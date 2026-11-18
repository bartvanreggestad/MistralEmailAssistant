// Mistral Email Assistant - Compose Dialog
(function () {
    "use strict";

    // Office initialiseren
    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("Office initialized in compose dialog: " + reason);
            setupEventHandlers();
            loadSettings();
        });
    };

    // Stel event handlers in
    function setupEventHandlers() {
        // Creativiteit slider
        $("#compose-dialog-creativity").on("input", function () {
            $("#compose-dialog-creativity-value").text($(this).val());
        });

        // Annuleer knop
        $("#btn-compose-cancel").click(function () {
            Office.context.ui.closeContainer();
        });

        // Genereer knop
        $("#btn-compose-generate").click(function () {
            generateEmail();
        });
    }

    // Laad instellingen
    function loadSettings() {
        const savedModel = localStorage.getItem("mistralModel") || "mistral-small";
        $("#compose-dialog-model").val(savedModel);
        
        const savedTone = localStorage.getItem("defaultTone") || "";
        $("#compose-dialog-tone").val(savedTone);
        
        const savedStyle = localStorage.getItem("defaultStyle") || "";
        $("#compose-dialog-style").val(savedStyle);
    }

    // Genereer een e-mail
    function generateEmail() {
        const subject = $("#compose-dialog-subject").val();
        const to = $("#compose-dialog-to").val().split(",").map(email => email.trim()).filter(email => email);
        const cc = $("#compose-dialog-cc").val().split(",").map(email => email.trim()).filter(email => email);
        const content = $("#compose-dialog-content").val();
        const tone = $("#compose-dialog-tone").val();
        const style = $("#compose-dialog-style").val();
        const model = $("#compose-dialog-model").val();
        const creativity = parseFloat($("#compose-dialog-creativity").val());

        if (!subject) {
            showError("Voer een onderwerp in");
            return;
        }

        if (!to || to.length === 0) {
            showError("Voer tenminste één ontvanger in");
            return;
        }

        if (!content) {
            showError("Voer inhoud of context in");
            return;
        }

        showLoading(true);

        const prompt = buildEmailPrompt({
            subject: subject,
            to: to,
            cc: cc,
            content: content,
            tone: tone,
            style: style,
            model: model,
            creativity: creativity
        });

        callMistralAPI(prompt).then(response => {
            // Stuur het gegenereerde e-mail terug naar de parent
            window.parent.postMessage({
                action: "generate",
                content: response,
                options: {
                    subject: subject,
                    to: to,
                    cc: cc,
                    tone: tone,
                    style: style,
                    model: model,
                    creativity: creativity
                }
            }, "*");
            
            showLoading(false);
            Office.context.ui.closeContainer();
        }).catch(error => {
            showError("Fout bij genereren e-mail: " + error.message);
            showLoading(false);
        });
    }

    // Bouw de prompt voor e-mail generatie
    function buildEmailPrompt(options) {
        const systemPrompt = `Jij bent een professionele Nederlandse e-mail assistant.
Schrijf een complete e-mail volgens de volgende specificaties:

Onderwerp: ${options.subject}
Aan: ${options.to.join(", ")}
${options.cc.length > 0 ? "CC: " + options.cc.join(", ") : ""}

Inhoud/Context: ${options.content}
${options.tone ? "Toon: " + options.tone : ""}
${options.style ? "Stijl: " + options.style : ""}

Schrijf de complete e-mail met:
- Een passende aanhef
- De inhoud die past bij de context
- Een passende afsluiting
- Handtekening (gebruik een generieke handtekening)

Let op:
- Gebruik een ${options.tone || "professionele"} toon
- Houd de e-mail ${options.style || "beknopt maar compleet"}
- Gebruik correct Nederlands
- Voeg geen technische uitleg toe over hoe je de e-mail hebt gegenereerd

E-mail:`;

        return {
            model: options.model,
            messages: [
                { role: "system", content: systemPrompt }
            ],
            temperature: options.creativity,
            max_tokens: 1500
        };
    }

    // Roep Mistral API aan
    function callMistralAPI(promptConfig) {
        return new Promise((resolve, reject) => {
            getMistralApiKey().then(apiKey => {
                const apiUrl = "https://api.mistral.ai/v1/chat/completions";
                
                const data = {
                    model: promptConfig.model,
                    messages: promptConfig.messages,
                    temperature: promptConfig.temperature,
                    max_tokens: promptConfig.max_tokens
                };

                $.ajax({
                    url: apiUrl,
                    headers: {
                        "Authorization": "Bearer " + apiKey,
                        "Content-Type": "application/json"
                    },
                    type: "POST",
                    data: JSON.stringify(data)
                }).done(function (response) {
                    if (response.choices && response.choices.length > 0) {
                        resolve(response.choices[0].message.content);
                    } else {
                        reject(new Error("Geen antwoord ontvangen van Mistral API"));
                    }
                }).fail(function (jqXHR, textStatus, errorThrown) {
                    let errorMessage = "Fout bij aanroepen Mistral API";
                    
                    if (jqXHR.status === 401) {
                        errorMessage = "Ongeldige API key. Controleer je Mistral API key.";
                    } else if (jqXHR.status === 429) {
                        errorMessage = "Te veel verzoeken. Wacht even en probeer opnieuw.";
                    } else if (jqXHR.responseJSON && jqXHR.responseJSON.error) {
                        errorMessage = jqXHR.responseJSON.error.message || errorMessage;
                    }
                    
                    reject(new Error(errorMessage));
                });
            }).catch(reject);
        });
    }

    // Haal de Mistral API key op
    function getMistralApiKey() {
        return new Promise((resolve, reject) => {
            // Probeer eerst uit Office settings
            Office.context.document.settings.getAsync(function (settingsResult) {
                if (settingsResult.status === "succeeded" && settingsResult.value.mistralApiKey) {
                    resolve(settingsResult.value.mistralApiKey);
                } else {
                    // Probeer uit localStorage
                    const storedKey = localStorage.getItem("mistralApiKey");
                    if (storedKey) {
                        resolve(storedKey);
                    } else {
                        reject(new Error("Geen Mistral API key gevonden. Configureer eerst je API key in de leesmodus."));
                    }
                }
            });
        });
    }

    // Toon laadindicator
    function showLoading(show) {
        $("#compose-loading-overlay").toggle(show);
    }

    // Toon foutmelding
    function showError(message) {
        console.error(message);
        alert(message);
    }

})();
