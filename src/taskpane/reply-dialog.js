// Mistral Email Assistant - Reply Dialog
(function () {
    "use strict";

    let emailDetails = null;

    // Office initialiseren
    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("Office initialized in reply dialog: " + reason);
            setupEventHandlers();
            
            // Luister naar berichten van de parent
            window.addEventListener("message", function (event) {
                if (event.data && event.data.action === "init") {
                    emailDetails = event.data.email;
                    displayEmailDetails();
                }
            });
        });
    };

    // Stel event handlers in
    function setupEventHandlers() {
        $("#btn-cancel").click(function () {
            Office.context.ui.closeContainer();
        });

        $("#btn-generate").click(function () {
            generateReply();
        });
    }

    // Toon de e-mail details
    function displayEmailDetails() {
        if (!emailDetails) return;

        $("#dialog-subject").text(emailDetails.subject || "Geen onderwerp");
        $("#dialog-from").text(emailDetails.from || "Onbekend");
        $("#dialog-body").text(emailDetails.body || "Geen inhoud");
    }

    // Genereer een antwoord
    function generateReply() {
        if (!emailDetails) {
            showError("Geen e-mail details geladen");
            return;
        }

        const tone = $("#reply-tone").val();
        const style = $("#reply-style").val();
        const model = $("#reply-model").val();
        const additionalContext = $("#additional-context").val();

        showLoading(true);

        const prompt = buildReplyPrompt(emailDetails, { tone, style, model, additionalContext });

        callMistralAPI(prompt).then(response => {
            // Stuur het gegenereerde antwoord terug naar de parent
            window.parent.postMessage({
                action: "generate",
                content: response,
                options: { tone, style, model, additionalContext }
            }, "*");
            
            showLoading(false);
            Office.context.ui.closeContainer();
        }).catch(error => {
            showError("Fout bij genereren antwoord: " + error.message);
            showLoading(false);
        });
    }

    // Bouw de prompt voor antwoord generatie
    function buildReplyPrompt(emailDetails, options) {
        const systemPrompt = `Jij bent een professionele Nederlandse e-mail assistant.
Schrijf een duidelijk, beleefd en professioneel antwoord op de volgende e-mail.

Richtlijnen:
- Gebruik een ${options.tone || "professionele"} toon
- Houd rekening met de ${options.style || "standaard"} stijl
- Antwoord in het Nederlands
- Beantwoord alle belangrijke punten uit de originele e-mail
- Wees beknopt maar compleet

Originele e-mail:
Onderwerp: ${emailDetails.subject}
Van: ${emailDetails.from}
Aan: ${emailDetails.to.join(", ")}

Inhoud:
${emailDetails.body}

${options.additionalContext ? "Aanvullende context: " + options.additionalContext : ""}

Schrijf een antwoord:`;

        return {
            model: options.model,
            messages: [
                { role: "system", content: systemPrompt }
            ],
            temperature: 0.7,
            max_tokens: 1000
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
                        reject(new Error("Geen Mistral API key gevonden. Configureer eerst je API key."));
                    }
                }
            });
        });
    }

    // Toon laadindicator
    function showLoading(show) {
        $("#loading-overlay").toggle(show);
    }

    // Toon foutmelding
    function showError(message) {
        console.error(message);
        alert(message);
    }

})();
