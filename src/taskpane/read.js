// Mistral Email Assistant - Read Taskpane
(function () {
    "use strict";

    // Office initialiseren
    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("Office initialized in read taskpane: " + reason);
            loadEmailDetails();
            setupEventHandlers();
            loadSettings();
        });
    };

    // Laad de e-mail details
    function loadEmailDetails() {
        Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
            if (result.status === "succeeded") {
                const item = Office.context.mailbox.item;
                
                // Haal basisgegevens op
                $("#email-subject").text(item.subject || "Geen onderwerp");
                
                // Haal afzender op
                if (item.from) {
                    $("#email-from").text(item.from.displayName || item.from.emailAddress.address);
                }
                
                // Haal ontvangers op
                if (item.to && item.to.length > 0) {
                    const toList = item.to.map(r => r.displayName || r.emailAddress.address).join(", ");
                    $("#email-to").text(toList);
                }
                
                // Haal de body op
                item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {
                    if (bodyResult.status === "succeeded") {
                        $("#email-body").html(bodyResult.value || "<p>Geen inhoud</p>");
                    } else {
                        $("#email-body").html("<p>Fout bij laden e-mail body</p>");
                    }
                });
            } else {
                console.error("Error loading email details:", result.error);
            }
        });
    }

    // Stel event handlers in
    function setupEventHandlers() {
        // Samenvatten knop
        $("#btn-summarize").click(function () {
            generateSummary();
        });

        // Vertalen knop
        $("#btn-translate").click(function () {
            translateEmail();
        });

        // Antwoord genereren knop
        $("#btn-generate-reply").click(function () {
            generateReply();
        });

        // API key opslaan
        $("#btn-save-key").click(function () {
            saveApiKey();
        });
    }

    // Laad instellingen
    function loadSettings() {
        // Laad API key status
        getMistralApiKey().then(apiKey => {
            if (apiKey) {
                $("#api-key-status").text("API key is ingesteld").css("color", "green");
            } else {
                $("#api-key-status").text("Geen API key ingesteld").css("color", "red");
            }
        }).catch(error => {
            $("#api-key-status").text("Fout bij laden API key").css("color", "red");
        });

        // Laad geselecteerd model
        const savedModel = localStorage.getItem("mistralModel") || "mistral-small";
        $("#model-select").val(savedModel);
    }

    // Genereer een samenvatting
    function generateSummary() {
        showLoading(true);
        
        const item = Office.context.mailbox.item;
        
        // Haal de body op
        item.body.getAsync(Office.CoercionType.Text, function (result) {
            if (result.status === "succeeded") {
                const emailContent = result.value;
                const prompt = buildSummaryPrompt(emailContent, item.subject);
                
                callMistralAPI(prompt).then(response => {
                    showResponse("Samenvatting:", response);
                    showLoading(false);
                }).catch(error => {
                    showError("Fout bij genereren samenvatting: " + error.message);
                    showLoading(false);
                });
            } else {
                showError("Fout bij ophalen e-mail content");
                showLoading(false);
            }
        });
    }

    // Vertaal de e-mail
    function translateEmail() {
        showLoading(true);
        
        const item = Office.context.mailbox.item;
        const targetLanguage = "Nederlands"; // TODO: Maak dit configureerbaar
        
        // Haal de body op
        item.body.getAsync(Office.CoercionType.Text, function (result) {
            if (result.status === "succeeded") {
                const emailContent = result.value;
                const prompt = buildTranslationPrompt(emailContent, targetLanguage);
                
                callMistralAPI(prompt).then(response => {
                    showResponse(`Vertaald naar ${targetLanguage}:`, response);
                    showLoading(false);
                }).catch(error => {
                    showError("Fout bij vertalen: " + error.message);
                    showLoading(false);
                });
            } else {
                showError("Fout bij ophalen e-mail content");
                showLoading(false);
            }
        });
    }

    // Genereer een antwoord
    function generateReply() {
        showLoading(true);
        
        const item = Office.context.mailbox.item;
        
        // Haal alle details op
        Promise.all([
            getItemPropertyAsync(item, "subject"),
            getItemPropertyAsync(item, "from"),
            getItemPropertyAsync(item, "to"),
            item.body.getAsync(Office.CoercionType.Text)
        ]).then(values => {
            const subject = values[0];
            const from = values[1];
            const to = values[2];
            const body = values[3].value;
            
            const emailDetails = {
                subject: subject,
                from: from ? (from.displayName || from.emailAddress.address) : "",
                to: to ? to.map(r => r.displayName || r.emailAddress.address) : [],
                body: body
            };
            
            const prompt = buildReplyPrompt(emailDetails);
            
            callMistralAPI(prompt).then(response => {
                showResponse("Voorgesteld antwoord:", response);
                showLoading(false);
                
                // Toon knop om het antwoord te gebruiken
                $("#ai-response").append(`
                    <div class="action-buttons" style="margin-top: 15px;">
                        <button id="btn-use-reply" class="btn btn-primary">Gebruik dit antwoord</button>
                    </div>
                `);
                
                $("#btn-use-reply").click(function () {
                    useGeneratedReply(response);
                });
            }).catch(error => {
                showError("Fout bij genereren antwoord: " + error.message);
                showLoading(false);
            });
        }).catch(error => {
            showError("Fout bij ophalen e-mail details: " + error.message);
            showLoading(false);
        });
    }

    // Gebruik het gegenereerde antwoord
    function useGeneratedReply(replyContent) {
        Office.context.mailbox.item.replyAllAsync({
            asyncContext: { content: replyContent }
        }, function (result) {
            if (result.status === "succeeded") {
                const replyItem = result.value;
                
                // Voeg de content toe
                replyItem.body.setAsync(replyContent, { coercionType: Office.CoercionType.Html }, function (setResult) {
                    if (setResult.status === "succeeded") {
                        showResponse("Antwoord ingevuld", "Het gegenereerde antwoord is succesvol ingevuld in de reply.");
                    } else {
                        showError("Fout bij invullen antwoord: " + setResult.error.message);
                    }
                });
            } else {
                showError("Fout bij maken reply: " + result.error.message);
            }
        });
    }

    // Haal een property van het item op
    function getItemPropertyAsync(item, propertyName) {
        return new Promise((resolve, reject) => {
            item.getAsync(propertyName, function (result) {
                if (result.status === "succeeded") {
                    resolve(result.value);
                } else {
                    reject(new Error(`Fout bij ophalen ${propertyName}: ${result.error.message}`));
                }
            });
        });
    }

    // Bouw prompt voor samenvatting
    function buildSummaryPrompt(emailContent, subject) {
        const systemPrompt = `Jij bent een professionele assistent. Geef een korte, heldere samenvatting van de volgende e-mail.

Onderwerp: ${subject || "Onbekend"}

E-mail inhoud:
${emailContent}

Geef een samenvatting in maximaal 5 regels:`;

        return {
            model: $("#model-select").val(),
            messages: [
                { role: "system", content: systemPrompt }
            ],
            temperature: 0.3,
            max_tokens: 200
        };
    }

    // Bouw prompt voor vertaling
    function buildTranslationPrompt(emailContent, targetLanguage) {
        const systemPrompt = `Jij bent een vertaler. Vertaal de volgende tekst naar ${targetLanguage}. 
Behoud de toon en stijl van de originele tekst. Vertaal alleen de inhoud, voeg geen uitleg toe.

Originele tekst:
${emailContent}

Vertaalde tekst:`;

        return {
            model: $("#model-select").val(),
            messages: [
                { role: "system", content: systemPrompt }
            ],
            temperature: 0.2,
            max_tokens: 1500
        };
    }

    // Bouw prompt voor antwoord
    function buildReplyPrompt(emailDetails) {
        const systemPrompt = `Jij bent een professionele e-mail assistant. 
Schrijf een duidelijk, beleefd en professioneel antwoord op de volgende e-mail.
Gebruik een formele toon. Houd rekening met de context en de relatie met de afzender.
Antwoord in het Nederlands.

Originele e-mail:
Onderwerp: ${emailDetails.subject}
Van: ${emailDetails.from}
Aan: ${emailDetails.to.join(", ")}

Inhoud:
${emailDetails.body}

Schrijf een antwoord:`;

        return {
            model: $("#model-select").val(),
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

    // Sla de API key op
    function saveApiKey() {
        const apiKey = $("#api-key").val();
        
        if (!apiKey) {
            showError("Voer een API key in");
            return;
        }

        // Sla op in localStorage
        localStorage.setItem("mistralApiKey", apiKey);
        
        // Sla ook op in Office settings
        Office.context.document.settings.set("mistralApiKey", apiKey);
        Office.context.document.settings.saveAsync(function (result) {
            if (result.status === "succeeded") {
                $("#api-key-status").text("API key opgeslagen").css("color", "green");
                $("#api-key").val("");
            } else {
                showError("Fout bij opslaan API key in Office settings");
            }
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
        if (show) {
            $("#ai-response-section").show();
            $("#loading-indicator").show();
            $("#ai-response").hide().empty();
        } else {
            $("#loading-indicator").hide();
            $("#ai-response").show();
        }
    }

    // Toon respons
    function showResponse(title, content) {
        $("#ai-response-section").show();
        $("#ai-response").append(`
            <h3>${title}</h3>
            <div class="response-text">${content}</div>
        `);
    }

    // Toon foutmelding
    function showError(message) {
        console.error(message);
        $("#ai-response-section").show();
        $("#ai-response").append(`
            <div class="error-message">
                <strong>Fout:</strong> ${message}
            </div>
        `);
    }

})();
