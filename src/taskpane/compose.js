// Mistral Email Assistant - Compose Taskpane
(function () {
    "use strict";

    // Office initialiseren
    Office.initialize = function (reason) {
        $(document).ready(function () {
            console.log("Office initialized in compose taskpane: " + reason);
            setupEventHandlers();
            loadSettings();
        });
    };

    // Stel event handlers in
    function setupEventHandlers() {
        // Creativiteit slider
        $("#compose-creativity").on("input", function () {
            $("#creativity-value").text($(this).val());
        });

        // Genereer e-mail knop
        $("#btn-generate-email").click(function () {
            generateEmail();
        });

        // Gebruik deze e-mail knop
        $("#btn-use-email").click(function () {
            useGeneratedEmail();
        });

        // Opnieuw genereren knop
        $("#btn-regenerate").click(function () {
            $("#generated-email-section").hide();
            generateEmail();
        });
    }

    // Laad instellingen
    function loadSettings() {
        const savedModel = localStorage.getItem("mistralModel") || "mistral-small";
        $("#compose-model").val(savedModel);
        
        const savedTone = localStorage.getItem("defaultTone") || "";
        $("#compose-tone").val(savedTone);
        
        const savedStyle = localStorage.getItem("defaultStyle") || "";
        $("#compose-style").val(savedStyle);
    }

    // Genereer een e-mail
    function generateEmail() {
        const subject = $("#compose-subject").val();
        const to = $("#compose-to").val().split(",").map(email => email.trim()).filter(email => email);
        const cc = $("#compose-cc").val().split(",").map(email => email.trim()).filter(email => email);
        const content = $("#compose-content").val();
        const tone = $("#compose-tone").val();
        const style = $("#compose-style").val();
        const model = $("#compose-model").val();
        const creativity = parseFloat($("#compose-creativity").val());

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

        showGenerating(true);

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
            displayGeneratedEmail(response, { subject, to, cc });
            showGenerating(false);
        }).catch(error => {
            showError("Fout bij genereren e-mail: " + error.message);
            showGenerating(false);
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
- Gebruik een professionele toon tenzij anders gespecificeerd
- Houd de e-mail beknopt maar compleet
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

    // Toon gegenereerde e-mail
    function displayGeneratedEmail(content, emailDetails) {
        $("#generated-email-section").show();
        
        // Format de e-mail content
        const formattedContent = formatEmailContent(content, emailDetails);
        
        $("#generated-email").html(`
            <div class="email-header">
                <p><strong>Onderwerp:</strong> ${emailDetails.subject}</p>
                <p><strong>Aan:</strong> ${emailDetails.to.join(", ")}</p>
                ${emailDetails.cc.length > 0 ? `<p><strong>CC:</strong> ${emailDetails.cc.join(", ")}</p>` : ""}
            </div>
            <div class="email-body">
                ${formattedContent}
            </div>
        `);
    }

    // Formatteer de e-mail content
    function formatEmailContent(content, emailDetails) {
        // Voeg een standaard handtekening toe als die niet in de content zit
        if (content.indexOf("Met vriendelijke groet") === -1 && 
            content.indexOf("Mvg") === -1 && 
            content.indexOf("Groet") === -1) {
            return content + "\n\n-- \nMet vriendelijke groet,\n[Jouw Naam]\n[Jouw Positie]\n[Jouw Bedrijf]\n[Telefoonnummer]\n[E-mailadres]";
        }
        return content;
    }

    // Gebruik de gegenereerde e-mail
    function useGeneratedEmail() {
        const subject = $("#compose-subject").val();
        const to = $("#compose-to").val().split(",").map(email => email.trim()).filter(email => email);
        const cc = $("#compose-cc").val().split(",").map(email => email.trim()).filter(email => email);
        const content = $("#generated-email .email-body").text();

        Office.context.mailbox.displayNewMessageForm({
            toRecipients: to.map(email => ({ emailAddress: { address: email } })),
            ccRecipients: cc.map(email => ({ emailAddress: { address: email } })),
            subject: subject
        }, function (result) {
            if (result.status === "succeeded") {
                const newItem = result.value;
                
                // Voeg de content toe
                newItem.body.setAsync(content, { coercionType: Office.CoercionType.Html }, function (setResult) {
                    if (setResult.status === "succeeded") {
                        console.log("New email created successfully");
                        // Sluit het paneel
                        Office.context.ui.closeContainer();
                    } else {
                        showError("Fout bij invullen e-mail: " + setResult.error.message);
                    }
                });
            } else {
                showError("Fout bij maken nieuwe e-mail: " + result.error.message);
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
                        reject(new Error("Geen Mistral API key gevonden. Configureer eerst je API key in de leesmodus."));
                    }
                }
            });
        });
    }

    // Toon generatie-indicator
    function showGenerating(show) {
        if (show) {
            $("#generated-email-section").show();
            $("#generating-indicator").show();
            $("#generated-email").hide().empty();
        } else {
            $("#generating-indicator").hide();
            $("#generated-email").show();
        }
    }

    // Toon foutmelding
    function showError(message) {
        console.error(message);
        alert(message); // TODO: Vervang door een betere UI notificatie
    }

})();
