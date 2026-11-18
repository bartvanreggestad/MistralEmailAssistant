// Mistral Email Assistant - Command Functions
// Deze functies worden aangeroepen vanuit de Office ribbon

(function () {
    "use strict";

    // Office initialiseren
    Office.initialize = function (reason) {
        console.log("Office initialized: " + reason);
    };

    // Functie voor het genereren van een antwoord
    function generateReply(event) {
        console.log("Generating reply...");
        
        const item = Office.context.mailbox.item;
        
        // Haal de basisgegevens op
        Promise.all([
            getItemPropertyAsync(item, "subject"),
            getItemPropertyAsync(item, "from"),
            getItemPropertyAsync(item, "to"),
            getItemPropertyAsync(item, "cc"),
            item.body.getAsync(Office.CoercionType.Text)
        ]).then(values => {
            const subject = values[0];
            const from = values[1];
            const to = values[2];
            const cc = values[3];
            const body = values[4].value;
            
            const emailDetails = {
                id: item.itemId,
                subject: subject || "Geen onderwerp",
                from: from ? (from.displayName || from.emailAddress.address) : "",
                to: to ? to.map(r => r.displayName || r.emailAddress.address) : [],
                cc: cc ? cc.map(r => r.displayName || r.emailAddress.address) : [],
                body: body || "",
                receivedDateTime: item.receivedDateTime
            };
            
            // Open een dialoogvenster om het antwoord te genereren
            Office.context.ui.displayDialogAsync(
                window.location.origin + "/src/taskpane/reply-dialog.html",
                { height: 60, width: 40, displayInIframe: true },
                function (asyncResult) {
                    if (asyncResult.status === "succeeded") {
                        const dialog = asyncResult.value;
                        
                        // Stuur de e-mail details naar het dialoogvenster
                        dialog.messageChild.postMessage({
                            action: "init",
                            email: emailDetails
                        }, "*");
                        
                        // Luister naar berichten van het dialoogvenster
                        const messageListener = function (messageEvent) {
                            if (messageEvent.data && messageEvent.data.action === "generate") {
                                const generatedReply = messageEvent.data.content;
                                
                                // Voeg het gegenereerde antwoord in de reply
                                insertReplyIntoEmail(generatedReply);
                                dialog.close();
                                event.completed();
                                
                                // Verwijder de listener
                                window.removeEventListener("message", messageListener);
                            }
                        };
                        
                        window.addEventListener("message", messageListener);
                        
                        // Timeout voor de listener
                        setTimeout(() => {
                            window.removeEventListener("message", messageListener);
                        }, 300000); // 5 minuten
                    } else {
                        console.error("Dialog failed:", asyncResult.error);
                        showError("Fout bij openen dialoogvenster: " + asyncResult.error.message);
                        event.completed();
                    }
                }
            );
        }).catch(error => {
            console.error("Error getting email details:", error);
            showError("Fout bij ophalen e-mail details: " + error.message);
            event.completed();
        });
    }

    // Functie voor het genereren van een nieuwe e-mail
    function generateEmail(event) {
        console.log("Generating new email...");
        
        // Open een dialoogvenster voor het genereren van een nieuwe e-mail
        Office.context.ui.displayDialogAsync(
            window.location.origin + "/src/taskpane/compose-dialog.html",
            { height: 60, width: 40, displayInIframe: true },
            function (asyncResult) {
                if (asyncResult.status === "succeeded") {
                    const dialog = asyncResult.value;
                    
                    // Luister naar berichten van het dialoogvenster
                    const messageListener = function (messageEvent) {
                        if (messageEvent.data && messageEvent.data.action === "generate") {
                            const generatedContent = messageEvent.data.content;
                            const options = messageEvent.data.options;
                            
                            // Maak een nieuwe e-mail aan met de gegenereerde content
                            createNewEmail(generatedContent, options);
                            dialog.close();
                            event.completed();
                            
                            // Verwijder de listener
                            window.removeEventListener("message", messageListener);
                        }
                    };
                    
                    window.addEventListener("message", messageListener);
                    
                    // Timeout voor de listener
                    setTimeout(() => {
                        window.removeEventListener("message", messageListener);
                    }, 300000); // 5 minuten
                } else {
                    console.error("Dialog failed:", asyncResult.error);
                    showError("Fout bij openen dialoogvenster: " + asyncResult.error.message);
                    event.completed();
                }
            }
        );
    }

    // Functie voor het samenvatten van een e-mail
    function summarizeEmail(event) {
        console.log("Summarizing email...");
        
        const item = Office.context.mailbox.item;
        
        // Haal de body op
        item.body.getAsync(Office.CoercionType.Text, function (result) {
            if (result.status === "succeeded") {
                const emailContent = result.value;
                const subject = item.subject || "Geen onderwerp";
                
                const prompt = {
                    model: "mistral-small",
                    messages: [
                        {
                            role: "system",
                            content: `Jij bent een professionele assistent. Geef een korte, heldere samenvatting van de volgende e-mail. Onderwerp: ${subject}. E-mail inhoud: ${emailContent}. Geef een samenvatting in maximaal 5 regels.`
                        }
                    ],
                    temperature: 0.3,
                    max_tokens: 200
                };
                
                callMistralAPI(prompt).then(response => {
                    // Toon de samenvatting in een notificatie
                    Office.context.mailbox.item.notificationMessages.replaceAsync(
                        "summary",
                        { type: "informationalMessage", message: "Samenvatting: " + response }
                    );
                    event.completed();
                }).catch(error => {
                    console.error("Error generating summary:", error);
                    showError("Fout bij genereren samenvatting: " + error.message);
                    event.completed();
                });
            } else {
                showError("Fout bij ophalen e-mail content: " + result.error.message);
                event.completed();
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

    // Roep de Mistral API aan
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
                        errorMessage = "Ongeldige API key. Controleer je Mistral API key in de instellingen.";
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

    // Voeg het gegenereerde antwoord in de reply
    function insertReplyIntoEmail(content) {
        Office.context.mailbox.item.replyAllAsync({
            asyncContext: { content: content }
        }, function (result) {
            if (result.status === "succeeded") {
                const replyItem = result.value;
                
                // Voeg de gegenereerde content toe
                replyItem.body.setAsync(content, { coercionType: Office.CoercionType.Html }, function (setResult) {
                    if (setResult.status === "succeeded") {
                        console.log("Reply inserted successfully");
                    } else {
                        console.error("Error setting reply body:", setResult.error);
                        showError("Fout bij invullen antwoord: " + setResult.error.message);
                    }
                });
            } else {
                console.error("Error creating reply:", result.error);
                showError("Fout bij maken reply: " + result.error.message);
            }
        });
    }

    // Maak een nieuwe e-mail aan
    function createNewEmail(content, options) {
        const toRecipients = options.to.map(email => ({ 
            emailAddress: { address: email } 
        }));
        
        const ccRecipients = options.cc.map(email => ({ 
            emailAddress: { address: email } 
        }));
        
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: toRecipients,
            ccRecipients: ccRecipients,
            subject: options.subject
        }, function (result) {
            if (result.status === "succeeded") {
                const newItem = result.value;
                
                // Voeg de content toe
                newItem.body.setAsync(content, { coercionType: Office.CoercionType.Html }, function (setResult) {
                    if (setResult.status === "succeeded") {
                        console.log("New email created successfully");
                    } else {
                        console.error("Error setting new email body:", setResult.error);
                        showError("Fout bij invullen e-mail: " + setResult.error.message);
                    }
                });
            } else {
                console.error("Error displaying new message form:", result.error);
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
                        reject(new Error("Geen Mistral API key gevonden. Configureer eerst je API key in de instellingen."));
                    }
                }
            });
        });
    }

    // Toon foutmelding
    function showError(message) {
        console.error(message);
        Office.context.mailbox.item.notificationMessages.replaceAsync(
            "error",
            { type: "errorMessage", message: message }
        );
    }

    // Registreer de functies voor Office
    Office.actions.associate("generateReply", generateReply);
    Office.actions.associate("generateEmail", generateEmail);
    Office.actions.associate("summarizeEmail", summarizeEmail);

})();
