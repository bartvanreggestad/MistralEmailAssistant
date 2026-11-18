# Mistral Email Assistant - Outlook Add-in

Een krachtige Outlook-add-in voor het genereren, beantwoorden en analyseren van e-mails met Mistral AI.

## Functies

- **E-mails genereren**: Maak nieuwe e-mails aan met AI-assistentie
- **Antwoorden genereren**: Genereer intelligente antwoorden op inkomende e-mails
- **E-mails samenvatten**: Kreeg een korte samenvatting van lange e-mails
- **E-mails vertalen**: Vertaal e-mails naar andere talen
- **Aanpasbare instellingen**: Kies uit verschillende Mistral modellen en pas de toon en stijl aan

## Vereisten

- Outlook 2016 of nieuwer (Windows, Mac, of Outlook op het web)
- Een geldige [Mistral AI API key](https://mistral.ai/)
- Internetverbinding

## Installatie

### Methode 1: Handmatige installatie (voor ontwikkeling)

1. **Kloon of download dit project**
   ```bash
   git clone <repository-url>
   cd OutlookPlugin
   ```

2. **Maak een netwerk gedeelde map** (voor lokale ontwikkeling)
   - Plaats de map op een locatie die bereikbaar is via HTTP/HTTPS
   - Bijvoorbeeld: `C:\inetpub\wwwroot\MistralEmailAssistant`

3. **Pas het manifest.xml aan**
   - Open `manifest.xml` in een teksteditor
   - Update de `Id` met een unieke GUID (gebruik `uuidgen` of een online generator)
   - Pas de `SupportUrl` aan naar je eigen website
   - Update de `AppDomains` met je domein

4. **Voeg iconen toe**
   - Vervang de placeholder iconen in de `assets` map door echte PNG-bestanden
   - Aanbevolen afmetingen: 16x16, 32x32, 48x48, 64x64, 80x80, 128x128 pixels

5. **Implementeer de add-in in Outlook**
   
   **Voor Outlook op het web:**
   - Ga naar [Outlook Developer Dashboard](https://outlook.office.com/mail/inbox)
   - Klik op het tandwiel icoon (Instellingen) > Alle instellingen bekijken
   - Ga naar Mail > Aangepaste add-ins
   - Klik op "Add custom add-in" en upload het manifest.xml bestand
   
   **Voor Outlook Desktop (Windows/Mac):**
   - Open Outlook
   - Ga naar Bestand > Beheren > Add-ins
   - Klik op "Mijn add-ins" > "Add a custom add-in"
   - Browse naar het manifest.xml bestand en selecteer het
   
   **Voor side-loading (ontwikkeling):**
   - Gebruik de Office Add-in Developer Tools
   - Voer uit: `npm install -g office-addin-cli`
   - Navigeer naar de projectmap en voer uit: `office-addin start`

### Methode 2: AppSource publicatie (voor productie)

1. **Registreer je add-in in Partner Center**
   - Ga naar [Microsoft Partner Center](https://partner.microsoft.com/)
   - Maak een nieuwe Office Add-in registratie aan
   
2. **Upload je add-in pakket**
   - Zip de volledige projectmap
   - Upload naar Partner Center
   
3. **Certificeer en publiceer**
   - Voltooi de certificeringsproces
   - Publiceer naar AppSource

## Configuratie

### Mistral API Key instellen

1. **Krijg een API key**
   - Ga naar [Mistral AI](https://mistral.ai/) en registreer je
   - Maak een API key aan in je account dashboard

2. **Configureer in Outlook**
   - Open een e-mail in Outlook
   - Klik op de "Open Mistral" knop in de ribbon
   - Voer je API key in en klik op "Opslaan"

## Gebruik

### Een nieuwe e-mail genereren

1. Klik op "Nieuw bericht" in Outlook
2. Klik op de "E-mail genereren" knop in de Mistral groep
3. Voer de benodigde informatie in:
   - Onderwerp
   - Ontvangers (Aan en CC)
   - Inhoud of context
   - Toon (formeel, informeel, vriendelijk, etc.)
   - Stijl (beknopt, uitgebreid, technisch, etc.)
4. Klik op "Genereer E-mail"
5. De gegenereerde e-mail wordt automatisch ingevuld

### Een antwoord genereren

1. Open de e-mail waaraan je wilt antwoorden
2. Klik op de "Antwoord genereren" knop in de Mistral groep
3. (Optioneel) Pas de toon, stijl en andere opties aan
4. Klik op "Genereer Antwoord"
5. Het gegenereerde antwoord wordt automatisch in een reply ingevuld

### Een e-mail samenvatten

1. Open de e-mail die je wilt samenvatten
2. Klik op de "Samenvatten" knop in de Mistral groep
3. Een samenvatting verschijnt als een notificatie

## Bestanden structuur

```
OutlookPlugin/
├── manifest.xml                    # Add-in manifest
├── README.md                        # Dit bestand
├── src/
│   ├── commands/
│   │   ├── commands.html            # HTML voor command handling
│   │   └── commands.js              # JavaScript voor command functies
│   ├── taskpane/
│   │   ├── read.html                # HTML voor leesmodus taskpane
│   │   ├── read.js                  # JavaScript voor leesmodus
│   │   ├── compose.html             # HTML voor nieuw bericht taskpane
│   │   ├── compose.js               # JavaScript voor nieuw bericht
│   │   ├── reply-dialog.html        # HTML voor antwoord generatie dialog
│   │   ├── reply-dialog.js          # JavaScript voor antwoord dialog
│   │   ├── compose-dialog.html      # HTML voor e-mail generatie dialog
│   │   └── compose-dialog.js        # JavaScript voor e-mail dialog
│   └── styles/
│       └── style.css                # CSS stijlen
└── assets/
    ├── icon-16.png                  # Icoon 16x16
    ├── icon-32.png                  # Icoon 32x32
    ├── icon-48.png                  # Icoon 48x48
    ├── icon-64.png                  # Icoon 64x64
    ├── icon-80.png                  # Icoon 80x80
    └── icon-128.png                 # Icoon 128x128
```

## Aangepaste ontwikkelingen

### Eigene modellen toevoegen

Pas het `model-select` dropdown menu aan in de HTML-bestanden en voeg de nieuwe modelnaam toe aan de API calls.

### Taalondersteuning uitbreiden

Voeg vertaalopties toe door de `buildTranslationPrompt` functie uit te breiden met meer talen.

### Stijl templates maken

Maak voorgedefinieerde templates voor verschillende soorten e-mails (bijvoorbeeld: klantenservice, verkoop, technische ondersteuning).

## Veiligheid

- **API Key opslag**: De API key wordt lokaal opgeslagen in Office settings en localStorage
- **HTTPS vereist**: Zorg ervoor dat je add-in via HTTPS wordt geladen voor productie
- **Input validatie**: Alle gebruiker input wordt gevalideerd voor het aanroepen van de API

## Foutoplossing

### Algemene problemen

**"Geen Mistral API key gevonden"**
- Zorg ervoor dat je een API key hebt ingesteld in de instellingen
- Controleer of de key correct is opgeslagen

**"Ongeldige API key"**
- Controleer of je API key nog geldig is
- Maak een nieuwe key aan als de huidige is verlopen

**"Te veel verzoeken"**
- Wacht een paar minuten en probeer opnieuw
- Overweeg om een betaald abonnement te nemen voor meer verzoeken

### Ontwikkelingsproblemen

**"Add-in wordt niet geladen"**
- Controleer of het manifest.xml bestand geldig XML is
- Zorg ervoor dat alle bestandsriften correct zijn
- Controleer de console log in de browser (F12)

**"Office API fouten"**
- Zorg ervoor dat je de juiste Office.js versie gebruikt
- Controleer of je add-in de benodigde permissies heeft

## Ondersteuning

Voor vragen en ondersteuning:
- E-mail: support@reggestad.nl
- Website: https://www.reggestad.nl

## Licentie

Deze software is eigendom van Reggestad ICT Specialisten. Voor licentie-informatie, neem contact op met ons.

## Versie geschiedenis

- **1.0.0** (2024): Eerste versie met basis functionaliteit
  - E-mail generatie
  - Antwoord generatie
  - Samenvatting functionaliteit
  - Mistral API integratie
