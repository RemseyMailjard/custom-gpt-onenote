# OneNote Prompt-voorbeelden

Voorbeeldprompts voor de OneNote-acties van de Microsoft Graph Custom GPT. De GPT vertaalt deze
natuurlijke taal automatisch naar de juiste Graph-operaties.

## Overzicht van de acties

| Actie | operationId | Doel |
| --- | --- | --- |
| Notitieblokken ophalen | `listOneNoteNotebooks` | Alle notitieblokken tonen |
| Notitieblok aanmaken | `createOneNoteNotebook` | Nieuw notitieblok maken |
| Secties ophalen | `listOneNoteNotebookSections` | Secties van een notitieblok tonen |
| Sectie aanmaken | `createOneNoteSection` | Nieuwe sectie maken |
| Pagina's ophalen | `listOneNoteSectionPages` | Pagina's van een sectie tonen |
| Pagina aanmaken | `createOneNotePage` | Nieuwe HTML-pagina maken |
| Pagina-inhoud ophalen | `getOneNotePageContent` | HTML van een pagina lezen (met `includeIDs`) |
| Pagina bijwerken | `updateOneNotePageContent` | Elementen op een pagina wijzigen |
| Pagina verwijderen | `deleteOneNotePage` | Pagina verwijderen |

---

## 1. Pagina aanmaken

### Voorbeeldprompt — pagina in "remsey" notitieblok, sectie "skills4it"

> Maak in mijn **remsey** notitieblok, sectie **skills4it**, een nieuwe OneNote-pagina aan met de titel
> **"Skills4IT – Weekoverzicht"**. Zet bovenaan een kop met de titel, daaronder een korte samenvatting
> en een sectie met actiepunten als bulletlijst. Voeg de volgende actiepunten toe:
> - Cursusmateriaal Azure Fundamentals afronden
> - Demo-omgeving klaarzetten voor de workshop
> - Feedbackformulier versturen naar deelnemers
>
> Zoek eerst het juiste sectie-ID op via de secties van het remsey-notitieblok voordat je de pagina aanmaakt.

De GPT voert dan achter de schermen uit:
1. `listOneNoteNotebooks` → vindt het notitieblok **remsey**.
2. `listOneNoteNotebookSections` → vindt de sectie **skills4it** en haalt het `section-id` op.
3. `createOneNotePage` → maakt de pagina aan met dit HTML-document:

```html
<!DOCTYPE html>
<html>
  <head>
    <title>Skills4IT – Weekoverzicht</title>
  </head>
  <body>
    <h1>Skills4IT – Weekoverzicht</h1>
    <div data-id="samenvatting">
      <p>Overzicht van de belangrijkste taken en voortgang van deze week.</p>
    </div>
    <div data-id="acties">
      <h2>Actiepunten</h2>
      <ul>
        <li>Cursusmateriaal Azure Fundamentals afronden</li>
        <li>Demo-omgeving klaarzetten voor de workshop</li>
        <li>Feedbackformulier versturen naar deelnemers</li>
      </ul>
    </div>
  </body>
</html>
```

### Kortere variant

> Maak een OneNote-pagina "Klantgesprek Contoso" aan in mijn remsey-notitieblok, sectie skills4it,
> met een kop, een samenvatting en een lijst met vervolgacties.

---

## 2. Pagina bijwerken

> Open de pagina "Skills4IT – Weekoverzicht" in mijn remsey-notitieblok en voeg onderaan een nieuw
> actiepunt toe: **"Certificaten uitreiken aan geslaagde deelnemers"**.

De GPT gebruikt hier:
1. `listOneNoteSectionPages` → zoekt de pagina en het `page-id`.
2. `getOneNotePageContent` met `includeIDs=true` → haalt de actuele element-ID's op.
3. `updateOneNotePageContent` → stuurt een wijzigingsopdracht, bijvoorbeeld:

```json
{
  "commands": [
    {
      "target": "#acties",
      "action": "append",
      "content": "<li>Certificaten uitreiken aan geslaagde deelnemers</li>"
    }
  ]
}
```

### Andere update-voorbeelden

> Vervang de samenvatting op de pagina "Skills4IT – Weekoverzicht" door: "Alle workshops zijn afgerond
> en de evaluaties zijn verwerkt."

> Voeg bovenaan de body van mijn pagina "Klantgesprek Contoso" een paragraaf toe met de datum van vandaag.

---

## 3. Pagina's en inhoud ophalen

> Toon alle pagina's in de sectie skills4it van mijn remsey-notitieblok.

> Haal de volledige inhoud op van de pagina "Skills4IT – Weekoverzicht" zodat ik de tekst kan lezen.

> Haal de inhoud van pagina "Skills4IT – Weekoverzicht" op met de element-ID's, zodat we een gericht
> onderdeel kunnen bijwerken.

---

## 4. Secties en notitieblokken beheren

> Laat al mijn OneNote-notitieblokken zien.

> Maak in mijn remsey-notitieblok een nieuwe sectie aan met de naam "Projecten 2026".

> Toon alle secties in het remsey-notitieblok.

---

## 5. Pagina verwijderen

> Verwijder de pagina "Skills4IT – Weekoverzicht" uit mijn remsey-notitieblok, sectie skills4it.

---

## Tips voor betrouwbare resultaten

- **Noem het notitieblok en de sectie expliciet** ("remsey", "skills4it"), zodat de GPT het juiste
  `section-id` opzoekt voordat een pagina wordt aangemaakt.
- **Vraag bij gerichte wijzigingen eerst om de inhoud met ID's** (`includeIDs=true`); door Graph
  gegenereerde ID's kunnen na elke wijziging veranderen.
- **Gebruik `data-id`-attributen** (zoals `#acties`) in aangemaakte pagina's; die blijven stabiel en
  zijn makkelijker als `target` te gebruiken dan gegenereerde element-ID's.
- **Houd de HTML eenvoudig**: koppen, paragrafen, lijsten en `div`-elementen werken het meest voorspelbaar.
