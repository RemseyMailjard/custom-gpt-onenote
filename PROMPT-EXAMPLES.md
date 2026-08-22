# Prompt examples — agentic OneNote workflows

Real example prompts for the Custom GPT built on this Action, written the way
you'd actually type them (Dutch, since that's the daily-use language). Each
one shows the **tool-call chain** the GPT should run to satisfy it — useful
both as a manual test script after deploying a change, and as a reference for
what "good" tool use looks like when tuning the GPT's Instructions.

None of these need code changes — they're all compositions of operations
already in [openapi.yaml](openapi.yaml). See [README.md](README.md#onenote-routing)
for the underlying routing rules these flows follow.

## Find & understand

### 1. Search across notebooks with instant relevance judgment

> "Wat heb ik allemaal genoteerd over Microsoft Purview?"

```
listAllSections  (or listNotebooks → listSectionsInNotebook, if scoping to one notebook)
  → listPages?search=Purview&includePreview=true
  → [judge relevance from the preview snippets — no extra call per candidate]
  → getPageContent?format=markdown  (only for the 2-3 pages that actually look relevant)
  → synthesize an answer, citing notebook/section/page per README's "grounding" convention
```

Without `includePreview=true` this would cost one extra `getPageContent` round
trip per candidate page just to *check* relevance. With it, that check is
already in the search results.

### 2. "What changed recently" — sorting and filtering, not scrolling

> "Wat heb ik deze week aangepast in mijn cursusnotities?"

```
listNotebooks  → resolve the "Cursussen" notebook id
listPagesInSection?$filter=lastModifiedDateTime ge 2026-08-17T00:00:00Z&$orderby=lastModifiedDateTime desc&includePreview=true
  → present as a dated list with snippets
```

`$filter` and `$orderby` are plain OData pass-through (documented on
`listPages`/`listPagesInSection`/`listNotebooks`) — the proxy does no work
here, Graph does the filtering server-side.

### 3. Cross-reference multiple notes into one grounded answer

> "Wat heb ik besloten over de trainingsomgeving voor Klant X?"

```
listPages?search=trainingsomgeving&includePreview=true
listPages?search=Klant X&includePreview=true
  → [pick overlapping/relevant candidates from both result sets]
  → getPageContent?format=markdown  (each candidate)
  → answer, explicitly separating what's stated in the notes from what's inferred
    (per README's hallucination-protection convention)
```

## Setting up a new course or client

### 4. Clone a course template for a new client

> "Maak een nieuwe sectie 'Azure Fundamentals — Klant Y' aan in mijn
> Cursussen-notitieboek, gebaseerd op mijn 'Azure Fundamentals Template'
> sectie."

```
listSectionsInNotebook  → resolve the template section's id
listNotebooks           → resolve the destination notebook's id
copySectionToNotebook(sectionId=<template>, id=<destinationNotebookId>, renameAs="Azure Fundamentals — Klant Y")
  → 202 accepted; confirm the copy was requested, note it finishes asynchronously
```

All pages in the template section come along in one call — no per-page
`createPage` loop.

### 5. Start a whole new client engagement from a consultancy template

> "Ik start een nieuw traject bij Klant Z. Gebruik mijn 'Consultancy
> Sjabloon' notitieboek als basis."

```
listNotebooks  → resolve the "Consultancy Sjabloon" notebook id
copyNotebook(notebookId=<template>, renameAs="Klant Z")
  → 202 accepted
```

Every section and page in the template notebook is duplicated — this is the
one-call alternative to rebuilding a client's starting structure by hand.

### 6. Structure a new client with section groups

> "Maak een sectiegroep 'Klant ABC' aan in mijn Consultancy-notitieboek, met
> daarin een sectie 'Kickoff'."

```
listNotebooks           → resolve notebook id
createSectionGroup(notebookId, displayName="Klant ABC")
createSectionInGroup(sectionGroupId=<new group's id>, displayName="Kickoff")
```

## Capturing content

### 7. Attach a screenshot to meeting notes

> "Voeg deze screenshot toe aan mijn 'Kickoff Klant X'-pagina met als
> bijschrift 'Architectuuroverzicht besproken tijdens de kickoff'."

```
listPagesInSection  → resolve the "Kickoff Klant X" page id
addImageToPage(pageId, data=<base64 of the attached image>, contentType="image/png", alt="Architectuuroverzicht besproken tijdens de kickoff")
  → 204, confirm briefly
```

The image has to actually be attached to the conversation — never fabricate
image bytes.

### 8. Attach a handout PDF

> "Voeg deze PDF toe aan de pagina 'Week 3 — Networking' als handout."

```
listPagesInSection  → resolve the page id
addAttachmentToPage(pageId, data=<base64 of the attached PDF>, contentType="application/pdf", fileName="week3-networking-handout.pdf")
  → 204, confirm briefly
```

### 9. Append structured content to an existing page

> "Voeg onderaan mijn 'Klant X Status'-pagina een lijst met deze drie
> openstaande acties toe: contract tekenen, kickoff plannen, toegang
> aanvragen."

```
listPagesInSection  → resolve the page id
getPageContent(pageId)  [default format=html — need it for a precise data-id target, or fall back to target:"body"]
updatePageContent(pageId, commands=[{target:"body", action:"append", content:"<h2>Openstaande acties</h2><ul><li>...</li></ul>"}])
  → 204, confirm briefly
```

## Organizing and cleaning up

### 10. Move a stray page to the right place

> "Verplaats de pagina 'Losse notitie over DNS' van mijn Inbox-sectie naar de
> sectie 'Networking Basics' in mijn Cursussen-notitieboek."

```
listPagesInSection(Inbox section)     → resolve the page id
listSectionsInNotebook(Cursussen)     → resolve the destination section id
movePage(pageId, destinationSectionId=<Networking Basics id>)
  → 200 {success:true, newPageId, deletedOriginal:true}
  → confirm the move; if 202 comes back instead, say the copy is still
    running and the original wasn't touched yet — don't claim it's done
```

## Full research-to-storage loop

### 11. Research, synthesize, and file the result — end to end

> "Zoek al mijn notities over 'Zero Trust', vat ze samen, en sla die
> samenvatting op als nieuwe pagina 'Zero Trust — Samenvatting' in mijn
> Security-sectie."

```
listPages?search=Zero Trust&includePreview=true
  → [pick the relevant pages from the previews]
getPageContent?format=markdown  (each relevant page)
  → [synthesize a summary from the retrieved content — never from memory]
listSectionsInNotebook  → resolve the "Security" section id
createPage(sectionId, content="<!DOCTYPE html>...<title>Zero Trust — Samenvatting</title>...")
  → 201, confirm with the page title
```

This is the shape most daily requests should collapse into: **retrieve with
Graph, reason in the model, write back with Graph** — never skip the retrieve
step and answer from memory, and never claim a write succeeded before the
Action actually returns success.
