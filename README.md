# OneNote Graph Proxy — Custom GPT Action

## What this is

**NoteBuddy** (working name) lets you talk to your Microsoft OneNote notebooks
from a ChatGPT Custom GPT: find notes, read pages, and create or update pages,
all through natural conversation instead of clicking through OneNote.

This repo is the backend piece that makes that possible — a small **Azure
Functions proxy** in front of Microsoft Graph's OneNote API. It is *not* an
MCP server; it's a REST API described by an OpenAPI spec ([openapi.yaml](openapi.yaml))
that a ChatGPT Custom GPT calls as an "Action". The GPT itself (via its
Instructions, see below) does the reasoning: deciding which notebook/section/
page to use, resolving names to IDs, and generating page content. The proxy's
only real job is a small technical fix — Microsoft Graph requires
`multipart/form-data` for creating/updating page content, but ChatGPT Actions
can only send plain JSON — plus passing every other request straight through
to Graph with the caller's Microsoft token.

```
You (in ChatGPT) → Custom GPT → Action (this proxy) → Microsoft Graph → OneNote
```

What it currently supports (via Graph, exposed as GPT Action operations):

- List notebooks, sections, section groups, pages (`listNotebooks`, `listAllSections`,
  `listSectionGroupsInNotebook`, `listPagesInSection`, ...), with optional
  `$filter` (e.g. by `lastModifiedDateTime`) and `includePreview=true` for a
  plain-text snippet per result
- Read a page's content as HTML or as clean Markdown (`getPage`, `getPageContent?format=markdown`)
- Create a notebook, section, section group or page (`createNotebook`,
  `createSection`, `createSectionGroup`, `createSectionInGroup`, `createPage`)
- Update an existing page's content (`updatePageContent`)
- Add an image or file attachment to a page (`addImageToPage`, `addAttachmentToPage`)
- Copy or move a page to another section (`copyPageToSection`, `movePage`)
- Copy a whole section — template and all its pages — into another notebook
  or section group (`copySectionToNotebook`, `copySectionToSectionGroup`)
- Delete a page (`deletePage`)

Authentication is delegated Microsoft sign-in (OAuth via Entra ID) — you only
ever see and edit your own OneNote content, scoped to `Notes.ReadWrite`.

**Not supported, by design:** renaming a notebook or section. Microsoft Graph's
OneNote API only supports List/Get/Create on notebooks and sections — there is
no update operation, so a "rename" would have to be faked as delete-and-recreate,
which is destructive and loses page history/links. Rename sections directly in
OneNote instead.

## Installation

Everything needed to run this yourself is below: local development, deploying
to Azure, registering the Entra ID app, and wiring it up as a Custom GPT
Action. Follow the sections in order — Azure deploy → Entra app registration →
Custom GPT setup.

## Why a proxy?

Microsoft Graph's one-step OneNote page creation endpoint
(`POST /me/onenote/sections/{id}/pages`) requires a `multipart/form-data` body
with a `Presentation` part. ChatGPT Actions can only send plain JSON request
bodies, so a thin proxy sits in between: it accepts `{"content": "<html>...</html>"}`
as JSON and builds the multipart body Graph expects. Every other request
(GET/POST/PATCH/DELETE on any other Graph path) is forwarded unchanged —
the caller's `Authorization` header is passed straight through to Graph.

```
ChatGPT Action → Azure Function (this repo) → https://graph.microsoft.com/v1.0/...
```

## Repo contents

| File | Purpose |
|---|---|
| [src/functions/graphProxy.js](src/functions/graphProxy.js) | The proxy itself — one HTTP-triggered function, route `v1.0/{*restOfPath}` |
| [openapi.yaml](openapi.yaml) | OpenAPI 3.1 spec — import this as the GPT's Action |
| [host.json](host.json) / [local.settings.json](local.settings.json) | Azure Functions host & local runtime config |

## Prerequisites

- Node.js 18/20/22 (this proxy targets Node 20 in production — see
  [Troubleshooting](#troubleshooting))
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) (`az`), logged in: `az login`
- An Azure subscription and a Microsoft Entra ID (Azure AD) tenant you can register apps in
- A ChatGPT plan that supports Custom GPTs with Actions

## Local development

```bash
npm install
npm start          # runs `func start`, listens on http://localhost:7071
```

Test without a real token (expects a 401 from the proxy's own check):

```bash
curl http://localhost:7071/v1.0/me/onenote/notebooks
```

## Deploying to Azure

```bash
# variables — adjust to your own naming
RG=remsey-onenote-proxy-rg
LOCATION=westeurope
STORAGE=remseyonenoteproxyz3oh
FUNCTIONAPP=remsey-onenote-proxy-z3oh

az group create --name $RG --location $LOCATION

az storage account create --name $STORAGE --location $LOCATION \
  --resource-group $RG --sku Standard_LRS

az functionapp create --resource-group $RG --consumption-plan-location $LOCATION \
  --runtime node --runtime-version 20 --functions-version 4 \
  --name $FUNCTIONAPP --storage-account $STORAGE --os-type Linux

# deploy the code
func azure functionapp publish $FUNCTIONAPP
```

**Important:** use `--runtime-version 20` (or 22). Node 24 is listed in
`az functionapp list-runtimes` but is not reliably available on Linux
Consumption yet — see [Troubleshooting](#troubleshooting) for what happens if
you use it.

This project's live deployment: Function App `remsey-onenote-proxy-z3oh` in
resource group `remsey-onenote-proxy-rg`, Linux Consumption plan, `Node|20`.

## Microsoft Entra ID app registration (for ChatGPT OAuth)

The proxy itself doesn't care how the caller got its token — it just forwards
whatever `Authorization` header it receives to Graph. ChatGPT needs an app
registration to run the OAuth authorization-code flow against.

```bash
TENANT_ID=$(az account show --query tenantId -o tsv)

# 1. create a multitenant app registration — any work/school Microsoft 365
# tenant can sign in, not just this one. Use AzureADMyOrg instead if you only
# ever want your own tenant's users.
az ad app create --display-name "OneNote ChatGPT Action" \
  --sign-in-audience AzureADMultipleOrgs

APP_ID=<appId from the output above>

# 2. add delegated Microsoft Graph permissions: Notes.ReadWrite + offline_access
az ad app permission add --id $APP_ID \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions \
    615e26af-c38a-4150-ae3e-c3b0d4cb1d6a=Scope \
    7427e0e9-2fba-42fe-b0c0-848c9e6a8182=Scope

# 3. create a client secret (shown once — copy it into ChatGPT's Action auth panel)
az ad app credential reset --id $APP_ID --append --display-name "chatgpt-action" --years 2
```

Both `Notes.ReadWrite` and `offline_access` are user-consentable scopes, so
**admin consent is not required from any tenant** — every signed-in user, in
their own organization, approves them individually on first login through
ChatGPT's OAuth screen. No one needs to pre-register or be added anywhere;
sign-in and consent alone is what grants access, scoped to that user's own
OneNote.

### Already have a single-tenant registration?

You don't need to create a new one — switch the existing registration to
multitenant in place (this keeps the same Client ID/Secret, so the ChatGPT
Action doesn't need to change):

```bash
az ad app update --id $APP_ID --sign-in-audience AzureADMultipleOrgs

# verify:
az ad app show --id $APP_ID --query signInAudience -o tsv
# expect: AzureADMultipleOrgs
```

This project's app registration: `OneNote ChatGPT Action`, app id
`f1d0f17d-e587-40d2-a698-fb0bd8979404`, home tenant `a300b38b-6acb-43db-b02a-af94b4300d87`
(skills4-it.nl) — multitenant (`AzureADMultipleOrgs`), so users from any
other Microsoft 365 work/school tenant can sign in too. Personal Microsoft
accounts (`@outlook.com`, `@live.com`) are intentionally not supported.

## Setting up the Custom GPT

1. In the GPT Builder, go to **Configure → Actions → Create new action**.
2. Import [openapi.yaml](openapi.yaml) (paste the raw content or upload the file).
   All parameters are inlined rather than using `$ref` — ChatGPT's importer
   doesn't resolve `$ref` inside a `parameters` array, only inside schemas.
3. Set **Authentication → OAuth** with:

   | Field | Value |
   |---|---|
   | Client ID | `f1d0f17d-e587-40d2-a698-fb0bd8979404` |
   | Client Secret | *(the secret from step 3 above — never commit this)* |
   | Authorization URL | `https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize` |
   | Token URL | `https://login.microsoftonline.com/organizations/oauth2/v2.0/token` |
   | Scope | `https://graph.microsoft.com/Notes.ReadWrite offline_access` |
   | Token Exchange Method | Default (POST request) |

   Use the `/organizations/` path, not a specific tenant ID — that's what
   makes sign-in multitenant. It accepts a user from *any* Microsoft 365
   work/school tenant (each signs in against their own tenant and consents
   individually), while still rejecting personal `@outlook.com`/`@live.com`
   accounts. Using a tenant-specific GUID here instead would restrict
   sign-in to that one organization only.

   #### What each field means

   - **Authentication Type** — set to **OAuth**, not "API Key". This tells
     ChatGPT to run a full browser-based sign-in flow (redirect to Microsoft,
     user logs in, redirect back with a code, exchange for a token) instead
     of just sending a static key on every request. Required here because
     Graph needs a *delegated, per-user* token — there's no fixed "API key"
     that would let the proxy know which user's OneNote to read.

   - **Client ID** — `f1d0f17d-e587-40d2-a698-fb0bd8979404`, the public
     identifier of the Entra app registration (`az ad app create` in step 1
     above returns this as `appId`). Identifies *which application* is asking
     Microsoft for access — not a secret, safe to show in a URL or commit.

   - **Client Secret** — the value from `az ad app credential reset` in step
     3 above, shown once at creation time. Proves to Microsoft that it's
     really *this* app (not an impersonator using the same Client ID) asking
     for a token. Never commit it, never put it in this repo — it only lives
     in ChatGPT's Action auth panel, which stores it encrypted and never
     shows it back to you after saving.

   - **Authorization URL** —
     `https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize`.
     The page ChatGPT sends the user's browser to *first* — this is the
     actual Microsoft login/consent screen. The path segment after
     `login.microsoftonline.com` controls who's allowed to sign in: a
     specific tenant GUID there restricts sign-in to that one Entra tenant;
     `organizations` (used here) accepts any work/school tenant; `common`
     would also accept personal Microsoft accounts, which this app
     deliberately doesn't support.

   - **Token URL** —
     `https://login.microsoftonline.com/organizations/oauth2/v2.0/token`.
     Where ChatGPT calls, server-to-server (browser never sees this step),
     *after* the user consents — it exchanges the authorization code from
     the callback for an actual access token, using the Client ID + Client
     Secret as proof of identity. Must use the same path segment
     (`organizations`) as the Authorization URL above.

   - **Scope** — `https://graph.microsoft.com/Notes.ReadWrite offline_access`,
     space-separated. This is what ChatGPT actually asks Microsoft for
     permission to do, shown to the user on the consent screen:
     - `Notes.ReadWrite` — delegated permission to read/write the signed-in
       user's own OneNote notebooks (never more than that — see
       [Permissions strategy](../onenote-graph-proxy/NoteBuddy%20—%20Product%20&%20Technical%20Specification.md)
       for why this repo deliberately doesn't request broader Graph scopes).
     - `offline_access` — lets Microsoft also hand back a *refresh* token,
       so ChatGPT can silently get new access tokens later without forcing
       the user to log in again every ~60–90 minutes.
     Must exactly match (a subset of) what was granted with
     `az ad app permission add` in step 2 further up — asking for a scope
     that isn't on the app registration causes consent to fail.

   - **Token Exchange Method** — **Default (POST request)** vs **Basic
     authorization header**. Both send the Client ID + Client Secret to the
     Token URL; they only differ in *how*:
     - *Default (POST request)* — `client_id` and `client_secret` go in the
       POST body, alongside `code`, `redirect_uri`, `grant_type`. This is
       what Microsoft's `v2.0/token` endpoint expects, and what this project
       uses.
     - *Basic authorization header* — the same credentials go in an
       `Authorization: Basic base64(client_id:client_secret)` HTTP header
       instead of the body. Some OAuth providers require this form, but
       Microsoft's endpoint accepts the body form, so there's no reason to
       switch it here.

   - **The warning about redirect URLs** ("*OAuth may fail if you don't allow
     our redirect URLs*") is exactly the callback-URL requirement explained
     below — ChatGPT is telling you it will redirect back to
     `https://chat.openai.com/aip/<this-gpt-id>/oauth/callback`, and that URL
     must be registered on the Entra app or Microsoft will refuse the
     redirect (`AADSTS50011`).

4. Save. ChatGPT now shows a callback URL like
   `https://chat.openai.com/aip/g-.../oauth/callback`. Add it as a **Web**
   redirect URI on the app registration:

   ```bash
   az ad app update --id $APP_ID --web-redirect-uris "<callback URL from ChatGPT>"
   ```

   #### What this callback URL actually is

   This URL is generated **by ChatGPT**, not by you or by this proxy. It only
   appears after you fill in Client ID/Secret and the two Microsoft endpoints
   in step 3 and save the Action — ChatGPT then shows something like:

   ```
   https://chat.openai.com/aip/g-d0ce268413020bbc99dd5b70bc3aa6be2f2de5eb/oauth/callback
   ```

   The `g-d0ce2684...` part is a ChatGPT-side ID for *this specific GPT*, not
   for the Action or for your Azure Function. It's where **Microsoft** (not
   your proxy) redirects the user's browser back to *ChatGPT's own servers*
   after they sign in and consent — ChatGPT then exchanges the returned
   authorization code for an access token behind the scenes and uses that
   token on every Action call it makes to your proxy afterwards.

   The full flow, end to end:

   ```
   1. User (in ChatGPT) asks something that needs OneNote
   2. ChatGPT redirects the browser to the Authorization URL
      (login.microsoftonline.com/.../authorize)
   3. User signs in with their Microsoft account and consents
      to Notes.ReadWrite + offline_access
   4. Microsoft redirects the browser back to the callback URL
      https://chat.openai.com/aip/g-.../oauth/callback?code=...
   5. ChatGPT exchanges that code for an access token by calling
      the Token URL (login.microsoftonline.com/.../token), using
      the Client ID + Client Secret you configured
   6. ChatGPT stores the token and attaches it as
      "Authorization: Bearer <token>" on every subsequent call to
      this proxy (e.g. GET /v1.0/me/onenote/notebooks)
   7. This proxy just forwards that header straight through to
      Microsoft Graph — it never sees the Client Secret and never
      handles the OAuth exchange itself
   ```

   Why it must be registered on the Entra app: Microsoft Entra only redirects
   back to URIs it was explicitly told to trust (the **redirect URI**
   allowlist). If `https://chat.openai.com/aip/g-.../oauth/callback` isn't on
   that list, sign-in fails with `AADSTS50011` (see
   [Troubleshooting](#troubleshooting)) — Microsoft refuses to send the
   authorization code anywhere it doesn't recognize, which is what stops a
   malicious site from intercepting it.

   Two things worth knowing:

   - **It's per-GPT, not per-Action-config.** If you delete this GPT (or its
     Action) and recreate it, ChatGPT generates a *new* `g-...` id and thus a
     *new* callback URL — the old one stays registered on the Entra app
     (harmless, but you should add the new one too) and sign-in breaks until
     you do.
   - **It's not a secret.** It's a public redirect target, safe to commit —
     unlike the Client Secret, which must never be committed. This repo's
     concrete callback URL is captured in
     [openapi.notebuddy-gpt.yaml](openapi.notebuddy-gpt.yaml) for reference.

5. Test the connection with a read-only prompt first (forces the OAuth
   consent screen), then a write prompt:

   ```
   Show my OneNote notebooks.
   ```
   ```
   Create a page titled "Test" in my "remsey-os" notebook with a short note
   confirming the OneNote Action works.
   ```

6. Custom GPTs with Actions require a **Privacy Policy URL** (GPT Builder →
   **Configure**, near the bottom). This repo includes one:
   [privacy-policy.html](privacy-policy.html). Host it somewhere public (e.g.
   serve it as a static file from the same Function App, or any static host)
   and paste that URL into the GPT Builder's Privacy Policy field.

### Suggested GPT instructions (OneNote section)

Paste this into the GPT's **Instructions** field (merge it into a broader
agent prompt if you have one covering other Actions too):

```
## OneNote routing

- For every OneNote request — reading, searching, creating, updating, copying,
  moving or deleting notebooks, sections, section groups or pages — always call
  the OneNote Graph Proxy Action (`listNotebooks`, `createNotebook`,
  `listSectionsInNotebook`, `createSection`, `listAllSections`,
  `listSectionGroupsInNotebook`, `createSectionGroup`, `createSectionInGroup`,
  `listPagesInSection`, `createPage`, `listPages`, `getPage`, `getPageContent`,
  `updatePageContent`, `addImageToPage`, `addAttachmentToPage`,
  `copyPageToSection`, `movePage`, `copySectionToNotebook`,
  `copySectionToSectionGroup`, `deletePage`). Never answer a OneNote question
  from memory and never invent notebook, section or page names. Renaming or
  deleting a notebook or section is **not possible** — Graph has no
  update/delete operation for them; tell the user to do it directly in
  OneNote instead.
- To reuse a course or client template, call `copySectionToNotebook` (or
  `copySectionToSectionGroup`) with the template `sectionId` and the
  destination notebook/section-group `id` — this copies the whole section
  with all its pages in one call, instead of recreating pages one by one.
  Like `copyPageToSection`, Graph processes this asynchronously (202
  response); don't claim the copy is finished until the user confirms it if
  that matters.
- The user will refer to notebooks, sections and pages by name, not by ID. Resolve
  the exact `id` first with `listNotebooks`, `listSectionsInNotebook`, `listAllSections`
  or `listPagesInSection` (use `$select=id,displayName` or `id,title` to keep the
  response small), matching on `displayName`/`title`. Only then call the operation
  that needs that ID.
- If more than one notebook, section or page matches closely, present the short
  list of candidates and ask which one is meant instead of guessing.
- Prefer an existing relevant notebook and section over creating duplicates. To
  create a new section for a course/client, call `createSection` with the
  notebook id and a `displayName` rather than asking the user to do it manually.
- When reading a page to answer a question, search, or summarize, call
  `getPageContent` with `format=markdown` — it's cleaner and cheaper in tokens
  than raw HTML. Only use the default (no `format`, i.e. HTML) when you're about
  to call `updatePageContent` and need real element `data-id`s.
- When searching across several candidate pages (`listPages`/`listPagesInSection`
  with `search` or `$filter`), add `includePreview=true` to get a short text
  snippet per result (first 10 only) instead of calling `getPageContent` on each
  candidate separately to judge relevance.
- To change an existing page, read its content first with `getPageContent`
  (default HTML format), then call `updatePageContent` with commands describing
  the change (see below) rather than creating a duplicate page.

## Creating OneNote pages (createPage)

1. Call `createPage` with `sectionId` (the exact section id, never its display
   name) and a JSON body `{"content": "<full HTML document as a string>"}`.
   This is the only accepted shape — never send raw `text/html`, never omit
   `content`.
2. Resolve `sectionId` first via `listNotebooks` / `listSectionsInNotebook` /
   `listAllSections`, matching on `displayName`. Ask if the target section is
   unclear.
3. The HTML inside `content` must be a complete document: `<!DOCTYPE html>`,
   one `<html>` with `<head>`/`<body>`, a `<title>` (becomes the page title),
   all visible content inside `<body>`.
4. Use only HTML OneNote supports: `<h1>`–`<h6>`, `<p>`, `<ul>`/`<ol>`/`<li>`,
   `<table>`/`<tr>`/`<td>`, `<b>`, `<i>`, `<a>`, `<img>`. No `<script>`, no
   external CSS, no JavaScript.
5. Generate the HTML yourself in full — don't ask the user for raw HTML unless
   they explicitly offer it.
6. Never invent a `sectionId` or content. On success (201), confirm briefly
   with the page title and link if available. On error, state the error code
   and likely cause (wrong `sectionId`, or missing/invalid `content`).

## Updating OneNote pages (updatePageContent)

1. Read the page first with `getPageContent` so you know its current structure
   before proposing a change.
2. Call `updatePageContent` with `pageId` and a JSON body
   `{"commands": [...]}` — each command has `target` (`"body"` or an element's
   `data-id`), `action` (`append`, `insert`, `prepend`, `replace` or `delete`),
   optional `position` (`before`/`after`, with `insert`), and `content` (an
   HTML fragment, omitted for `delete`).
3. Never invent a `pageId` or guess at element `data-id`s that weren't in the
   content you just read. If the requested change needs a `data-id` you don't
   have, fall back to `target: "body"` with `append`/`prepend`.
4. A successful update returns 204 with no body — confirm briefly that the
   page was updated. On error, state the error code and likely cause (wrong
   `pageId`, or missing/invalid `commands`).

## Adding images and attachments (addImageToPage / addAttachmentToPage)

1. Call `addImageToPage` with `pageId` and `{"data": "<base64>", "contentType":
   "image/png", "alt": "optional description", "placement": "append"}` for
   screenshots, diagrams or slide excerpts. `contentType` must be an `image/*`
   MIME type.
2. Call `addAttachmentToPage` with `pageId` and `{"data": "<base64>", "contentType":
   "application/pdf", "fileName": "handout.pdf", "placement": "append"}` for any
   other file (PDFs, handouts, documents).
3. Both accept files up to 25 MB decoded; a larger file returns 413 — tell the
   user to shrink it. A successful call returns 204 with no body.
4. Never invent a `pageId`, and never fabricate base64 data — only use file
   content the user actually provided in the conversation.

## Copying and moving pages (copyPageToSection / movePage)

1. `copyPageToSection` duplicates a page into another section — call it with
   `pageId` and `{"id": "<destination section id>"}`. Graph processes this
   asynchronously and returns 202; the copy itself is usually done within
   seconds, but this Action does not wait for it, so don't claim the copy is
   finished until the user has confirmed it if that matters.
2. `movePage` moves a page (copy, verify, then delete the original) — call it
   with `pageId` and `{"destinationSectionId": "<section id>"}`. It returns 200
   with `newPageId` once genuinely done. If it returns 202 the copy is still
   running and the original was **not** deleted — tell the user to retry
   shortly rather than assuming it worked. On a `502`, the original may still
   exist unchanged, or (rarely) a duplicate may exist — surface the error
   message as-is rather than guessing what happened.
3. Resolve both the page and the destination section's exact `id`s first
   (never their display names) via the usual list/find operations.
```

### Updating existing page content (updatePageContent)

`PATCH .../pages/{pageId}/content` accepts a JSON body with a `commands`
array — the proxy wraps it in the multipart/form-data `Commands` part Graph
requires:

```json
{
  "commands": [
    { "target": "body", "action": "append", "content": "<p>Extra note added later.</p>" }
  ]
}
```

Each command has:

| Field | Required | Notes |
|---|---|---|
| `target` | yes | `"body"`, or the `data-id` of an existing element on the page |
| `action` | yes | one of `append`, `insert`, `prepend`, `replace`, `delete` |
| `position` | no | `"before"` / `"after"`, used with `insert` |
| `content` | no | HTML fragment; omit for `delete` |

A successful update returns `204 No Content`. See
[Microsoft's command reference](https://learn.microsoft.com/graph/onenote-update-page)
for the full set of supported targets and actions.

## Known limitations

- **No CORS handling.** Fine for server-side Action calls; add CORS headers
  if you ever call this from a browser.
- **Anonymous `authLevel`.** The function itself checks for an `Authorization`
  header and forwards it to Graph, but there's no Azure Functions-level key —
  anyone who finds the URL can attempt a call (Graph still rejects it without
  a valid token, but keep that in mind before treating the URL as a secret).

## Troubleshooting

**Everything returns `503 The service is unavailable`, including the Kudu/SCM
site.** This happened when `linuxFxVersion` was set to `Node|24` on the Linux
Consumption plan — `24` shows up in `az functionapp list-runtimes` but the
Consumption compute cluster didn't reliably run it, so the whole container
failed to start (confirmed via `az functionapp function list` failing with
`ServiceUnavailable from host runtime`, and a plain `az functionapp restart`
did **not** fix it). Fix:

```bash
az functionapp config set --name $FUNCTIONAPP --resource-group $RG \
  --linux-fx-version "Node|20"
```

**`AADSTS50011: The redirect URI ... does not match ...`** — the redirect URI
ChatGPT sends isn't registered on the app. Add it (see step 4 above); it
changes if you delete and recreate the GPT's Action.

**GPT Action import warnings like `parameter {'$ref': ...} has missing or
non-string name; skipping`.** ChatGPT's OpenAPI importer doesn't resolve
`$ref` inside a `parameters` array. Keep parameters inlined per operation, as
in the current [openapi.yaml](openapi.yaml) (`$ref` in schemas/responses is fine).
