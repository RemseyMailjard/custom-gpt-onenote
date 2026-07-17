# OneNote Graph Proxy — Custom GPT Action

An Azure Functions (Node.js) proxy in front of Microsoft Graph `v1.0`, built so a
ChatGPT Custom GPT Action can read and write OneNote notebooks, sections and
pages on behalf of a signed-in Microsoft user.

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

# 1. create a single-tenant app registration
az ad app create --display-name "OneNote ChatGPT Action" \
  --sign-in-audience AzureADMyOrg

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
**admin consent is not required** — the signed-in user approves them on first
login through ChatGPT's OAuth screen.

This project's app registration: `OneNote ChatGPT Action`, app id
`f1d0f17d-e587-40d2-a698-fb0bd8979404`, tenant `a300b38b-6acb-43db-b02a-af94b4300d87`
(skills4-it.nl).

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
   | Authorization URL | `https://login.microsoftonline.com/a300b38b-6acb-43db-b02a-af94b4300d87/oauth2/v2.0/authorize` |
   | Token URL | `https://login.microsoftonline.com/a300b38b-6acb-43db-b02a-af94b4300d87/oauth2/v2.0/token` |
   | Scope | `https://graph.microsoft.com/Notes.ReadWrite offline_access` |
   | Token Exchange Method | Default (POST request) |

4. Save. ChatGPT now shows a callback URL like
   `https://chat.openai.com/aip/g-.../oauth/callback`. Add it as a **Web**
   redirect URI on the app registration:

   ```bash
   az ad app update --id $APP_ID --web-redirect-uris "<callback URL from ChatGPT>"
   ```

5. Test the connection with a read-only prompt first (forces the OAuth
   consent screen), then a write prompt:

   ```
   Show my OneNote notebooks.
   ```
   ```
   Create a page titled "Test" in my "remsey-os" notebook with a short note
   confirming the OneNote Action works.
   ```

### Suggested GPT instructions (OneNote section)

Paste this into the GPT's **Instructions** field (merge it into a broader
agent prompt if you have one covering other Actions too):

```
## OneNote routing

- For every OneNote request — reading, searching, creating, updating or deleting
  notebooks, sections or pages — always call the OneNote Graph Proxy Action
  (`listNotebooks`, `createNotebook`, `listSectionsInNotebook`, `listAllSections`,
  `listPagesInSection`, `createPage`, `listPages`, `getPage`, `getPageContent`,
  `updatePageContent`, `deletePage`). Never answer a OneNote question from memory
  and never invent notebook, section or page names.
- The user will refer to notebooks, sections and pages by name, not by ID. Resolve
  the exact `id` first with `listNotebooks`, `listSectionsInNotebook`, `listAllSections`
  or `listPagesInSection` (use `$select=id,displayName` or `id,title` to keep the
  response small), matching on `displayName`/`title`. Only then call the operation
  that needs that ID.
- If more than one notebook, section or page matches closely, present the short
  list of candidates and ask which one is meant instead of guessing.
- Prefer an existing relevant notebook and section over creating duplicates.
- To change an existing page, read its content first with `getPageContent`, then
  call `updatePageContent` with commands describing the change (see below) rather
  than creating a duplicate page.

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
