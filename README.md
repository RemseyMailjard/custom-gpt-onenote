# OneNote Graph Proxy — Custom GPT Action

## What this is

**NoteBuddy** (working name) lets you talk to your Microsoft OneNote notebooks
from a ChatGPT Custom GPT: find notes, read pages, and create or update pages,
all through natural conversation instead of clicking through OneNote.

This repo is the backend piece that makes that possible — a small **Azure
Functions proxy** in front of Microsoft Graph's OneNote API. It is *not* an
MCP server; it's a REST API described by an OpenAPI spec ([openapi.template.yaml](openapi.template.yaml))
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
  `$filter` (e.g. by `lastModifiedDateTime`), `$orderby`, `$expand` and
  `includePreview=true` for a plain-text snippet per result
- Read a page's content as HTML or as clean Markdown (`getPage`, `getPageContent?format=markdown`)
- Create a notebook, section, section group or page (`createNotebook`,
  `createSection`, `createSectionGroup`, `createSectionInGroup`, `createPage`)
- Update an existing page's content (`updatePageContent`)
- Add an image or file attachment to a page (`addImageToPage`, `addAttachmentToPage`)
- Copy or move a page to another section (`copyPageToSection`, `movePage`)
- Copy a whole section, or an entire notebook, template-and-all into another
  notebook/section group (`copySectionToNotebook`, `copySectionToSectionGroup`,
  `copyNotebook`)
- Delete a page (`deletePage`)

Authentication is delegated Microsoft sign-in (OAuth via Entra ID) — you only
ever see and edit your own OneNote content, scoped to `Notes.ReadWrite`.

**Not supported, by design:** renaming or deleting a notebook or section.
Microsoft Graph's OneNote API only supports List/Get/Create on notebooks and
sections — there is no update or delete operation, so faking either would mean
delete-and-recreate, which is destructive and loses page history/links. Do
this directly in OneNote instead.

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
as JSON and builds the multipart body Graph expects. The same trick covers
updating page content (`updatePageContent`'s `commands` array) and embedding
images/attachments (`addImageToPage`/`addAttachmentToPage`'s base64 `data`
field) — anywhere Graph itself demands multipart, the proxy turns it back
into plain JSON. It also adds a couple of conveniences Graph doesn't offer
directly: `getPageContent?format=markdown` (clean Markdown instead of raw
HTML), `includePreview=true` on page lists (a text snippet per result), and
`movePage` (copy, verify, then delete — Graph has no native cross-section
page move). Every other request (GET/POST/PATCH/DELETE on any other Graph
path, including the native async `copyPageToSection`, `copySectionToNotebook`,
`copySectionToSectionGroup` and `copyNotebook`) is forwarded unchanged — the
caller's `Authorization` header is passed straight through to Graph.

```
ChatGPT Action → Azure Function (this repo) → https://graph.microsoft.com/v1.0/...
```

## Repo contents

| File | Purpose |
|---|---|
| [src/functions/graphProxy.js](src/functions/graphProxy.js) | The proxy itself — one HTTP-triggered function, route `v1.0/{*restOfPath}` |
| [src/functions/setupGuide.js](src/functions/setupGuide.js) | Serves the static files in [setup-guide/](setup-guide/) (onboarding page, privacy policy, downloadable OpenAPI spec), substituting `{{PLACEHOLDER}}` tokens from App Settings at request time |
| [openapi.template.yaml](openapi.template.yaml) | OpenAPI 3.1 spec, generic — `{{PLACEHOLDER}}` tokens filled in per deployment by `scripts/render-openapi.js` |
| [azure.yaml](azure.yaml) / [infra/](infra/) | Azure Developer CLI (`azd`) project: Bicep for the Azure resources, plus a `postprovision` hook for the Entra app registration |
| [scripts/render-openapi.js](scripts/render-openapi.js) | Fills `openapi.template.yaml`'s placeholders from a `deployments/<name>/config.json` — used by the `azd` postprovision hook |
| [deployments/](deployments/) | One folder per deployment: `config.json` (public deployment values) and the generated `openapi.yaml` to import into that deployment's Custom GPT |
| [PROMPT-EXAMPLES.md](PROMPT-EXAMPLES.md) | Example prompts with their expected tool-call chains — a manual test script and a reference for tuning GPT Instructions |
| [setup-guide/](setup-guide/) | Self-service onboarding page, privacy policy, and a copy of the OpenAPI spec served live from the deployed Function App |
| [host.json](host.json) / [local.settings.json](local.settings.json) | Azure Functions host & local runtime config |

## Prerequisites

- Node.js 22 (this proxy targets Node 22 in production — Node 20 reached
  end-of-life on 2026-04-30; see [Troubleshooting](#troubleshooting))
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) (`az`), logged in: `az login`
- [Azure Developer CLI](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
  (`azd`) — for the [fast deploy path](#deploying-to-azure); the manual path only needs `az`
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

This repo is designed to be deployed **once per tenant/customer** — every
deployment is single-tenant (its own Entra app registration, its own Azure
resources), and every deployment-specific value (Function App URL, tenant ID,
Client ID) is templated out of the source, never hardcoded. See
[Deployment config](#deployment-config) below for how that works.

### Fast path: Azure Developer CLI (`azd`)

Requires [`azd`](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
(already installs alongside Azure CLI on most setups; `azd version` to check).

```bash
az login --tenant <target-tenant-id>   # log into the tenant this deployment is for

azd auth login                         # azd needs its own auth context too
azd env new <deployment-name>          # one azd environment per tenant/customer
azd up
```

`azd up` provisions the resource group, storage account, Application
Insights and Function App from [infra/main.bicep](infra/main.bicep) (Linux
Consumption, Node 22), deploys the code, then runs
[infra/hooks/postprovision.sh](infra/hooks/postprovision.sh) automatically:
that hook creates the single-tenant Entra ID app registration with
`Notes.ReadWrite`/`offline_access` permissions, wires `ENTRA_TENANT_ID`/
`ENTRA_CLIENT_ID` into the Function App's App Settings, and generates
`deployments/<deployment-name>/config.json` + `openapi.yaml`. It prints the
Client ID/Secret you need for the ChatGPT Action's Authentication panel.
**Copy the Client Secret immediately — it's shown once.**

Re-running `azd up` (or `azd provision`) later is idempotent — it reconciles
the existing resources instead of failing on "already exists", and the
postprovision hook reuses the existing Entra app registration by display
name instead of creating a duplicate. `azd down` tears the deployment back
down if a customer churns.

Bicep can't create Entra ID / Microsoft Graph objects (app registrations
aren't ARM resources), which is why that one step still shells out to `az ad`
inside a hook rather than living in `infra/*.bicep`.

### Manual path

```bash
# variables — adjust to your own naming
RG=<your-resource-group>
LOCATION=westeurope
STORAGE=<your-storage-account>   # lowercase, alphanumeric, ≤24 chars
FUNCTIONAPP=<your-function-app-name>

az group create --name $RG --location $LOCATION

az storage account create --name $STORAGE --location $LOCATION \
  --resource-group $RG --sku Standard_LRS

az functionapp create --resource-group $RG --consumption-plan-location $LOCATION \
  --runtime node --runtime-version 22 --functions-version 4 \
  --name $FUNCTIONAPP --storage-account $STORAGE --os-type Linux

# deploy the code
func azure functionapp publish $FUNCTIONAPP
```

**Important:** use `--runtime-version 22`. Node 20 reached end-of-life on
2026-04-30 and Azure now refuses to create new Function Apps on it; Node 24 is
listed in `az functionapp list-runtimes` but has not reliably worked on Linux
Consumption — see [Troubleshooting](#troubleshooting) for what happens if you
use it.

## Deployment config

`setupGuide.js` reads three App Settings at request time and substitutes them
into the pages it serves (`setup-guide/index.html`, `privacy.html`,
`openapi.notebuddy-gpt.yaml`) wherever they contain a `{{PLACEHOLDER}}` token
— nothing deployment-specific is hardcoded in those files:

| App Setting | Used for |
|---|---|
| `PROXY_BASE_URL` | The Function App's own URL, e.g. `https://acme-onenote-proxy.azurewebsites.net` |
| `ENTRA_TENANT_ID` | The Entra tenant this deployment's app registration lives in |
| `ENTRA_CLIENT_ID` | The Entra app registration's public Client ID |

`infra/hooks/postprovision.sh` sets these automatically as part of `azd up`.
`PROXY_BASE_URL` is actually set twice — once by `infra/resources.bicep`
during provisioning (a deterministic value derived from the Function App
name), once by the hook (harmless, keeps both paths self-contained). To set
them by hand:

```bash
az functionapp config appsettings set --name $FUNCTIONAPP --resource-group $RG \
  --settings \
    PROXY_BASE_URL="https://$FUNCTIONAPP.azurewebsites.net" \
    ENTRA_TENANT_ID="<tenant id>" \
    ENTRA_CLIENT_ID="<client id>"
```

The root [openapi.template.yaml](openapi.template.yaml) uses the same
`{{PLACEHOLDER}}` tokens; it's not served, so it's filled in with
`scripts/render-openapi.js` instead (also run automatically by the
`postprovision` hook):

```bash
node scripts/render-openapi.js deployments/<deployment-name>/config.json
# writes deployments/<deployment-name>/openapi.yaml
```

Every deployment's resolved config and generated spec live under
`deployments/<deployment-name>/` — check that folder in so a deployment can
be recreated or audited later. `config.json` holds only public identifiers
(tenant ID, Client ID, resource names); the Client Secret is **never** written
to disk anywhere in this repo.

## Microsoft Entra ID app registration (for ChatGPT OAuth)

The proxy itself doesn't care how the caller got its token — it just forwards
whatever `Authorization` header it receives to Graph. ChatGPT needs an app
registration to run the OAuth authorization-code flow against.
`infra/hooks/postprovision.sh` does the following automatically as part of
`azd up`; shown here for the manual path or to understand what it's doing:

```bash
TENANT_ID=$(az account show --query tenantId -o tsv)

# 1. create a single-tenant app registration — only users in this tenant can
# sign in. Use AzureADMultipleOrgs instead if you want any work/school
# Microsoft 365 tenant to be able to sign in (a multitenant SaaS-style setup).
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

### Multitenant instead of single-tenant?

If you'd rather run one shared app registration that any work/school
Microsoft 365 tenant can sign into (a multitenant SaaS-style setup, not the
default this repo is built around), switch the sign-in audience:

```bash
az ad app update --id $APP_ID --sign-in-audience AzureADMultipleOrgs

# verify:
az ad app show --id $APP_ID --query signInAudience -o tsv
# expect: AzureADMultipleOrgs
```

With a multitenant app, use `/organizations/` (not the tenant ID) in the
Authorization/Token URLs in the next section.

## Setting up the Custom GPT

1. In the GPT Builder, go to **Configure → Actions → Create new action**.
2. Import `deployments/<your-deployment-name>/openapi.yaml` (paste the raw
   content or upload the file) — or point ChatGPT's importer at
   `https://<your-function-app>.azurewebsites.net/setup-guide/openapi.notebuddy-gpt.yaml`,
   which serves the same spec with your live values already filled in.
   All parameters are inlined rather than using `$ref` — ChatGPT's importer
   doesn't resolve `$ref` inside a `parameters` array, only inside schemas.
3. Set **Authentication → OAuth** with the values `azd up` printed
   (or your own, if you registered the app manually):

   | Field | Value |
   |---|---|
   | Client ID | *(your Entra app's Client ID — public, see `deployments/<name>/config.json`)* |
   | Client Secret | *(shown once by `az ad app credential reset` — never commit this)* |
   | Authorization URL | `https://login.microsoftonline.com/<your tenant ID>/oauth2/v2.0/authorize` |
   | Token URL | `https://login.microsoftonline.com/<your tenant ID>/oauth2/v2.0/token` |
   | Scope | `https://graph.microsoft.com/Notes.ReadWrite offline_access` |
   | Token Exchange Method | Default (POST request) |

   Use your Entra tenant's GUID in the path — that's what restricts sign-in
   to a single-tenant app's own tenant. If you switched the app to
   `AzureADMultipleOrgs` per the section above, use `/organizations/` instead
   so any work/school Microsoft 365 tenant can sign in; `common` would also
   accept personal Microsoft accounts, which this app deliberately doesn't
   support.

   #### What each field means

   - **Authentication Type** — set to **OAuth**, not "API Key". This tells
     ChatGPT to run a full browser-based sign-in flow (redirect to Microsoft,
     user logs in, redirect back with a code, exchange for a token) instead
     of just sending a static key on every request. Required here because
     Graph needs a *delegated, per-user* token — there's no fixed "API key"
     that would let the proxy know which user's OneNote to read.

   - **Client ID** — the public identifier of the Entra app registration
     (`az ad app create` returns this as `appId`; the `azd` postprovision hook prints
     it, and it's saved in `deployments/<name>/config.json`). Identifies
     *which application* is asking Microsoft for access — not a secret, safe
     to show in a URL or commit.

   - **Client Secret** — the value from `az ad app credential reset`, shown
     once at creation time. Proves to Microsoft that it's really *this* app
     (not an impersonator using the same Client ID) asking for a token.
     Never commit it, never put it in this repo — it only lives in ChatGPT's
     Action auth panel, which stores it encrypted and never shows it back to
     you after saving.

   - **Authorization URL** —
     `https://login.microsoftonline.com/<tenant ID>/oauth2/v2.0/authorize`
     for a single-tenant app (the default this repo builds). The page
     ChatGPT sends the user's browser to *first* — this is the actual
     Microsoft login/consent screen. The path segment after
     `login.microsoftonline.com` controls who's allowed to sign in: a
     specific tenant GUID there restricts sign-in to that one Entra tenant;
     `organizations` accepts any work/school tenant (only if the app was
     switched to `AzureADMultipleOrgs`); `common` would also accept personal
     Microsoft accounts, which this app deliberately doesn't support.

   - **Token URL** —
     `https://login.microsoftonline.com/<tenant ID>/oauth2/v2.0/token`.
     Where ChatGPT calls, server-to-server (browser never sees this step),
     *after* the user consents — it exchanges the authorization code from
     the callback for an actual access token, using the Client ID + Client
     Secret as proof of identity. Must use the same path segment as the
     Authorization URL above.

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
  `copySectionToSectionGroup`, `copyNotebook`, `deletePage`). Never answer a
  OneNote question from memory and never invent notebook, section or page
  names. Renaming or deleting a notebook or section is **not possible** —
  Graph has no update/delete operation for them; tell the user to do it
  directly in OneNote instead.
- To reuse a course or client template, call `copySectionToNotebook` (or
  `copySectionToSectionGroup`) with the template `sectionId` and the
  destination notebook/section-group `id` — this copies the whole section
  with all its pages in one call, instead of recreating pages one by one.
  For duplicating an entire notebook (all sections and pages) as a starting
  point for a new client/cohort, use `copyNotebook` instead. Like
  `copyPageToSection`, these are asynchronous (202 response); don't claim
  the copy is finished until the user confirms it if that matters.
- `listNotebooks`, `listPages` and `listPagesInSection` accept `$orderby`
  (e.g. `lastModifiedDateTime desc`) and `$expand` (e.g. `parentSection` on
  pages) in addition to `$select`/`$top`/`$filter` — use them instead of
  sorting/filtering results yourself when the user asks for "most recently
  edited" or similar.
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
- **25 MB cap on images/attachments.** `addImageToPage`/`addAttachmentToPage`
  reject a decoded payload over 25 MB with `413`. Graph itself may have a
  lower effective limit depending on tenant/page size.
- **`movePage` can leave a duplicate in a rare failure window.** It only
  deletes the original after Graph confirms the copy succeeded; if that copy
  succeeds but the subsequent delete call itself fails, you'll have two
  copies of the page and get a `502` naming the new page's id so you can
  clean up manually. If the copy is still running after ~20 seconds it
  returns `202` and leaves the original untouched — safe to retry.
- **Notebooks and sections can't be renamed or deleted.** Not a proxy
  limitation — Microsoft Graph's OneNote API has no update/delete operation
  for them. Do this directly in OneNote.

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
  --linux-fx-version "Node|22"
```

(The original fix used `Node|20`; Node 20 reached end-of-life on 2026-04-30, so use `Node|22` now.)

**`AADSTS50011: The redirect URI ... does not match ...`** — the redirect URI
ChatGPT sends isn't registered on the app. Add it (see step 4 above); it
changes if you delete and recreate the GPT's Action.

**GPT Action import warnings like `parameter {'$ref': ...} has missing or
non-string name; skipping`.** ChatGPT's OpenAPI importer doesn't resolve
`$ref` inside a `parameters` array. Keep parameters inlined per operation, as
in [openapi.template.yaml](openapi.template.yaml) (`$ref` in schemas/responses is fine).
