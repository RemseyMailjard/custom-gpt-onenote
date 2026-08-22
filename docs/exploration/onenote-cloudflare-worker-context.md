# OneNote Capability Gateway — Cloudflare Worker Context

## 1. Purpose

Build a **Cloudflare Worker-based OneNote Capability Gateway** that exposes a safe, agent-friendly API over Microsoft Graph OneNote.

The gateway is intended for AI agents and automation clients such as:

- ChatGPT
- Microsoft Copilot / Copilot Studio
- Claude
- MCP clients
- Power Automate
- custom .NET applications
- other API consumers

The gateway MUST NOT be a generic Microsoft Graph proxy.

Its purpose is to translate high-level semantic OneNote capabilities into deterministic Microsoft Graph calls.

Core principle:

> One business capability = one predictable agent tool.

Examples:

- `onenote_list_notebooks`
- `onenote_get_notebook_tree`
- `onenote_find_pages`
- `onenote_get_page_content`
- `onenote_append_page_content`
- `onenote_add_attachment_to_page`
- `onenote_copy_page_to_section`

Do NOT expose generic operations such as:

- `callGraphApi`
- `executeGraphRequest`
- `callUrl`
- `rawOneNoteRequest`
- arbitrary OData strings
- arbitrary Graph URLs

---

# 2. High-Level Architecture

```text
AI Agent / Client
        |
        v
Cloudflare Worker
OneNote Capability Gateway
        |
        +-- Authentication / token validation
        +-- Capability routing
        +-- Context resolution
        +-- Input validation
        +-- Ambiguity detection
        +-- Confirmation policies
        +-- OData query builder
        +-- Markdown -> OneNote HTML
        +-- HTML sanitization
        +-- Multipart request builder
        +-- Cursor manager
        +-- Async operation manager
        +-- Retry / throttling
        +-- Error normalization
        +-- Audit / telemetry
        |
        v
Microsoft Graph v1.0
        |
        v
OneNote
```

The Worker is the policy and abstraction layer.

Microsoft Graph is an implementation detail and MUST NOT leak unnecessarily into the public API.

---

# 3. Target Runtime

Primary runtime:

```text
Cloudflare Workers
TypeScript
```

Recommended stack:

- Cloudflare Workers
- TypeScript
- Hono or native Worker routing
- Web Crypto API
- Cloudflare KV and/or Durable Objects for state
- Cloudflare D1 only if persistent relational audit/configuration becomes necessary
- Cloudflare Secrets for application configuration and cryptographic secrets
- Microsoft Identity Platform
- Microsoft Graph REST API v1.0

Prefer Web Platform APIs supported natively by Workers.

Avoid Node.js-only dependencies unless compatibility mode is intentionally enabled.

---

# 4. Microsoft Graph OneNote Authentication

## Critical rule

Microsoft Graph OneNote MUST be treated as **delegated-only**.

Do NOT design OneNote operations around:

```text
client_credentials
application-only
daemon/service-principal-only access
```

App-only support for Microsoft Graph OneNote is no longer supported.

The gateway should operate in the context of a signed-in Microsoft user.

Supported architectural patterns may include:

```text
Authorization Code + PKCE
delegated access token passthrough
On-Behalf-Of architecture where appropriate
```

The Worker MUST validate that the delegated token has sufficient OneNote scopes before executing an operation.

Relevant delegated scopes:

```text
Notes.Create
Notes.Read
Notes.Read.All
Notes.ReadWrite
Notes.ReadWrite.All
```

Apply least privilege per capability.

Example mapping:

```text
read operations
    -> Notes.Read or Notes.Read.All

create notebook / section / page
    -> Notes.Create where supported

modify page content
    -> Notes.ReadWrite

delete page
    -> Notes.ReadWrite
```

Do not infer permissions silently when Graph denies access.

Normalize the response to `INSUFFICIENT_PERMISSION`.

---

# 5. OneNote Context Model

OneNote Graph exposes content through different ownership/location contexts.

The gateway MUST abstract these into a typed context.

Supported context types:

```json
{
  "type": "me"
}
```

```json
{
  "type": "sharedUser",
  "userId": "..."
}
```

```json
{
  "type": "group",
  "groupId": "..."
}
```

```json
{
  "type": "site",
  "siteId": "..."
}
```

Internal mappings:

```text
me
-> /me/onenote

sharedUser
-> /users/{id}/onenote

group
-> /groups/{id}/onenote

site
-> /sites/{id}/onenote
```

Important:

`sharedUser` MUST NOT be described or implemented as unrestricted impersonation of another user's OneNote.

It represents OneNote content of that user that is accessible to the signed-in user.

Not every operation supports every context.

Each capability definition MUST declare:

```yaml
supportedContexts:
  - me
  - sharedUser
  - group
  - site
```

or the actual supported subset.

Never automatically assume all four contexts are valid for every Graph operation.

---

# 6. Core OneNote Resource Model

```text
OneNote
|
+-- notebooks[]
|   |
|   +-- sections[]
|   |   |
|   |   +-- pages[]
|   |       |
|   |       +-- HTML content
|   |       +-- images/resources
|   |       +-- attachments/resources
|   |
|   +-- sectionGroups[]
|       |
|       +-- sections[]
|       +-- sectionGroups[]
|
+-- pages[]
+-- sections[]
+-- sectionGroups[]
+-- resources/{id}/content
+-- operations/{id}
```

Section groups are recursive.

Pages have two important representations:

1. metadata as JSON
2. page body/content as OneNote HTML

Binary files and images are referenced from HTML and retrieved separately.

---

# 7. Public Capability Categories

Every public operation MUST be categorized as one of:

```text
GRAPH_PRIMITIVE
GATEWAY_COMPOSITE
AGENT_CONVENIENCE
```

Definitions:

## GRAPH_PRIMITIVE

A predictable wrapper around one supported Microsoft Graph capability.

Example:

```text
onenote_get_page
onenote_delete_page
```

## GATEWAY_COMPOSITE

The Worker performs multiple Graph calls and/or transformations.

Example:

```text
onenote_get_notebook_tree
onenote_get_page_resources
onenote_append_page_content
onenote_move_page
```

## AGENT_CONVENIENCE

A semantic agent operation built on safe typed filtering/resolution.

Example:

```text
onenote_find_notebooks
onenote_find_sections
onenote_find_pages
```

---

# 8. Recommended Public Capability Surface

## Notebook discovery

```text
onenote_list_notebooks
onenote_get_notebook
onenote_get_recent_notebooks
onenote_find_notebooks
onenote_get_notebook_tree
```

## Notebook creation

```text
onenote_create_notebook
```

## Sections

```text
onenote_list_sections
onenote_get_section
onenote_find_sections
onenote_create_section
onenote_create_section_in_group
```

## Section groups

```text
onenote_list_section_groups
onenote_get_section_group
onenote_create_section_group
```

## Pages

```text
onenote_list_pages
onenote_get_page
onenote_get_page_preview
onenote_find_pages
onenote_get_section_page_tree
onenote_create_page
onenote_delete_page
```

## Page content

```text
onenote_get_page_content
onenote_append_page_content
onenote_prepend_page_content
onenote_insert_page_content
onenote_replace_page_content
onenote_set_page_title
```

## Resources

```text
onenote_get_page_resources
onenote_get_resource
onenote_add_image_to_page
onenote_add_attachment_to_page
```

## Organization

```text
onenote_copy_page_to_section
onenote_move_page
onenote_copy_section_to_notebook
onenote_copy_section_to_section_group
onenote_copy_notebook
```

## Operations / safety

```text
onenote_prepare_confirmation
onenote_get_operation
```

---

# 9. Explicitly Unsupported Public Operations

Do NOT publish fake OneNote capabilities for operations that Microsoft Graph OneNote does not natively support.

Do not expose:

```text
onenote_delete_notebook
onenote_delete_section
onenote_delete_section_group

onenote_rename_notebook
onenote_rename_section
onenote_rename_section_group

onenote_move_notebook
onenote_move_section

onenote_export_notebook
onenote_import_notebook
onenote_backup_notebook
onenote_restore_notebook

onenote_subscribe
onenote_delta

onenote_manage_permissions
```

A page title may be changed through page-content semantics and can therefore be implemented as a composite:

```text
onenote_set_page_title
```

A page move can be implemented as a destructive composite:

```text
copy
-> poll completion
-> verify destination
-> delete original
```

It MUST be marked high risk and require confirmation.

---

# 10. API Routing

Recommended public REST shape:

```text
GET    /onenote/v1/notebooks
POST   /onenote/v1/notebooks

GET    /onenote/v1/notebooks/:notebookId
GET    /onenote/v1/notebooks/:notebookId/tree

GET    /onenote/v1/notebooks/:notebookId/sections
POST   /onenote/v1/notebooks/:notebookId/sections

GET    /onenote/v1/sections/:sectionId

GET    /onenote/v1/sections/:sectionId/pages
POST   /onenote/v1/sections/:sectionId/pages

GET    /onenote/v1/pages/:pageId
DELETE /onenote/v1/pages/:pageId

GET    /onenote/v1/pages/:pageId/preview
GET    /onenote/v1/pages/:pageId/content

POST   /onenote/v1/pages/:pageId/content/append
POST   /onenote/v1/pages/:pageId/content/prepend
POST   /onenote/v1/pages/:pageId/content/insert
POST   /onenote/v1/pages/:pageId/content/replace

PUT    /onenote/v1/pages/:pageId/title

GET    /onenote/v1/pages/:pageId/resources

POST   /onenote/v1/pages/:pageId/images
POST   /onenote/v1/pages/:pageId/attachments

POST   /onenote/v1/pages/:pageId/copy
POST   /onenote/v1/pages/:pageId/move

POST   /onenote/v1/confirmations

GET    /onenote/v1/operations/:operationId
```

Use plural nouns and resource-oriented routing.

Keep semantic behavior in `operationId` and handler definitions.

---

# 11. Capability Definition as Source of Truth

OpenAPI SHOULD NOT be the primary source of truth.

Create an internal TypeScript capability registry.

Example:

```ts
type CapabilityCategory =
  | "read"
  | "create"
  | "update"
  | "copy"
  | "delete";

type ImplementationType =
  | "graph_primitive"
  | "gateway_composite"
  | "agent_convenience";

interface CapabilityDefinition {
  operationId: string;
  category: CapabilityCategory;
  implementation: ImplementationType;

  risk: "low" | "medium" | "high";
  destructive: boolean;
  confirmationRequired: boolean;
  deterministicResourceRequired: boolean;

  delegatedScope: string;
  supportedContexts: Array<
    "me" | "sharedUser" | "group" | "site"
  >;

  method: string;
  route: string;

  graph?: {
    method: string;
    pathTemplate: string;
  };
}
```

Example:

```ts
{
  operationId: "onenote_append_page_content",
  category: "update",
  implementation: "gateway_composite",

  risk: "medium",
  destructive: false,
  confirmationRequired: false,
  deterministicResourceRequired: true,

  delegatedScope: "Notes.ReadWrite",
  supportedContexts: ["me", "sharedUser", "group", "site"],

  method: "POST",
  route: "/pages/:pageId/content/append",

  graph: {
    method: "PATCH",
    pathTemplate: "/{context}/onenote/pages/{pageId}/content"
  }
}
```

Use this registry to derive:

- route registration
- authorization rules
- OpenAPI
- MCP tool metadata
- tests
- telemetry tags

---

# 12. Agent Safety Model

Every capability MUST include agent metadata.

Example:

```yaml
x-agent:
  category: update
  risk: medium
  destructive: false
  confirmationRequired: false
  requiresDeterministicResource: true
```

Delete example:

```yaml
x-agent:
  category: delete
  risk: high
  destructive: true
  confirmationRequired: true
  requiresDeterministicResource: true
```

Risk classes:

```text
READ   -> LOW
CREATE -> MEDIUM
UPDATE -> MEDIUM
COPY   -> MEDIUM
DELETE -> HIGH
```

---

# 13. Deterministic Resource Rule

Mutating operations MUST work on resolved resource IDs.

Never perform a mutation directly from a free-text name.

Bad:

```text
append content to "Training Notes"
```

without resolving the resource first.

Required flow:

```text
find
-> candidate resolution
-> deterministic ID
-> mutation
```

If no match:

```text
RESOURCE_NOT_FOUND
```

If multiple plausible matches:

```text
AMBIGUOUS_RESOURCE
```

Never use:

```ts
results[0]
```

as an ambiguity strategy.

---

# 14. Ambiguity Handling

Normalized ambiguity response:

```json
{
  "success": false,
  "data": null,
  "paging": null,
  "meta": {
    "requestId": "gw_...",
    "apiVersion": "v1.0",
    "source": "Microsoft Graph OneNote"
  },
  "error": {
    "code": "AMBIGUOUS_RESOURCE",
    "message": "Multiple resources match the requested name.",
    "candidates": [
      {
        "id": "...",
        "title": "Power Automate"
      },
      {
        "id": "...",
        "title": "Power Automate Training"
      }
    ]
  }
}
```

The agent should resolve the ambiguity explicitly.

---

# 15. Destructive Operation Confirmation

Destructive actions require a two-step flow.

Supported initially:

```text
deletePage
movePage
```

Step 1:

```http
POST /onenote/v1/confirmations
```

Request:

```json
{
  "resourceId": "page-id",
  "operation": "deletePage"
}
```

Gateway MUST:

1. resolve the page;
2. fetch current metadata;
3. validate user access;
4. record identifying metadata;
5. create a short-lived signed confirmation token.

Response:

```json
{
  "success": true,
  "data": {
    "confirmationToken": "confirm_...",
    "operation": "deletePage",
    "resource": {
      "id": "...",
      "title": "Old Training Notes",
      "notebook": "Skills4-IT",
      "section": "Archive"
    },
    "expiresAt": "..."
  }
}
```

The token MUST be cryptographically bound to:

```text
tenant
user
operation
resourceId
resource version or lastModifiedDateTime
expiration
```

Step 2 performs the actual destructive action.

Do not accept a token prepared for another operation/resource.

---

# 16. Cloudflare State Design

Use Cloudflare state products deliberately.

## KV

Suitable for:

- short-lived opaque cursor state
- operation token mappings
- idempotency records
- capability/config cache

KV is eventually consistent.

Do not depend on KV alone when strong serialization is required.

## Durable Objects

Recommended for:

- destructive confirmation state
- idempotent mutation coordination
- async copy orchestration
- avoiding duplicate writes
- per-user or per-operation locking

Potential Durable Object classes:

```text
OperationCoordinator
ConfirmationCoordinator
IdempotencyCoordinator
```

## D1

Optional.

Use only if persistent queryable data is needed for:

- audit records
- tenant configuration
- operation history
- analytics

Do not introduce D1 merely to store transient continuation tokens.

---

# 17. Opaque Cursor Strategy

Never expose Graph:

```text
@odata.nextLink
```

Public response:

```json
{
  "paging": {
    "hasMore": true,
    "nextCursor": "cursor_..."
  }
}
```

Cursor MUST be opaque.

Suggested internal cursor payload:

```json
{
  "version": 1,
  "tenantId": "...",
  "userId": "...",
  "operationId": "onenote_list_pages",
  "context": {
    "type": "me"
  },
  "queryHash": "...",
  "graphContinuation": "...",
  "expiresAt": 0
}
```

Never serialize this as clear unsigned JSON to the client.

Use either:

1. encrypted/signed self-contained cursor; or
2. random cursor ID mapped to state in KV/Durable Objects.

A cursor MUST NOT be reusable by:

- another user;
- another tenant;
- another operation;
- another query.

---

# 18. Async Operation Strategy

OneNote copy operations may return async operation locations.

Never expose Graph `Operation-Location`.

Instead return:

```json
{
  "success": true,
  "data": {
    "operationId": "op_...",
    "status": "running"
  }
}
```

Internally store:

```text
Gateway operation ID
-> Graph Operation-Location
-> signed-in user
-> tenant
-> capability
-> created time
-> status
```

Expose:

```http
GET /onenote/v1/operations/:operationId
```

The Worker retrieves/polls Graph internally.

Possible normalized statuses:

```text
notStarted
running
completed
failed
```

---

# 19. Idempotency

Support:

```http
Idempotency-Key
```

especially for:

```text
create notebook
create section
create page
copy notebook
copy section
copy page
move page
```

Bind idempotency to:

```text
tenant
user
operationId
idempotencyKey
request payload hash
```

If the same key is reused with a different request payload, return:

```text
CONFLICT
```

Recommended behavior:

```text
new key
-> reserve execution
-> run operation
-> persist result

existing completed key
-> return original result

existing running key
-> return existing operation state
```

Durable Objects are a strong fit for mutation idempotency.

---

# 20. Graph Request Builder

All Microsoft Graph URLs MUST be generated internally.

Example:

```ts
function buildOneNoteRoot(context: GatewayContext): string {
  switch (context.type) {
    case "me":
      return "/me/onenote";

    case "sharedUser":
      return `/users/${encodeURIComponent(context.userId)}/onenote`;

    case "group":
      return `/groups/${encodeURIComponent(context.groupId)}/onenote`;

    case "site":
      return `/sites/${encodeURIComponent(context.siteId)}/onenote`;
  }
}
```

Base URL:

```text
https://graph.microsoft.com/v1.0
```

Never concatenate raw untrusted Graph path fragments from API clients.

---

# 21. Typed Query Builder

Do NOT accept:

```json
{
  "odata": "$filter=..."
}
```

Public typed filters may include:

```text
nameContains
title
createdAfter
modifiedAfter
pageSize
cursor
includeHierarchy
```

The gateway generates OData internally.

Example:

```ts
function notebookFilter(input: {
  nameContains?: string;
}): string | undefined {
  if (!input.nameContains) return undefined;

  const escaped = input.nameContains.replaceAll("'", "''");

  return `contains(displayName,'${escaped}')`;
}
```

Always URL encode generated query strings.

---

# 22. Page Content Model

OneNote page content is HTML-based.

A page consists conceptually of:

```text
metadata JSON
+
HTML content
+
binary resources
```

Graph is NOT a generic page JSON CRUD API.

The gateway must treat page-content updates separately from page metadata.

---

# 23. Create Page

Public API may accept:

```json
{
  "title": "Agent Architecture",
  "contentFormat": "markdown",
  "content": "# Architecture\n..."
}
```

or:

```json
{
  "title": "Agent Architecture",
  "contentFormat": "html",
  "content": "<h1>Architecture</h1>"
}
```

Gateway flow:

```text
validate input
-> convert Markdown when needed
-> sanitize HTML
-> wrap into valid OneNote-compatible HTML
-> POST to Graph section pages
```

Graph destination pattern:

```text
/{context}/onenote/sections/{sectionId}/pages
```

---

# 24. Markdown Support

Markdown is a Gateway feature, not a Microsoft Graph OneNote feature.

Implement:

```text
Markdown
-> sanitized OneNote-compatible HTML
```

Avoid exposing raw Markdown as though Graph natively supports it.

Recommended conversion support:

```text
headings
paragraphs
bold
italic
links
ordered lists
unordered lists
tables
code blocks as preformatted text
horizontal rules where compatible
```

Unsupported markdown features should degrade predictably or be rejected.

---

# 25. HTML Sanitization

Maintain a strict allowlist.

Possible supported elements:

```text
html
head
title
body
div
p
h1
h2
h3
h4
h5
h6
ul
ol
li
table
thead
tbody
tr
th
td
a
img
object
strong
b
em
i
u
s
strike
sub
sup
br
pre
span
```

Remove or reject:

```text
script
style injection
event handlers
onclick
onerror
javascript: URLs
unknown dangerous iframe/embed content
```

Also validate URL schemes.

Allow only explicitly approved schemes, for example:

```text
https
http where intentionally allowed
data only when intentionally supported
```

Do not rely solely on Microsoft Graph to sanitize content.

---

# 26. Reading Page Content

Public tool:

```text
onenote_get_page_content
```

Supported output formats:

```text
html
plainText
```

The Graph source remains HTML.

For `plainText`:

```text
Graph HTML
-> parse
-> remove markup
-> retain useful structural line breaks
-> return normalized text
```

Do NOT expose Graph-generated element IDs by default.

---

# 27. Page Update Mechanics

Graph page updates require:

```text
PATCH /pages/{pageId}/content
```

with change commands.

The gateway MUST hide this complexity.

The agent should not construct raw:

```json
[
  {
    "target": "...",
    "action": "...",
    "content": "..."
  }
]
```

Instead expose semantic capabilities:

```text
append
prepend
insert
replace
set title
```

---

# 28. Semantic DOM Resolution

For targeted updates:

```text
GET current content with generated IDs
-> parse HTML
-> resolve semantic target
-> require deterministic match
-> build Graph change command
-> PATCH
```

Possible semantic selectors:

```json
{
  "type": "heading",
  "text": "Architecture"
}
```

```json
{
  "type": "text",
  "text": "Current architecture"
}
```

Special targets may include:

```text
body
title
```

If no target matches:

```text
TARGET_NOT_FOUND
```

If multiple targets match:

```text
AMBIGUOUS_TARGET
```

Never guess.

Important:

Graph-generated element IDs can change after updates.

Therefore:

> Resolve IDs immediately before mutation.

Do not cache OneNote DOM IDs as durable public identifiers.

---

# 29. Page Resources

There is no normal `list resources` collection that should be exposed directly.

Implement:

```text
onenote_get_page_resources
```

as a composite:

```text
GET page HTML
-> inspect img/object/resource references
-> normalize resources
```

Example output:

```json
{
  "resources": [
    {
      "resourceId": "...",
      "type": "image",
      "contentType": "image/png",
      "name": null
    },
    {
      "resourceId": "...",
      "type": "attachment",
      "contentType": "application/pdf",
      "name": "architecture.pdf"
    }
  ]
}
```

Only return gateway-safe resource identifiers.

Do not expose arbitrary Graph resource URLs.

---

# 30. Image and Attachment Uploads

Expose simple operations:

```text
onenote_add_image_to_page
onenote_add_attachment_to_page
```

Client input:

```text
multipart/form-data
```

Gateway responsibilities:

```text
validate size
validate MIME type
sanitize filename
construct Graph multipart
construct OneNote content command
execute PATCH/POST
normalize response
```

Do not require clients/agents to understand OneNote multipart internals.

---

# 31. Page Preview Optimization

Use page preview as a lightweight discovery mechanism.

Recommended agent strategy:

```text
find candidate pages
-> retrieve previews
-> select deterministic target
-> fetch full page content only when necessary
```

This reduces:

- token consumption
- Graph traffic
- latency
- unnecessary content exposure

---

# 32. Notebook Tree

Implement:

```text
onenote_get_notebook_tree
```

as a composite optimized around Graph `$expand` where supported.

Normalized output:

```json
{
  "notebook": {
    "id": "...",
    "name": "Skills4-IT"
  },
  "children": [
    {
      "type": "section",
      "id": "...",
      "name": "Training"
    },
    {
      "type": "sectionGroup",
      "id": "...",
      "name": "Projects",
      "children": []
    }
  ]
}
```

Keep recursion depth bounded.

Protect against unexpectedly huge notebook structures.

---

# 33. Page Tree

Implement:

```text
onenote_get_section_page_tree
```

using page hierarchy metadata such as:

```text
level
order
```

and OneNote page-level query behavior.

Normalize the result into an actual nested tree.

---

# 34. Find Operations

Recommended:

```text
onenote_find_notebooks
onenote_find_sections
onenote_find_pages
```

These are NOT general full-text search endpoints.

Do not promise native OneNote Graph full-text search.

Initial page finding should rely on supported metadata and safe filters.

Potential future content search must be explicitly implemented as a Gateway feature using:

- an external index;
- stored page text;
- a Worker AI/vector search solution;
- or controlled candidate retrieval.

Do not disguise Gateway-side search as native Graph `$search`.

---

# 35. Copy Operations

Supported high-level capabilities:

```text
onenote_copy_notebook
onenote_copy_section_to_notebook
onenote_copy_section_to_section_group
onenote_copy_page_to_section
```

These can be asynchronous.

Return Gateway operation tokens.

Do not return Graph `Operation-Location`.

---

# 36. Move Page

`onenote_move_page` is a Gateway composite.

Required workflow:

```text
validate destination
-> copy page
-> poll until copy completes
-> verify copied page exists
-> confirm source still matches expected resource/version
-> delete original
-> return destination resource
```

This is destructive.

Required:

```text
risk: high
confirmationRequired: true
```

Never implement move as:

```text
copy
-> immediately delete
```

without verified copy completion.

---

# 37. Normalized Response Envelope

All JSON endpoints should use one consistent response model.

Success:

```json
{
  "success": true,
  "data": {},
  "paging": null,
  "meta": {
    "requestId": "gw_...",
    "graphRequestId": "...",
    "apiVersion": "v1.0",
    "source": "Microsoft Graph OneNote"
  },
  "error": null
}
```

List response:

```json
{
  "success": true,
  "data": {
    "items": []
  },
  "paging": {
    "hasMore": false,
    "nextCursor": null
  },
  "meta": {
    "requestId": "gw_...",
    "graphRequestId": "...",
    "apiVersion": "v1.0",
    "source": "Microsoft Graph OneNote"
  },
  "error": null
}
```

Error:

```json
{
  "success": false,
  "data": null,
  "paging": null,
  "meta": {
    "requestId": "gw_...",
    "graphRequestId": "...",
    "apiVersion": "v1.0",
    "source": "Microsoft Graph OneNote"
  },
  "error": {
    "code": "PAGE_NOT_FOUND",
    "message": "The requested page was not found.",
    "graphStatus": 404,
    "graphCode": "itemNotFound",
    "oneNoteCode": null,
    "requestId": "...",
    "retryAfterSeconds": null
  }
}
```

---

# 38. Gateway Error Codes

Normalize Microsoft Graph and OneNote errors to stable gateway codes.

Initial set:

```text
BAD_REQUEST

AUTHENTICATION_REQUIRED
INSUFFICIENT_PERMISSION

RESOURCE_NOT_FOUND
NOTEBOOK_NOT_FOUND
SECTION_NOT_FOUND
PAGE_NOT_FOUND

AMBIGUOUS_RESOURCE
TARGET_NOT_FOUND
AMBIGUOUS_TARGET

INVALID_PAGE_CONTENT
UNSUPPORTED_CONTEXT

CONFLICT
PAYLOAD_TOO_LARGE

THROTTLED
GRAPH_UNAVAILABLE
GATEWAY_ERROR
```

Preserve diagnostics:

```text
graphStatus
graphCode
oneNoteCode
graphRequestId
retryAfterSeconds
```

Never send Graph stack traces or token details to clients.

---

# 39. Graph Error Mapping

Create a dedicated mapper:

```ts
interface GraphErrorContext {
  status: number;
  graphCode?: string;
  oneNoteCode?: string;
  message?: string;
  requestId?: string;
  retryAfterSeconds?: number;
}

function normalizeGraphError(
  error: GraphErrorContext
): GatewayError;
```

Mapping should consider:

```text
HTTP status
Graph error.code
deepest innerError.code
OneNote numeric error code
operation context
```

---

# 40. Throttling

Microsoft Graph may return:

```text
429 Too Many Requests
Retry-After
```

Gateway behavior:

```text
read Retry-After
-> determine whether retry is safe
-> retry or return normalized THROTTLED
```

Automatically retry safe reads with bounded retries.

For mutations:

- retry only when execution state is known;
- use idempotency;
- never blindly replay uncertain mutations.

Recommended retry classes:

```text
GET
-> retryable

POST create
-> only retry with idempotency protection

PATCH
-> only retry if safe/known

DELETE
-> never blindly retry after ambiguous network failure
```

Use exponential backoff plus Graph `Retry-After`.

---

# 41. Request Correlation

Every incoming gateway request gets:

```text
requestId
```

Use a Worker-generated UUID.

Forward Graph client request correlation where useful.

Capture Graph response request ID.

Logs should include:

```text
gatewayRequestId
graphRequestId
operationId
tenantId hash
userId hash
status
duration
retry count
```

Do NOT log:

- access tokens
- refresh tokens
- full page content by default
- attachment binary content
- sensitive user data unless explicitly required

---

# 42. Token Handling

Treat access tokens as secrets.

Rules:

```text
never log token
never return token
never store token in plain KV
prefer request-scoped use
```

If token caching/refresh is added later:

- encrypt at rest;
- minimize lifetime;
- isolate by tenant/user;
- rotate encryption keys safely.

---

# 43. Input Validation

Every public operation must validate:

```text
route parameters
context discriminators
body schema
content length
filenames
MIME types
date formats
enum values
page size limits
```

Reject unknown object properties where practical.

Use strict schemas.

Do not accept huge arbitrary objects.

---

# 44. Size Limits

Enforce gateway-level limits before sending content to Graph.

Configuration examples:

```text
MAX_HTML_BYTES
MAX_MARKDOWN_BYTES
MAX_ATTACHMENT_BYTES
MAX_IMAGE_BYTES
MAX_BATCH_RESOURCE_COUNT
```

Values should be configurable through environment variables/configuration.

Keep Graph product limits separately documented.

Do not hard-code guessed Microsoft limits without verification.

---

# 45. Secrets / Cloudflare Environment

Suggested Worker bindings:

```text
MS_CLIENT_ID
MS_TENANT_MODE
CURSOR_SIGNING_SECRET
CONFIRMATION_SIGNING_SECRET
INTERNAL_API_SECRET optional
```

Possible bindings:

```text
ONENOTE_STATE_KV
ONENOTE_OPERATIONS Durable Object
ONENOTE_CONFIRMATIONS Durable Object
```

Never commit secret values.

Use:

```bash
wrangler secret put ...
```

for production secrets.

---

# 46. Suggested Project Structure

```text
src/
|
+-- index.ts
|
+-- routes/
|   +-- notebooks.ts
|   +-- sections.ts
|   +-- section-groups.ts
|   +-- pages.ts
|   +-- content.ts
|   +-- resources.ts
|   +-- operations.ts
|   +-- confirmations.ts
|
+-- capabilities/
|   +-- registry.ts
|   +-- types.ts
|
+-- graph/
|   +-- client.ts
|   +-- onenote-paths.ts
|   +-- query-builder.ts
|   +-- errors.ts
|   +-- retry.ts
|
+-- auth/
|   +-- bearer.ts
|   +-- scopes.ts
|
+-- onenote/
|   +-- html.ts
|   +-- markdown.ts
|   +-- sanitizer.ts
|   +-- dom-resolver.ts
|   +-- multipart.ts
|   +-- resources.ts
|
+-- state/
|   +-- cursor.ts
|   +-- operations.ts
|   +-- confirmations.ts
|   +-- idempotency.ts
|
+-- api/
|   +-- envelopes.ts
|   +-- errors.ts
|   +-- schemas.ts
|
+-- telemetry/
|   +-- logging.ts
|   +-- correlation.ts
|
+-- openapi/
    +-- generate.ts
```

---

# 47. Graph Client

Create one controlled Graph client.

Example interface:

```ts
interface GraphRequest {
  method: "GET" | "POST" | "PATCH" | "DELETE";
  path: string;
  query?: URLSearchParams;
  headers?: HeadersInit;
  body?: BodyInit;
}

interface GraphClient {
  request<T>(
    accessToken: string,
    request: GraphRequest
  ): Promise<GraphResponse<T>>;
}
```

Only internal modules may instantiate Graph requests.

Routes should call capability services rather than Graph directly.

---

# 48. Capability Service Pattern

Preferred layering:

```text
route
-> validation
-> capability service
-> Graph helper(s)
-> normalized result
```

Example:

```ts
async function appendPageContent(
  ctx: RequestContext,
  input: AppendPageContentInput
): Promise<PageMutationResult> {
  const currentHtml =
    await oneNoteGraph.getPageContentWithIds(
      ctx,
      input.pageId
    );

  const target =
    resolveSemanticTarget(
      currentHtml,
      { type: "body" }
    );

  const safeHtml =
    await convertAndSanitizeContent(
      input.content,
      input.contentFormat
    );

  return oneNoteGraph.patchPageContent(
    ctx,
    input.pageId,
    [{
      target: target.graphId,
      action: "append",
      content: safeHtml
    }]
  );
}
```

Graph-specific IDs remain inside the service.

---

# 49. OpenAPI Generation

The existing target is an OpenAPI 3.1 contract.

The Worker project should be able to generate or validate OpenAPI from the capability registry.

Vendor metadata:

```yaml
x-agent:
  category: update
  risk: medium
  destructive: false
  confirmationRequired: false
  requiresDeterministicResource: true

x-implementation: gateway_composite

x-supported-contexts:
  - me
  - sharedUser
  - group
  - site
```

OpenAPI should expose only supported public capabilities.

---

# 50. Future MCP Generation

The same capability registry should later be usable to generate MCP tools.

Example mapping:

```text
CapabilityDefinition
        |
        +-- OpenAPI operation
        |
        +-- MCP tool
        |
        +-- route
        |
        +-- authorization rule
        |
        +-- test case
```

Do not duplicate capability descriptions manually across multiple systems if generation can prevent drift.

---

# 51. Testing Strategy

Every public capability must have:

```text
positive test
authentication test
permission test
invalid ID test
not found test
invalid body test
ambiguity test where relevant
paging test where relevant
throttling test
```

Mutation tests additionally require:

```text
create/modify
-> verify
-> cleanup where possible
```

High-risk operations additionally require:

```text
confirmation rejected
expired confirmation
wrong user confirmation
wrong resource confirmation
resource changed after confirmation
```

---

# 52. Contract Tests

Create tests asserting:

```text
every public route has CapabilityDefinition

every CapabilityDefinition has:
- operationId
- category
- risk
- delegated scope
- supported contexts
- implementation type

every public OpenAPI operation maps to one capability

no unsupported capability appears in OpenAPI
```

Also assert:

```text
no public parameter accepts arbitrary Graph URL
no public parameter accepts raw @odata.nextLink
no public parameter accepts raw Operation-Location
no public OData expression string is accepted
```

---

# 53. Security Tests

Explicitly test:

```text
Graph URL injection
OData injection
cross-user cursor reuse
cross-tenant cursor reuse
operation token reuse
confirmation token tampering
confirmation token replay
expired confirmations
HTML script injection
javascript: URL injection
malicious attachment filename
oversized upload
invalid MIME type
```

---

# 54. Observability

Recommended metrics:

```text
requests per capability
Graph latency
gateway latency
Graph errors
normalized error codes
429 rate
retry rate
ambiguity rate
confirmation rate
operation completion time
idempotency hits
```

Useful Cloudflare features:

- Workers Analytics
- structured `console.log` / Logpush
- Analytics Engine if deeper product metrics are needed

Do not log full OneNote content by default.

---

# 55. Versioning

Public gateway version:

```text
/onenote/v1
```

Microsoft Graph version used internally:

```text
v1.0
```

Do not expose Graph beta functionality in the stable public v1 unless explicitly isolated.

If beta capabilities are explored later:

```text
experimental capability flag
or
/onenote/experimental
```

Keep beta dependencies visibly separate.

---

# 56. Error Response HTTP Statuses

Recommended public mapping:

```text
400 BAD_REQUEST
401 AUTHENTICATION_REQUIRED
403 INSUFFICIENT_PERMISSION
404 *_NOT_FOUND
409 AMBIGUOUS_RESOURCE / CONFLICT
413 PAYLOAD_TOO_LARGE
429 THROTTLED
503 GRAPH_UNAVAILABLE
500 GATEWAY_ERROR
```

Use HTTP status plus stable gateway error code.

Do not return all Graph failures as HTTP 200.

---

# 57. Current Negative Capability Policy

A capability that sounds logical is NOT considered supported until it is:

```text
documented
or
observed and intentionally implemented as a Gateway composite
```

Classifications:

```text
SUPPORTED
PARTIALLY_SUPPORTED
NOT_SUPPORTED
INDIRECTLY_POSSIBLE
GATEWAY_COMPOSITE
UNKNOWN
```

Never invent missing Graph endpoints.

---

# 58. Implementation Priorities

Build in this order.

## Phase A — Read-only core

```text
auth
Graph client
context resolver
response envelope
error mapper

list notebooks
get notebook
list sections
list pages
get page
page preview
get page content
```

## Phase B — Safe discovery

```text
find notebooks
find sections
find pages
notebook tree
page tree
cursor pagination
```

## Phase C — Create

```text
create notebook
create section
create section group
create page
Markdown -> HTML
HTML sanitizer
```

## Phase D — Content mutation

```text
append
prepend
set page title
semantic target resolver
insert
replace
```

## Phase E — Resources

```text
discover page resources
download resource
add image
add attachment
multipart builder
```

## Phase F — Async organization

```text
copy page
copy section
copy notebook
operation tokens
operation polling
```

## Phase G — High-risk composites

```text
confirmation engine
delete page
move page
idempotency coordinator
```

---

# 59. Definition of Done for Worker v1

The Worker v1 is complete when:

- Microsoft delegated OneNote authentication works;
- no app-only OneNote path exists;
- no raw Graph proxy exists;
- all public operations map to capability definitions;
- all queries are typed;
- all Graph URLs are generated internally;
- cursor paging is opaque;
- async operations use gateway operation IDs;
- errors are normalized;
- Graph request IDs are preserved;
- HTML input is sanitized;
- Markdown conversion is deterministic;
- page mutation resolves current Graph element IDs;
- ambiguous resources are never silently selected;
- destructive actions require signed confirmations;
- mutation retries are safe/idempotent;
- unsupported OneNote functions are not exposed;
- every public capability has tests;
- OpenAPI 3.1 can be generated or validated from the capability registry.

---

# 60. Non-Goals for v1

Do not implement in v1:

```text
full notebook backup
full notebook restore
OneNote binary import/export
native OneNote webhook subscriptions
OneNote delta sync
raw ink editing
native OCR service
tenant-wide unattended OneNote crawling
generic Microsoft Graph proxying
arbitrary permissions management
arbitrary full-text OneNote Graph search
```

These may require different Microsoft APIs, storage-layer APIs, indexing solutions, or entirely separate services.

---

# 61. Engineering Principle

When deciding whether to expose a capability, use this decision sequence:

```text
1. Does Microsoft Graph OneNote support it?
      |
      +-- no -> do not fake it
      |
      +-- yes
            |
2. Can it be made deterministic for an agent?
      |
      +-- no -> create a safe resolver/composite first
      |
      +-- yes
            |
3. Is it destructive?
      |
      +-- yes -> confirmation boundary
      |
      +-- no
            |
4. What is the least delegated scope?
      |
5. Which OneNote contexts are actually supported?
      |
6. Create CapabilityDefinition
      |
7. Implement route/service
      |
8. Generate OpenAPI/MCP metadata
      |
9. Add tests
```

---

# 62. Final Architectural Rule

The OneNote Capability Gateway is not:

> Microsoft Graph exposed through Cloudflare.

It is:

> A controlled semantic capability layer that allows AI agents and applications to safely operate OneNote through Microsoft Graph.

The Worker owns:

```text
safety
policy
resolution
abstraction
normalization
orchestration
```

Microsoft Graph owns:

```text
OneNote storage
resource semantics
actual OneNote operations
```

Keep that boundary strict.
