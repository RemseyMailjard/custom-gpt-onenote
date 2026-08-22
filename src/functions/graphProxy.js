const { app } = require('@azure/functions');
const { NodeHtmlMarkdown } = require('node-html-markdown');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const ONENOTE_PAGE_CREATE_PATTERN = /^me\/onenote\/sections\/[^/]+\/pages$/i;
const ONENOTE_PAGE_UPDATE_PATTERN = /^me\/onenote\/pages\/[^/]+\/content$/i;
const ONENOTE_PAGE_CONTENT_GET_PATTERN = /^me\/onenote\/pages\/([^/]+)\/content$/i;
const ONENOTE_PAGE_IMAGE_PATTERN = /^me\/onenote\/pages\/([^/]+)\/images$/i;
const ONENOTE_PAGE_ATTACHMENT_PATTERN = /^me\/onenote\/pages\/([^/]+)\/attachments$/i;
const ONENOTE_PAGE_MOVE_PATTERN = /^me\/onenote\/pages\/([^/]+)\/move$/i;
const ONENOTE_PAGES_LIST_PATTERN = /^me\/onenote\/pages$/i;
const ONENOTE_SECTION_PAGES_LIST_PATTERN = /^me\/onenote\/sections\/[^/]+\/pages$/i;
const MAX_RESOURCE_BYTES = 25 * 1024 * 1024;
const PREVIEW_MAX_ITEMS = 10;
const PREVIEW_MAX_CHARS = 200;
const nodeHtmlMarkdown = new NodeHtmlMarkdown();

function buildMultipartBody(html) {
  const boundary = `OneNoteBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const body =
    `--${boundary}\r\n` +
    'Content-Disposition: form-data; name="Presentation"\r\n' +
    'Content-Type: text/html\r\n\r\n' +
    `${html}\r\n` +
    `--${boundary}--\r\n`;
  return { boundary, body };
}

function buildPageUpdateMultipartBody(commands) {
  const boundary = `OneNoteBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const body =
    `--${boundary}\r\n` +
    'Content-Disposition: form-data; name="Commands"\r\n' +
    'Content-Type: application/json\r\n\r\n' +
    `${JSON.stringify(commands)}\r\n` +
    `--${boundary}--\r\n`;
  return { boundary, body };
}

function buildResourceMultipartBody(commands, partName, contentType, fileBuffer) {
  const boundary = `OneNoteBoundary${Date.now()}${Math.random().toString(16).slice(2)}`;
  const commandsPart =
    `--${boundary}\r\n` +
    'Content-Disposition: form-data; name="Commands"\r\n' +
    'Content-Type: application/json\r\n\r\n' +
    `${JSON.stringify(commands)}\r\n`;
  const filePartHeader =
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="${partName}"\r\n` +
    `Content-Type: ${contentType}\r\n\r\n`;
  const closing = `\r\n--${boundary}--\r\n`;

  const body = Buffer.concat([
    Buffer.from(commandsPart, 'utf8'),
    Buffer.from(filePartHeader, 'utf8'),
    fileBuffer,
    Buffer.from(closing, 'utf8')
  ]);
  return { boundary, body };
}

function htmlToPlainTextPreview(html, maxLength) {
  const text = html
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength).trimEnd()}…`;
}

function escapeHtmlAttribute(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function decodeBase64Payload(data) {
  const buffer = Buffer.from(data, 'base64');
  if (buffer.length === 0) return null;
  if (buffer.length > MAX_RESOURCE_BYTES) return undefined;
  return buffer;
}

async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function pollGraphOperation(operationUrl, authHeader, { maxAttempts = 15, intervalMs = 1500 } = {}) {
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const response = await fetch(operationUrl, { headers: { Authorization: authHeader } });
    if (!response.ok) {
      return { status: 'failed', error: { message: `Kon operatiestatus niet ophalen (HTTP ${response.status}).` } };
    }
    const data = await response.json();
    if (data.status === 'completed') return { status: 'completed', resourceId: data.resourceId, resourceLocation: data.resourceLocation };
    if (data.status === 'failed') return { status: 'failed', error: data.error || { message: 'Graph-operatie is mislukt.' } };
    await sleep(intervalMs);
  }
  return { status: 'pending' };
}

async function forwardResponse(graphResponse) {
  const text = await graphResponse.text();
  return {
    status: graphResponse.status,
    headers: { 'Content-Type': graphResponse.headers.get('content-type') || 'application/json' },
    body: text
  };
}

app.http('graphProxy', {
  route: 'v1.0/{*restOfPath}',
  methods: ['GET', 'POST', 'PATCH', 'DELETE'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const authHeader = request.headers.get('authorization');
    if (!authHeader) {
      return { status: 401, jsonBody: { error: { message: 'Missing Authorization header.' } } };
    }

    const restOfPath = (request.params.restOfPath || '').replace(/^\/+/, '');
    const url = new URL(`${GRAPH_BASE}/${restOfPath}`);
    const format = request.query.get('format');
    const includePreview = request.query.get('includePreview');
    for (const [key, value] of request.query.entries()) {
      if (key === 'format' || key === 'includePreview') continue;
      url.searchParams.append(key, value);
    }

    const contentGetMatch = request.method === 'GET' && restOfPath.match(ONENOTE_PAGE_CONTENT_GET_PATTERN);
    if (contentGetMatch && format === 'markdown') {
      const graphResponse = await fetch(url.toString(), { method: 'GET', headers: { Authorization: authHeader } });
      if (!graphResponse.ok) return forwardResponse(graphResponse);
      const html = await graphResponse.text();
      const markdown = nodeHtmlMarkdown.translate(html);
      return { status: 200, headers: { 'Content-Type': 'text/markdown; charset=utf-8' }, body: markdown };
    }

    const isPagesListGet =
      request.method === 'GET' &&
      (ONENOTE_PAGES_LIST_PATTERN.test(restOfPath) || ONENOTE_SECTION_PAGES_LIST_PATTERN.test(restOfPath));
    if (isPagesListGet && includePreview === 'true') {
      const graphResponse = await fetch(url.toString(), { method: 'GET', headers: { Authorization: authHeader } });
      if (!graphResponse.ok) return forwardResponse(graphResponse);
      const data = await graphResponse.json();
      const items = Array.isArray(data.value) ? data.value : [];
      await Promise.all(
        items.slice(0, PREVIEW_MAX_ITEMS).map(async (page) => {
          try {
            const contentResponse = await fetch(`${GRAPH_BASE}/me/onenote/pages/${page.id}/content`, {
              headers: { Authorization: authHeader }
            });
            if (contentResponse.ok) {
              const html = await contentResponse.text();
              page.preview = htmlToPlainTextPreview(html, PREVIEW_MAX_CHARS);
            }
          } catch {
            // best-effort enrichment — leave preview unset on failure
          }
        })
      );
      return { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) };
    }

    const imageMatch = request.method === 'POST' && restOfPath.match(ONENOTE_PAGE_IMAGE_PATTERN);
    if (imageMatch) {
      let payload;
      try {
        payload = await request.json();
      } catch {
        return { status: 400, jsonBody: { error: { message: 'Ongeldige of ontbrekende JSON-body.' } } };
      }

      const { data, contentType, alt, placement } = payload || {};
      if (typeof data !== 'string' || data.length === 0) {
        return { status: 400, jsonBody: { error: { message: "Veld 'data' (base64) is verplicht." } } };
      }
      if (typeof contentType !== 'string' || !contentType.startsWith('image/')) {
        return { status: 400, jsonBody: { error: { message: "Veld 'contentType' moet een image/* MIME-type zijn." } } };
      }
      const fileBuffer = decodeBase64Payload(data);
      if (fileBuffer === undefined) {
        return { status: 413, jsonBody: { error: { message: 'Afbeelding is te groot (max 25 MB).' } } };
      }
      if (fileBuffer === null) {
        return { status: 400, jsonBody: { error: { message: "Veld 'data' bevat geen geldige base64-inhoud." } } };
      }

      const partName = 'image1';
      const altAttr = alt ? ` alt="${escapeHtmlAttribute(alt)}"` : '';
      const commands = [
        {
          target: 'body',
          action: placement === 'prepend' ? 'prepend' : 'append',
          content: `<img src="name:${partName}"${altAttr} />`
        }
      ];
      const { boundary, body } = buildResourceMultipartBody(commands, partName, contentType, fileBuffer);
      const contentUrl = new URL(`${GRAPH_BASE}/me/onenote/pages/${imageMatch[1]}/content`);
      const graphResponse = await fetch(contentUrl.toString(), {
        method: 'PATCH',
        headers: { Authorization: authHeader, 'Content-Type': `multipart/form-data; boundary=${boundary}` },
        body
      });
      return forwardResponse(graphResponse);
    }

    const attachmentMatch = request.method === 'POST' && restOfPath.match(ONENOTE_PAGE_ATTACHMENT_PATTERN);
    if (attachmentMatch) {
      let payload;
      try {
        payload = await request.json();
      } catch {
        return { status: 400, jsonBody: { error: { message: 'Ongeldige of ontbrekende JSON-body.' } } };
      }

      const { data, contentType, fileName, placement } = payload || {};
      if (typeof data !== 'string' || data.length === 0) {
        return { status: 400, jsonBody: { error: { message: "Veld 'data' (base64) is verplicht." } } };
      }
      if (typeof contentType !== 'string' || contentType.length === 0) {
        return { status: 400, jsonBody: { error: { message: "Veld 'contentType' is verplicht." } } };
      }
      if (typeof fileName !== 'string' || fileName.length === 0) {
        return { status: 400, jsonBody: { error: { message: "Veld 'fileName' is verplicht." } } };
      }
      const fileBuffer = decodeBase64Payload(data);
      if (fileBuffer === undefined) {
        return { status: 413, jsonBody: { error: { message: 'Bijlage is te groot (max 25 MB).' } } };
      }
      if (fileBuffer === null) {
        return { status: 400, jsonBody: { error: { message: "Veld 'data' bevat geen geldige base64-inhoud." } } };
      }

      const partName = 'file1';
      const commands = [
        {
          target: 'body',
          action: placement === 'prepend' ? 'prepend' : 'append',
          content: `<object data-attachment="${escapeHtmlAttribute(fileName)}" data="name:${partName}" type="${escapeHtmlAttribute(contentType)}"></object>`
        }
      ];
      const { boundary, body } = buildResourceMultipartBody(commands, partName, contentType, fileBuffer);
      const contentUrl = new URL(`${GRAPH_BASE}/me/onenote/pages/${attachmentMatch[1]}/content`);
      const graphResponse = await fetch(contentUrl.toString(), {
        method: 'PATCH',
        headers: { Authorization: authHeader, 'Content-Type': `multipart/form-data; boundary=${boundary}` },
        body
      });
      return forwardResponse(graphResponse);
    }

    const moveMatch = request.method === 'POST' && restOfPath.match(ONENOTE_PAGE_MOVE_PATTERN);
    if (moveMatch) {
      let payload;
      try {
        payload = await request.json();
      } catch {
        return { status: 400, jsonBody: { error: { message: 'Ongeldige of ontbrekende JSON-body.' } } };
      }

      const destinationSectionId = payload && payload.destinationSectionId;
      if (typeof destinationSectionId !== 'string' || destinationSectionId.length === 0) {
        return { status: 400, jsonBody: { error: { message: "Veld 'destinationSectionId' is verplicht." } } };
      }

      const pageId = moveMatch[1];
      const copyResponse = await fetch(`${GRAPH_BASE}/me/onenote/pages/${pageId}/copyToSection`, {
        method: 'POST',
        headers: { Authorization: authHeader, 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: destinationSectionId })
      });

      if (copyResponse.status !== 202) return forwardResponse(copyResponse);

      const operationUrl = copyResponse.headers.get('operation-location') || copyResponse.headers.get('Location');
      if (!operationUrl) {
        return { status: 502, jsonBody: { error: { message: 'Graph gaf geen operation-location terug voor de copy-operatie.' } } };
      }

      const result = await pollGraphOperation(operationUrl, authHeader);
      if (result.status === 'pending') {
        return {
          status: 202,
          jsonBody: {
            success: false,
            status: 'pending',
            message: 'Kopie is nog bezig; de originele pagina is nog niet verwijderd. Probeer het straks opnieuw.'
          }
        };
      }
      if (result.status === 'failed') {
        return { status: 502, jsonBody: { error: { message: 'Kopie mislukt, originele pagina is niet verwijderd.', details: result.error } } };
      }

      const deleteResponse = await fetch(`${GRAPH_BASE}/me/onenote/pages/${pageId}`, {
        method: 'DELETE',
        headers: { Authorization: authHeader }
      });
      if (!deleteResponse.ok) {
        return {
          status: 502,
          jsonBody: {
            error: {
              message: 'Pagina is gekopieerd maar de originele pagina kon niet worden verwijderd. Er staan nu twee kopieën.',
              newPageId: result.resourceId
            }
          }
        };
      }

      return {
        status: 200,
        jsonBody: { success: true, newPageId: result.resourceId, deletedOriginal: true }
      };
    }

    const isOneNotePageCreate =
      request.method === 'POST' && ONENOTE_PAGE_CREATE_PATTERN.test(restOfPath);

    if (isOneNotePageCreate) {
      let payload;
      try {
        payload = await request.json();
      } catch {
        return { status: 400, jsonBody: { error: { message: 'Ongeldige of ontbrekende JSON-body.' } } };
      }

      const html = payload && payload.content;
      if (typeof html !== 'string' || html.length === 0) {
        return {
          status: 400,
          jsonBody: { error: { message: "Veld 'content' (volledige HTML) is verplicht en moet een string zijn." } }
        };
      }

      const { boundary, body } = buildMultipartBody(html);
      const graphResponse = await fetch(url.toString(), {
        method: 'POST',
        headers: {
          Authorization: authHeader,
          'Content-Type': `multipart/form-data; boundary=${boundary}`
        },
        body
      });
      return forwardResponse(graphResponse);
    }

    const isOneNotePageUpdate =
      request.method === 'PATCH' && ONENOTE_PAGE_UPDATE_PATTERN.test(restOfPath);

    if (isOneNotePageUpdate) {
      let payload;
      try {
        payload = await request.json();
      } catch {
        return { status: 400, jsonBody: { error: { message: 'Ongeldige of ontbrekende JSON-body.' } } };
      }

      const commands = payload && payload.commands;
      if (!Array.isArray(commands) || commands.length === 0) {
        return {
          status: 400,
          jsonBody: { error: { message: "Veld 'commands' is verplicht en moet een niet-lege array zijn." } }
        };
      }

      const { boundary, body } = buildPageUpdateMultipartBody(commands);
      const graphResponse = await fetch(url.toString(), {
        method: 'PATCH',
        headers: {
          Authorization: authHeader,
          'Content-Type': `multipart/form-data; boundary=${boundary}`
        },
        body
      });
      return forwardResponse(graphResponse);
    }

    // Transparante pass-through voor de overige operaties: ongewijzigd doorsturen naar Microsoft Graph.
    const init = { method: request.method, headers: { Authorization: authHeader } };
    if (!['GET', 'DELETE'].includes(request.method)) {
      const incomingContentType = request.headers.get('content-type');
      if (incomingContentType) init.headers['Content-Type'] = incomingContentType;
      const rawBody = await request.text();
      if (rawBody) init.body = rawBody;
    }

    const graphResponse = await fetch(url.toString(), init);
    return forwardResponse(graphResponse);
  }
});
