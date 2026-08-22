const { app } = require('@azure/functions');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const ONENOTE_PAGE_CREATE_PATTERN = /^me\/onenote\/sections\/[^/]+\/pages$/i;
const ONENOTE_PAGE_UPDATE_PATTERN = /^me\/onenote\/pages\/[^/]+\/content$/i;

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
    for (const [key, value] of request.query.entries()) {
      url.searchParams.append(key, value);
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
