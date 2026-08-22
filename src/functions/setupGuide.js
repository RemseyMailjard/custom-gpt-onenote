const { app } = require('@azure/functions');
const { readFile } = require('fs/promises');
const path = require('path');

const GUIDE_DIR = path.join(__dirname, '..', '..', 'setup-guide');

const FILES = {
  '': 'index.html',
  'index.html': 'index.html',
  'privacy.html': 'privacy.html',
  'openapi.notebuddy-gpt.yaml': 'openapi.notebuddy-gpt.yaml'
};

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.yaml': 'text/yaml; charset=utf-8'
};

// Deployment-specific values are never hardcoded in the served files — they're
// filled in here from App Settings, so the same code/files work unmodified
// for any tenant's deployment. See deployments/<name>/config.json for the
// values a given deployment uses, and README.md#deployment-config for setup.
const TEMPLATE_VARS = {
  '{{PROXY_BASE_URL}}': process.env.PROXY_BASE_URL || 'https://YOUR-FUNCTION-APP.azurewebsites.net',
  '{{ENTRA_TENANT_ID}}': process.env.ENTRA_TENANT_ID || 'YOUR-TENANT-ID',
  '{{ENTRA_CLIENT_ID}}': process.env.ENTRA_CLIENT_ID || 'YOUR-CLIENT-ID'
};

const TEMPLATED_FILES = new Set(['index.html', 'privacy.html', 'openapi.notebuddy-gpt.yaml']);

function renderTemplate(contents) {
  let rendered = contents;
  for (const [placeholder, value] of Object.entries(TEMPLATE_VARS)) {
    rendered = rendered.split(placeholder).join(value);
  }
  return rendered;
}

app.http('setupGuide', {
  route: 'setup-guide/{file?}',
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    const requested = (request.params.file || '').replace(/^\/+/, '');
    const fileName = FILES[requested];
    if (!fileName) {
      return { status: 404, body: 'Not found' };
    }

    try {
      const filePath = path.join(GUIDE_DIR, fileName);
      let contents = await readFile(filePath, 'utf8');
      if (TEMPLATED_FILES.has(fileName)) {
        contents = renderTemplate(contents);
      }
      const ext = path.extname(fileName);
      return {
        status: 200,
        headers: { 'Content-Type': CONTENT_TYPES[ext] || 'application/octet-stream' },
        body: contents
      };
    } catch (err) {
      context.error('setupGuide: failed to read file', err);
      return { status: 404, body: 'Not found' };
    }
  }
});
