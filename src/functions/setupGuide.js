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
      const contents = await readFile(filePath, 'utf8');
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
