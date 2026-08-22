#!/usr/bin/env node
// Fills openapi.template.yaml's {{PLACEHOLDER}} tokens from a deployment
// config.json and writes the result next to that config file.
//
// Usage: node scripts/render-openapi.js deployments/<name>/config.json

const { readFileSync, writeFileSync } = require('fs');
const path = require('path');

const configPath = process.argv[2];
if (!configPath) {
  console.error('Usage: node scripts/render-openapi.js deployments/<name>/config.json');
  process.exit(1);
}

const repoRoot = path.join(__dirname, '..');
const config = JSON.parse(readFileSync(configPath, 'utf8'));

const required = ['proxyBaseUrl', 'entraTenantId', 'entraClientId'];
const missing = required.filter((key) => !config[key]);
if (missing.length > 0) {
  console.error(`Missing required field(s) in ${configPath}: ${missing.join(', ')}`);
  process.exit(1);
}

const templatePath = path.join(repoRoot, 'openapi.template.yaml');
let rendered = readFileSync(templatePath, 'utf8');
rendered = rendered
  .split('{{PROXY_BASE_URL}}').join(config.proxyBaseUrl)
  .split('{{ENTRA_TENANT_ID}}').join(config.entraTenantId)
  .split('{{ENTRA_CLIENT_ID}}').join(config.entraClientId);

const outputPath = path.join(path.dirname(configPath), 'openapi.yaml');
writeFileSync(outputPath, rendered, 'utf8');
console.log(`Wrote ${outputPath}`);
