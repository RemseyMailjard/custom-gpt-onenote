const { TableClient } = require('@azure/data-tables');

const TABLE_NAME = 'Licenses';
const PARTITION_KEY = 'license';

let cachedClient;

function getClient() {
  if (!cachedClient) {
    const connectionString = process.env.AzureWebJobsStorage;
    cachedClient = TableClient.fromConnectionString(connectionString, TABLE_NAME, {
      allowInsecureConnection: connectionString?.includes('UseDevelopmentStorage=true')
    });
  }
  return cachedClient;
}

async function ensureTableExists() {
  const client = getClient();
  await client.createTable();
}

async function createLicense({ code, email, stripeCustomerId, stripeSessionId }) {
  const client = getClient();
  await ensureTableExists();
  await client.createEntity({
    partitionKey: PARTITION_KEY,
    rowKey: code,
    email,
    status: 'active',
    stripeCustomerId: stripeCustomerId || '',
    stripeSessionId: stripeSessionId || '',
    createdAt: new Date().toISOString()
  });
}

async function getLicense(code) {
  const client = getClient();
  try {
    return await client.getEntity(PARTITION_KEY, code);
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function revokeLicense(code) {
  const client = getClient();
  const entity = await getLicense(code);
  if (!entity) return false;
  entity.status = 'revoked';
  await client.updateEntity(entity, 'Merge');
  return true;
}

module.exports = { createLicense, getLicense, revokeLicense };
