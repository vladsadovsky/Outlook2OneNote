const fs = require('fs');
const path = require('path');

function stripOptionalQuotes(value) {
  const trimmed = value.trim();
  if (
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

function loadDotEnvFile(envFilePath) {
  if (!fs.existsSync(envFilePath)) {
    return false;
  }

  const content = fs.readFileSync(envFilePath, 'utf8');
  const lines = content.split(/\r?\n/);

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) {
      continue;
    }

    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
    if (!match) {
      continue;
    }

    const key = match[1];
    const value = stripOptionalQuotes(match[2]);

    if (process.env[key] === undefined) {
      process.env[key] = value;
    }
  }

  return true;
}

function applyConvenienceMappings() {
  if (!process.env.GRAPH_CLIENT_ID && process.env.VITE_CLIENT_ID) {
    process.env.GRAPH_CLIENT_ID = process.env.VITE_CLIENT_ID;
  }

  if (!process.env.GRAPH_TENANT_ID && process.env.VITE_TENANT_ID) {
    process.env.GRAPH_TENANT_ID = process.env.VITE_TENANT_ID;
  }

  if (!process.env.AZURE_CLIENT_ID && process.env.GRAPH_CLIENT_ID) {
    process.env.AZURE_CLIENT_ID = process.env.GRAPH_CLIENT_ID;
  }

  if (!process.env.GRAPH_AUTHORITY) {
    const tenantId = process.env.GRAPH_TENANT_ID || 'common';
    process.env.GRAPH_AUTHORITY = `https://login.microsoftonline.com/${tenantId}`;
  }
}

function loadRootEnv() {
  const envFilePath =
    process.env.TEST_EXTSERVICES_ENV_FILE ||
    path.resolve(__dirname, '..', '..', '.env');

  const loaded = loadDotEnvFile(envFilePath);
  applyConvenienceMappings();

  return {
    loaded,
    envFilePath,
  };
}

function getEnvList(variableName, fallback = []) {
  const raw = process.env[variableName];
  if (!raw) {
    return fallback;
  }

  return raw
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean);
}

module.exports = {
  loadRootEnv,
  getEnvList,
};