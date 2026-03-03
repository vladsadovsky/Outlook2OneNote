import { createRequire } from 'module';
import envLoader from './load-root-env.cjs';

const require = createRequire(import.meta.url);
const msal = require('@azure/msal-node');
const { loadRootEnv, getEnvList } = envLoader;

loadRootEnv();

const clientId = process.env.GRAPH_CLIENT_ID;
const authority = process.env.GRAPH_AUTHORITY || 'https://login.microsoftonline.com/common';
const scopes = getEnvList('GRAPH_DEVICE_CODE_SCOPES', ['User.Read']);

if (!clientId) {
  console.error('❌ Missing GRAPH_CLIENT_ID (or VITE_CLIENT_ID) in .env');
  process.exit(1);
}

const config = {
  auth: {
    clientId,
    authority,
  }
};

const pca = new msal.PublicClientApplication(config);

const deviceCodeRequest = {
  deviceCodeCallback: (response) => {
    console.log(response.message);
  },
  scopes
};

pca.acquireTokenByDeviceCode(deviceCodeRequest).then((response) => {
  console.log('\n✅ Auth successful. Access token:');
  console.log(response.accessToken);
}).catch((error) => {
  console.error(error);
});
