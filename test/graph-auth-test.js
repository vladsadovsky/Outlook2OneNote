const msal = require('@azure/msal-node');
const readline = require('readline');

const config = {
  auth: {
    clientId: 'a73f5240-e06c-43a3-8328-1fbd80766263',
    authority: 'https://login.microsoftonline.com/common',
  }
};

const pca = new msal.PublicClientApplication(config);

const deviceCodeRequest = {
  deviceCodeCallback: (response) => {
    console.log(response.message);
  },
  scopes: ["User.Read"]
};

pca.acquireTokenByDeviceCode(deviceCodeRequest).then((response) => {
  console.log('\n✅ Auth successful. Access token:');
  console.log(response.accessToken);
}).catch((error) => {
  console.error(error);
});
