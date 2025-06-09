const msal = require('@azure/msal-node');
const readline = require('readline');

const config = {
  auth: {
    clientId: 'YOUR_CLIENT_ID_HERE',
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
