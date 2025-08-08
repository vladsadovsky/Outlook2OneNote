// Backend Service Example (Node.js/Express)
// File: backend/server.js

const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const app = express();
app.use(express.json());

// CORS for your Office Add-in
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://localhost:3000'); // Your add-in URL
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  if (req.method === 'OPTIONS') {
    res.sendStatus(200);
  } else {
    next();
  }
});

// Azure AD app configuration
const clientConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID, // Your Azure app client ID
    clientSecret: process.env.AZURE_CLIENT_SECRET, // Your Azure app secret
    authority: 'https://login.microsoftonline.com/common'
  }
};

const confidentialClientApp = new ConfidentialClientApplication(clientConfig);

// JWKS client for token validation
const jwksUri = 'https://login.microsoftonline.com/common/discovery/v2.0/keys';
const client = jwksClient({ jwksUri });

// Validate Office SSO token
function getKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    const signingKey = key.publicKey || key.rsaPublicKey;
    callback(null, signingKey);
  });
}

// Exchange Office SSO token for Graph API token
app.post('/api/auth/exchange-token', async (req, res) => {
  try {
    const { ssoToken } = req.body;
    
    if (!ssoToken) {
      return res.status(400).json({ error: 'SSO token is required' });
    }
    
    // Validate the SSO token
    const decoded = jwt.verify(ssoToken, getKey, {
      audience: `api://${process.env.AZURE_CLIENT_ID}`, // Your Azure app ID URI
      issuer: ['https://login.microsoftonline.com/common/v2.0', 'https://sts.windows.net/common/']
    });
    
    // Exchange SSO token for Graph API token using on-behalf-of flow
    const oboRequest = {
      oboAssertion: ssoToken,
      scopes: ['https://graph.microsoft.com/Notes.Read', 'https://graph.microsoft.com/User.Read'],
    };
    
    const response = await confidentialClientApp.acquireTokenOnBehalfOf(oboRequest);
    
    // Use the Graph API token to get notebooks
    const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
      headers: {
        'Authorization': `Bearer ${response.accessToken}`,
        'Accept': 'application/json'
      }
    });
    
    if (!graphResponse.ok) {
      throw new Error(`Graph API error: ${graphResponse.status} ${graphResponse.statusText}`);
    }
    
    const notebooks = await graphResponse.json();
    res.json(notebooks);
    
  } catch (error) {
    console.error('Token exchange error:', error);
    res.status(500).json({ 
      error: 'Token exchange failed', 
      details: error.message 
    });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Backend service running on port ${PORT}`);
});
