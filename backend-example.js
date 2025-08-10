// Backend Service Example (Node.js/Express)
// File: backend/server.js

const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
require('dotenv').config(); // Load environment variables from .env file

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

// Validate required environment variables
const requiredEnvVars = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID'];
const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
  console.error('Missing required environment variables:', missingVars);
  console.error('Please check your .env file and ensure all required variables are set.');
  process.exit(1);
}

// Azure AD app configuration from environment variables
const clientConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: process.env.AUTHORITY || `https://login.microsoftonline.com/${process.env.TENANT_ID}`
  }
};

console.log('Backend service initialized with:', {
  clientId: process.env.CLIENT_ID,
  authority: clientConfig.auth.authority,
  hasClientSecret: !!process.env.CLIENT_SECRET
});

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

// Exchange authorization code for Graph API token (PKCE + Client Secret)
app.post('/api/auth/exchange-code', async (req, res) => {
  try {
    const { code, codeVerifier, redirectUri } = req.body;
    
    if (!code || !codeVerifier || !redirectUri) {
      return res.status(400).json({ 
        error: 'invalid_request', 
        error_description: 'Missing required parameters: code, codeVerifier, redirectUri' 
      });
    }
    
    // Exchange authorization code for tokens using client secret
    const tokenRequest = {
      code: code,
      scopes: ['https://graph.microsoft.com/Notes.Read', 'https://graph.microsoft.com/User.Read'],
      redirectUri: redirectUri,
      codeVerifier: codeVerifier, // PKCE parameter
    };
    
    console.log('Exchanging authorization code for tokens...');
    const response = await confidentialClientApp.acquireTokenByCode(tokenRequest);
    
    // Use the Graph API token to get notebooks
    const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
      headers: {
        'Authorization': `Bearer ${response.accessToken}`,
        'Accept': 'application/json'
      }
    });
    
    if (!graphResponse.ok) {
      const errorText = await graphResponse.text();
      console.error('Graph API request failed:', graphResponse.status, errorText);
      throw new Error(`Graph API error: ${graphResponse.status} ${graphResponse.statusText}`);
    }
    
    const notebooks = await graphResponse.json();
    
    // Return both token info and notebooks
    res.json({
      access_token: response.accessToken,
      refresh_token: response.refreshToken,
      expires_in: response.expiresOn ? Math.floor((response.expiresOn.getTime() - Date.now()) / 1000) : 3600,
      token_type: 'Bearer',
      scope: response.scopes?.join(' ') || 'https://graph.microsoft.com/Notes.Read https://graph.microsoft.com/User.Read',
      notebooks: notebooks
    });
    
  } catch (error) {
    console.error('Code exchange error:', error);
    res.status(500).json({ 
      error: 'server_error', 
      error_description: `Token exchange failed: ${error.message}`
    });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Backend service running on port ${PORT}`);
});
