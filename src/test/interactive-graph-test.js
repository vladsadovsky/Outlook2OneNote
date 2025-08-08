#!/usr/bin/env node

/**
 * Interactive OneNote Test - Browser-based Authentication
 * 
 * This version opens a browser for authentication, similar to how
 * your Outlook add-in will work with SSO.
 * 
 * Prerequisites:
 * - npm install @microsoft/microsoft-graph-client isomorphic-fetch
 * 
 * This approach is closer to what your Outlook add-in will use.
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const { spawn } = require('child_process');
const http = require('http');
const url = require('url');

// Configuration - uses same client ID as your add-in
const CLIENT_ID = 'a73f5240-e06c-43a3-8328-1fbd80766263';
const REDIRECT_URI = 'http://localhost:8080/callback';
const TENANT = 'common'; // or your specific tenant ID

const SCOPES = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/User.Read',
  'openid',
  'profile'
];

class InteractiveOneNoteTest {
  constructor() {
    this.accessToken = null;
    this.graphClient = null;
  }

  /**
   * Generate the authorization URL
   */
  getAuthUrl() {
    const authParams = new URLSearchParams({
      client_id: CLIENT_ID,
      response_type: 'token',
      redirect_uri: REDIRECT_URI,
      scope: SCOPES.join(' '),
      response_mode: 'fragment'
    });

    return `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize?${authParams}`;
  }

  /**
   * Start local server to receive the auth callback
   */
  async startAuthServer() {
    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        const parsedUrl = url.parse(req.url, true);
        
        if (parsedUrl.pathname === '/callback') {
          // Send a simple HTML page that extracts the token from fragment
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>OneNote Test - Authentication</title>
                <style>
                    body { font-family: Arial, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; }
                    .success { color: green; }
                    .error { color: red; }
                </style>
            </head>
            <body>
                <h2>OneNote Test Authentication</h2>
                <div id="status">Processing authentication...</div>
                <script>
                    // Extract token from URL fragment
                    const fragment = window.location.hash.substr(1);
                    const params = new URLSearchParams(fragment);
                    const accessToken = params.get('access_token');
                    const error = params.get('error');
                    
                    if (accessToken) {
                        document.getElementById('status').innerHTML = 
                            '<div class="success">✅ Authentication successful! You can close this window and return to the terminal.</div>';
                        
                        // Send token to the Node.js server
                        fetch('/token', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ access_token: accessToken })
                        });
                    } else if (error) {
                        document.getElementById('status').innerHTML = 
                            '<div class="error">❌ Authentication failed: ' + error + '</div>';
                    } else {
                        document.getElementById('status').innerHTML = 
                            '<div class="error">❌ No authentication result found.</div>';
                    }
                </script>
            </body>
            </html>
          `);
        } else if (parsedUrl.pathname === '/token' && req.method === 'POST') {
          // Receive the token from the frontend
          let body = '';
          req.on('data', chunk => { body += chunk; });
          req.on('end', () => {
            try {
              const data = JSON.parse(body);
              this.accessToken = data.access_token;
              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ status: 'success' }));
              
              // Close the server and resolve
              server.close();
              resolve(this.accessToken);
            } catch (error) {
              res.writeHead(400, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ error: 'Invalid JSON' }));
              reject(error);
            }
          });
        } else {
          res.writeHead(404);
          res.end('Not found');
        }
      });

      server.listen(8080, () => {
        console.log('🌐 Local auth server started at http://localhost:8080');
      });

      // Timeout after 2 minutes
      setTimeout(() => {
        server.close();
        reject(new Error('Authentication timeout'));
      }, 120000);
    });
  }

  /**
   * Open browser for authentication
   */
  openBrowser(authUrl) {
    const platform = process.platform;
    let command;

    if (platform === 'win32') {
      command = 'start';
    } else if (platform === 'darwin') {
      command = 'open';
    } else {
      command = 'xdg-open';
    }

    console.log('🌐 Opening browser for authentication...');
    console.log('📝 If browser doesn\'t open, go to:', authUrl);
    
    try {
      spawn(command, [authUrl], { stdio: 'ignore', detached: true }).unref();
    } catch (error) {
      console.log('❌ Could not open browser automatically');
      console.log('📋 Please manually open:', authUrl);
    }
  }

  /**
   * Authenticate user
   */
  async authenticate() {
    console.log('🔐 Starting browser-based authentication...\n');
    
    try {
      const authUrl = this.getAuthUrl();
      
      // Start the auth server and open browser simultaneously
      const serverPromise = this.startAuthServer();
      this.openBrowser(authUrl);
      
      console.log('⏳ Waiting for authentication in browser...');
      console.log('   (This will timeout in 2 minutes)');
      
      const token = await serverPromise;
      
      if (token) {
        console.log('✅ Authentication successful!');
        
        // Initialize Graph client
        this.graphClient = Client.init({
          authProvider: {
            getAccessToken: async () => {
              return this.accessToken;
            }
          }
        });
        
        return true;
      }
    } catch (error) {
      console.error('❌ Authentication failed:', error.message);
      return false;
    }
  }

  /**
   * Test Graph API connection
   */
  async testConnection() {
    try {
      const user = await this.graphClient.api('/me').get();
      console.log(`👤 Authenticated as: ${user.displayName} (${user.userPrincipalName})`);
      return true;
    } catch (error) {
      console.error('❌ Graph API connection failed:', error.message);
      return false;
    }
  }

  /**
   * List OneNote notebooks
   */
  async listNotebooks() {
    try {
      console.log('\n📚 Fetching OneNote notebooks...');
      
      const response = await this.graphClient
        .api('/me/onenote/notebooks')
        .get();

      const notebooks = response.value || [];

      if (notebooks.length === 0) {
        console.log('📝 No notebooks found.');
        return [];
      }

      console.log(`\n📊 Found ${notebooks.length} notebook(s):\n`);
      
      notebooks.forEach((notebook, index) => {
        console.log(`${index + 1}. 📓 "${notebook.displayName}"`);
        console.log(`   🆔 ID: ${notebook.id}`);
        console.log(`   📅 Created: ${new Date(notebook.createdDateTime).toLocaleDateString()}`);
        console.log(`   📝 Modified: ${new Date(notebook.lastModifiedDateTime).toLocaleDateString()}`);
        console.log(`   ⭐ Default: ${notebook.isDefault ? 'Yes' : 'No'}`);
        console.log('');
      });

      return notebooks;
      
    } catch (error) {
      console.error('❌ Failed to fetch notebooks:', error.message);
      
      if (error.code === 'Forbidden') {
        console.log('💡 This might be a permissions issue:');
        console.log('   - Check that Notes.Read permission is granted');
        console.log('   - Ensure you have OneNote notebooks in your account');
      }
      
      return [];
    }
  }

  /**
   * Run the test
   */
  async run() {
    console.log('🚀 Interactive OneNote Test Starting...\n');
    
    // Authenticate
    const authSuccess = await this.authenticate();
    if (!authSuccess) {
      console.log('\n❌ Test failed: Could not authenticate');
      console.log('💡 Make sure your Azure AD app allows public client flows');
      return;
    }

    // Test connection
    const connectionOk = await this.testConnection();
    if (!connectionOk) {
      console.log('\n❌ Test failed: Could not connect to Graph API');
      return;
    }

    // List notebooks
    const notebooks = await this.listNotebooks();

    console.log('\n✅ Test completed successfully!');
    console.log(`📊 Summary: Found ${notebooks.length} OneNote notebook(s)`);
    console.log('\n🔗 Integration notes for your Outlook add-in:');
    console.log('   - This flow is similar to what Office.auth.getAccessToken() provides');
    console.log('   - Same Graph API endpoints will work in your add-in');
    console.log('   - Consider error handling for network issues');
  }
}

// Main execution
async function main() {
  console.log('Interactive OneNote Test - Microsoft Graph API');
  console.log('=' .repeat(50));
  
  try {
    const test = new InteractiveOneNoteTest();
    await test.run();
  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main().catch(console.error);
}

module.exports = InteractiveOneNoteTest;
