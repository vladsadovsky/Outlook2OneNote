const express = require('express');
const path = require('path');
const https = require('https');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Enable CORS for development
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    next();
});

// Serve static files from project root
app.use(express.static(path.join(__dirname, '.')));

// Serve the PKCE debug page
app.get('/pkce-debug', (req, res) => {
    res.sendFile(path.join(__dirname, 'test', 'pkce-debug-interactive.html'));
});

// Handle PKCE callback
app.get('/src/auth/callback.html', (req, res) => {
    const { code, state, error, error_description } = req.query;
    
    if (error) {
        res.send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Authentication Error</title>
                <style>
                    body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #fff0f0; }
                    .error { color: #d13438; background: white; padding: 20px; border-radius: 8px; border: 1px solid #d13438; }
                </style>
            </head>
            <body>
                <div class="error">
                    <h2>Authentication Error</h2>
                    <p><strong>Error:</strong> ${error}</p>
                    <p><strong>Description:</strong> ${error_description || 'No description provided'}</p>
                    <p><a href="/pkce-debug">Return to PKCE Debugger</a></p>
                </div>
            </body>
            </html>
        `);
    } else {
        res.send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Authentication Success</title>
                <style>
                    body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #f0fff0; }
                    .success { color: #107c10; background: white; padding: 20px; border-radius: 8px; border: 1px solid #107c10; }
                    .callback-url { background: #f8f8f8; padding: 10px; border-radius: 4px; word-break: break-all; margin: 10px 0; }
                    button { background: #0078d4; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; }
                </style>
            </head>
            <body>
                <div class="success">
                    <h2>🎉 Authentication Successful!</h2>
                    <p><strong>Authorization Code:</strong> ${code ? code.substring(0, 20) + '...' : 'Not provided'}</p>
                    <p><strong>State:</strong> ${state || 'Not provided'}</p>
                    
                    <h3>📋 Copy this URL and paste it in the PKCE Debugger (Phase 3):</h3>
                    <div class="callback-url" id="callback-url">${req.url}</div>
                    
                    <button onclick="copyUrl()">📋 Copy URL</button>
                    <button onclick="window.opener && window.opener.focus(); window.close();">🔄 Return to Debugger</button>
                    
                    <p><a href="/pkce-debug" target="_blank">Open PKCE Debugger in new tab</a></p>
                </div>
                
                <script>
                    function copyUrl() {
                        const urlText = document.getElementById('callback-url').textContent;
                        navigator.clipboard.writeText('${req.protocol}://${req.get('host')}' + urlText).then(() => {
                            alert('URL copied to clipboard!');
                        });
                    }
                    
                    // Try to communicate with parent window if opened from debugger
                    if (window.opener && window.opener.location.href.includes('pkce-debug')) {
                        try {
                            window.opener.postMessage({
                                type: 'authCallback',
                                url: '${req.protocol}://${req.get('host')}' + '${req.url}',
                                code: '${code || ''}',
                                state: '${state || ''}',
                                error: '${error || ''}'
                            }, '*');
                        } catch (e) {
                            console.log('Could not communicate with parent window:', e);
                        }
                    }
                </script>
            </body>
            </html>
        `);
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        timestamp: new Date().toISOString(),
        message: 'PKCE Debug Server is running'
    });
});

// Start HTTP server for development
app.listen(PORT, () => {
    console.log(`🚀 PKCE Debug Server running at:`);
    console.log(`   - HTTP: http://localhost:${PORT}`);
    console.log(`   - PKCE Debugger: http://localhost:${PORT}/pkce-debug`);
    console.log(`   - Health Check: http://localhost:${PORT}/health`);
    console.log(`\n📋 To test PKCE flow:`);
    console.log(`   1. Open http://localhost:${PORT}/pkce-debug`);
    console.log(`   2. Follow the step-by-step debugging process`);
    console.log(`   3. Use the interactive phases to test each part of PKCE`);
    console.log(`\n🔧 Make sure your Azure AD app is configured with:`);
    console.log(`   - Redirect URI: https://localhost:3000/src/auth/callback.html`);
    console.log(`   - Allow public client flows: Yes`);
    console.log(`   - API permissions: User.Read, Notes.Read (or Notes.ReadWrite)`);
});
