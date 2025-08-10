/**
 * Microsoft Graph Authentication Service
 * 
 * Implements Office SSO-first authentication with MSAL.js fallback
 * Designed for Office Add-ins with personal Microsoft accounts
 * 
 * Authentication Flow:
 * 1. Try Office SSO (Office.context.auth.getAccessTokenAsync)
 * 2. Fallback to MSAL.js popup authentication
 * 3. Handle permission consent like Microsoft's "Save to OneNote" add-in
 */

import { PublicClientApplication, LogLevel } from '@azure/msal-browser';

// MSAL Configuration for Office Add-ins with personal accounts
const msalConfig = {
    auth: {
        clientId: 'a73f5240-e06c-43a3-8328-1fbd80766263',
        authority: 'https://login.microsoftonline.com/consumers', // Personal accounts only
        redirectUri: 'https://localhost:3000/src/auth/msal-callback.html' // MSAL callback page
    },
    cache: {
        cacheLocation: 'sessionStorage', // Use sessionStorage for security
        storeAuthStateInCookie: false
    },
    system: {
        allowNativeBroker: false, // Disable native broker for Office Add-ins
        windowHashTimeout: 60000, // Increase timeout for popup flow
        iframeHashTimeout: 6000,
        loadFrameTimeout: 0,
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return;
                console.log('[MSAL]', message);
            },
            logLevel: LogLevel.Warning
        }
    }
};

// Required Microsoft Graph scopes
const GRAPH_SCOPES = [
    'https://graph.microsoft.com/Mail.Read',
    'https://graph.microsoft.com/Notes.Read', 
    'https://graph.microsoft.com/Notes.ReadWrite',
    'https://graph.microsoft.com/User.Read'
];

class AuthService {
    constructor() {
        this.msalInstance = null;
        this.currentAccount = null;
        this.accessToken = null;
        this.initializeMsal();
    }

    /**
     * Initialize MSAL instance
     */
    initializeMsal() {
        try {
            this.msalInstance = new PublicClientApplication(msalConfig);
            console.log('✅ MSAL initialized successfully');
        } catch (error) {
            console.error('❌ MSAL initialization failed:', error);
        }
    }

    /**
     * Main authentication method - tries SSO first, then MSAL popup
     */
    async authenticate() {
        console.log('🔐 Starting authentication flow...');
        
        try {
            // Step 1: Try Office SSO first
            console.log('🏢 Attempting Office SSO...');
            const ssoToken = await this.tryOfficeSso();
            if (ssoToken) {
                console.log('✅ Office SSO successful');
                this.accessToken = ssoToken;
                return ssoToken;
            }
        } catch (ssoError) {
            console.log('⚠️ Office SSO failed:', ssoError.message);
        }

        try {
            // Step 2: Fallback to MSAL popup authentication
            console.log('🌐 Falling back to MSAL popup authentication...');
            const msalToken = await this.authenticateWithMsal();
            if (msalToken) {
                console.log('✅ MSAL authentication successful');
                this.accessToken = msalToken;
                return msalToken;
            }
        } catch (msalError) {
            console.error('❌ MSAL authentication failed:', msalError);
            throw new Error(`Authentication failed: ${msalError.message}`);
        }

        throw new Error('All authentication methods failed');
    }

    /**
     * Try Office SSO authentication
     */
    async tryOfficeSso() {
        return new Promise((resolve, reject) => {
            if (!Office?.context?.auth?.getAccessTokenAsync) {
                reject(new Error('Office SSO API not available'));
                return;
            }

            Office.context.auth.getAccessTokenAsync({
                allowConsentPrompt: true,
                allowSignInPrompt: true
            }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('✅ Office SSO token obtained');
                    resolve(result.value);
                } else {
                    reject(new Error(`Office SSO failed: ${result.error.message}`));
                }
            });
        });
    }

    /**
     * MSAL popup authentication
     */
    async authenticateWithMsal() {
        try {
            console.log('🔧 MSAL Config Check:', {
                clientId: msalConfig.auth.clientId,
                authority: msalConfig.auth.authority,
                redirectUri: msalConfig.auth.redirectUri
            });
            
            // Check if we already have an account
            const accounts = this.msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                this.currentAccount = accounts[0];
                console.log('📱 Using existing MSAL account:', this.currentAccount.username);
            }

            // Prepare login request
            const loginRequest = {
                scopes: GRAPH_SCOPES,
                prompt: 'select_account' // Allow user to select account
            };

            // Try silent authentication first
            if (this.currentAccount) {
                try {
                    console.log('🔄 Attempting silent token acquisition...');
                    const silentResult = await this.msalInstance.acquireTokenSilent({
                        ...loginRequest,
                        account: this.currentAccount
                    });
                    console.log('✅ Silent token acquisition successful');
                    this.currentAccount = silentResult.account;
                    return silentResult.accessToken;
                } catch (silentError) {
                    console.log('⚠️ Silent authentication failed, using popup:', silentError.message);
                }
            }

            // Use popup for interactive authentication
            console.log('🖱️ Opening authentication popup...');
            console.log('🔧 Login Request:', loginRequest);
            
            const popupResult = await this.msalInstance.acquireTokenPopup(loginRequest);
            
            console.log('✅ Popup authentication successful');
            this.currentAccount = popupResult.account;
            return popupResult.accessToken;

        } catch (error) {
            console.error('❌ MSAL authentication error:', error);
            
            // Provide specific error handling
            if (error.errorCode === 'user_cancelled') {
                throw new Error('Authentication cancelled by user');
            } else if (error.message && error.message.includes('redirect_uri')) {
                throw new Error('Redirect URI configuration error. Please check Azure AD app registration.');
            } else {
                throw error;
            }
        }
    }

    /**
     * Get current access token (from memory or refresh if needed)
     */
    async getAccessToken() {
        // Return cached token if available and valid
        if (this.accessToken) {
            return this.accessToken;
        }

        // Try to get a fresh token
        console.log('🔄 No cached token, attempting authentication...');
        return await this.authenticate();
    }

    /**
     * Check if user is authenticated
     */
    isAuthenticated() {
        return !!this.accessToken || !!this.currentAccount;
    }

    /**
     * Sign out user
     */
    async signOut() {
        try {
            if (this.msalInstance && this.currentAccount) {
                await this.msalInstance.logoutPopup({
                    account: this.currentAccount
                });
            }
            this.currentAccount = null;
            this.accessToken = null;
            console.log('✅ User signed out successfully');
        } catch (error) {
            console.error('❌ Sign out error:', error);
        }
    }

    /**
     * Get current user account info
     */
    getCurrentUser() {
        return this.currentAccount;
    }

    /**
     * Make authenticated Graph API call
     */
    async callGraphApi(endpoint, method = 'GET', body = null) {
        try {
            const token = await this.getAccessToken();
            if (!token) {
                throw new Error('No access token available');
            }

            const headers = {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json',
                'Content-Type': 'application/json'
            };

            const options = {
                method: method,
                headers: headers
            };

            if (body && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
                options.body = JSON.stringify(body);
            }

            const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, options);
            
            console.log(`🌐 Graph API Request: ${method} https://graph.microsoft.com/v1.0${endpoint}`);
            console.log('🔑 Request Headers:', JSON.stringify(headers, null, 2));
            
            if (!response.ok) {
                // Try to get the detailed error message from the response
                let errorDetails = `${response.status} ${response.statusText}`;
                try {
                    const errorResponse = await response.json();
                    if (errorResponse.error) {
                        errorDetails += `\nError Code: ${errorResponse.error.code}`;
                        errorDetails += `\nError Message: ${errorResponse.error.message}`;
                        if (errorResponse.error.details) {
                            errorDetails += `\nDetails: ${JSON.stringify(errorResponse.error.details)}`;
                        }
                    }
                } catch (jsonError) {
                    console.log('Could not parse error response as JSON');
                }
                
                console.error('📋 Graph API Error Details:', errorDetails);
                console.error('🔗 Request URL:', `https://graph.microsoft.com/v1.0${endpoint}`);
                throw new Error(`Graph API call failed: ${errorDetails}`);
            }

            return await response.json();
        } catch (error) {
            console.error('❌ Graph API call failed:', error);
            throw error;
        }
    }
}

// Create singleton instance
const authService = new AuthService();

export default authService;

// Named exports for convenience
export {
    authService,
    GRAPH_SCOPES
};
