/* eslint-disable no-unused-vars */
/* global Office, window, document, localStorage, sessionStorage */

/**
 * OAuth 2.0 Authorization Code Flow with PKCE Authentication Service
 * 
 * This module implements secure authentication for Microsoft Graph API access
 * using OAuth 2.0 Authorization Code Flow with Proof Key for Code Exchange (PKCE).
 * 
 * Key Features:
 * - Client-side authentication without backend requirements
 * - PKCE for enhanced security (no client secret needed)
 * - SSO fallback for Office Add-ins
 * - Token management with automatic refresh
 * - Cross-platform support (desktop, web, mobile)
 * 
 * Security Benefits:
 * - Prevents authorization code interception attacks
 * - Eliminates client secret exposure risks
 * - Uses dynamic code verifier/challenge pairs
 * - Implements proper token storage and rotation
 * 
 * Dependencies:
 * - crypto-utils.js for PKCE cryptographic functions
 * - Office.js for SSO fallback
 */

import { 
  generateCodeVerifier, 
  generateCodeChallenge, 
  generateState, 
  validateCodeVerifier,
  checkCryptoSupport 
} from '../common/crypto-utils.js';

import { getAuthConfig, validateEnvironmentConfig } from '../common/env-config.js';

// Get configuration from environment
const config = getAuthConfig();
const AUTH_CONFIG = config.azureAd;
const PKCE_CONFIG = config.pkce;
const STORAGE_CONFIG = config.storage;
const ENDPOINTS = config.endpoints;

// Storage keys for authentication data
const STORAGE_KEYS = STORAGE_CONFIG.keys;

/**
 * Main authentication class implementing PKCE OAuth flow
 */
export class PKCEAuthenticator {
  constructor(customConfig = {}) {
    // Validate environment configuration on initialization
    try {
      validateEnvironmentConfig();
    } catch (error) {
      console.error('Environment configuration validation failed:', error);
      throw error;
    }
    
    // Merge custom configuration with defaults from environment
    this.config = { ...AUTH_CONFIG, ...customConfig };
    this.tokenEndpoint = ENDPOINTS.token;
    this.authEndpoint = ENDPOINTS.auth;
    this.backendEndpoint = ENDPOINTS.backend;
    
    // Check crypto support on initialization
    this.cryptoSupport = checkCryptoSupport();
    
    console.log('PKCE Authenticator initialized with config:', {
      clientId: this.config.clientId,
      authority: this.config.authority,
      scopes: this.config.scopes,
      cryptoSupport: this.cryptoSupport,
      hasBackendEndpoint: !!this.backendEndpoint
    });
  }

  /**
   * Main authentication entry point with fallback strategies
   * 
   * Authentication Priority:
   * 1. Existing valid tokens (silent authentication)
   * 2. PKCE Authorization Code Flow (popup-based)
   * 3. SSO fallback (Office Add-ins)
   * 4. Mock data (development/testing)
   */
  async authenticateAndGetNotebooks() {
    console.log('Starting PKCE authentication flow...');
    
    try {
      // Check for existing valid tokens first
      if (await this.hasValidToken()) {
        console.log('✅ Using existing valid access token');
        return await this.getNotebooks();
      }
      
      // Try to refresh token if available
      if (await this.canRefreshToken()) {
        console.log('🔄 Refreshing expired access token');
        await this.refreshAccessToken();
        return await this.getNotebooks();
      }
      
      // Start PKCE authorization flow (popup-based)
      console.log('🚀 Starting PKCE authorization code flow');
      const notebooks = await this.startPKCEFlow();
      
      if (notebooks) {
        console.log('✅ PKCE authentication completed successfully');
        return notebooks;
      }
      
      // If no notebooks returned, try to get them
      return await this.getNotebooks();
      
    } catch (error) {
      console.error('PKCE authentication failed:', error);
      
      // Fallback to SSO if available
      try {
        console.log('🔄 Falling back to SSO authentication');
        return await this.trySSoFallback();
      } catch (ssoError) {
        console.error('SSO fallback also failed:', ssoError);
        
        // Final fallback to mock data
        console.log('📱 Using mock data for development');
        return this.getMockNotebooks();
      }
    }
  }  /**
   * Starts the PKCE authorization code flow using popup window
   * This is better for Office Add-ins as it doesn't navigate away from the taskpane
   */
  async startPKCEFlow() {
    try {
      console.log('🚀 Starting PKCE flow with popup...');
      
      // Generate PKCE parameters
      const codeVerifier = generateCodeVerifier();
      const codeChallenge = await generateCodeChallenge(codeVerifier);
      const state = generateState();
      
      // Validate code verifier
      if (!validateCodeVerifier(codeVerifier)) {
        throw new Error('Generated code verifier is invalid');
      }
      
      // Store PKCE parameters securely
      this.storeSecurely(STORAGE_KEYS.CODE_VERIFIER, codeVerifier);
      this.storeSecurely(STORAGE_KEYS.STATE, state);
      
      // Build authorization URL
      const authUrl = this.buildAuthorizationUrl(codeChallenge, state);
      
      console.log('🔗 Opening authorization popup:', authUrl);
      console.log('🔧 Popup dimensions: 600x700');
      
      // Use popup for Office Add-in environment
      return new Promise((resolve, reject) => {
        const popup = window.open(
          authUrl,
          'pkce-auth-popup',
          'width=600,height=700,scrollbars=yes,resizable=yes,location=yes,status=yes,menubar=no,toolbar=no'
        );
        
        if (!popup) {
          console.error('❌ Failed to open popup - likely blocked by browser');
          reject(new Error('Failed to open authentication popup. Please allow popups for this site and try again.'));
          return;
        }
        
        console.log('✅ Popup opened successfully');
        
        // Listen for messages from the popup
        const messageHandler = async (event) => {
          console.log('📨 Received message from popup:', event.data, 'Origin:', event.origin);
          
          if (event.origin !== window.location.origin) {
            console.warn('⚠️ Ignoring message from different origin:', event.origin);
            return;
          }
          
          if (event.data.type === 'PKCE_AUTH_CODE') {
            console.log('🔑 Received authorization code from popup, processing in main window');
            
            try {
              // Validate state parameter in main window context
              const storedState = this.retrieveSecurely(STORAGE_KEYS.STATE);
              if (event.data.state !== storedState) {
                console.error('State mismatch:', { received: event.data.state, stored: storedState });
                throw new Error('Invalid state parameter - possible CSRF attack');
              }
              
              // Exchange code for tokens in main window (has access to code verifier)
              const tokens = await this.exchangeCodeForTokens(event.data.code);
              
              // Store tokens in main window
              await this.storeTokens(tokens);
              
              console.log('✅ Token exchange completed in main window');
              
              // Get notebooks
              const notebooks = await this.getNotebooks();
              
              // Clean up auth state
              this.cleanupAuthState();
              
              window.removeEventListener('message', messageHandler);
              popup.close();
              resolve(notebooks);
              
            } catch (error) {
              console.error('❌ Token exchange failed in main window:', error);
              window.removeEventListener('message', messageHandler);
              popup.close();
              reject(error);
            }
            
          } else if (event.data.type === 'PKCE_AUTH_SUCCESS') {
            console.log('✅ Authentication successful via popup');
            
            // Store the access token in the main window's session storage
            if (event.data.accessToken) {
              console.log('🔑 Storing access token from popup in main window');
              this.storeSecurely(STORAGE_KEYS.ACCESS_TOKEN, event.data.accessToken);
            }
            
            window.removeEventListener('message', messageHandler);
            popup.close();
            resolve(event.data.notebooks);
          } else if (event.data.type === 'PKCE_AUTH_ERROR') {
            console.error('❌ Authentication error via popup:', event.data.error);
            window.removeEventListener('message', messageHandler);
            popup.close();
            reject(new Error(event.data.error));
          }
        };
        
        window.addEventListener('message', messageHandler);
        
        // Check if popup was closed by user
        const checkClosed = setInterval(() => {
          if (popup.closed) {
            console.log('⚠️ Popup was closed by user');
            clearInterval(checkClosed);
            window.removeEventListener('message', messageHandler);
            reject(new Error('Authentication cancelled by user - popup was closed'));
          }
        }, 1000);
      });
      
    } catch (error) {
      console.error('Failed to start PKCE flow:', error);
      throw new Error(`PKCE flow initialization failed: ${error.message}`);
    }
  }

  /**
   * Builds the OAuth 2.0 authorization URL with PKCE parameters
   */
  buildAuthorizationUrl(codeChallenge, state) {
    const params = new URLSearchParams({
      client_id: this.config.clientId,
      response_type: PKCE_CONFIG.responseType,
      redirect_uri: this.config.redirectUri,
      scope: this.config.scopes.join(' '),
      state: state,
      code_challenge: codeChallenge,
      code_challenge_method: PKCE_CONFIG.codeChallengeMethod,
      response_mode: PKCE_CONFIG.responseMode,
      prompt: PKCE_CONFIG.prompt
    });
    
    return `${this.authEndpoint}?${params.toString()}`;
  }

  /**
   * Handles the authorization callback after user consent
   * Exchanges authorization code for access token
   */
  async handleAuthorizationCallback() {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const code = urlParams.get('code');
      const state = urlParams.get('state');
      const error = urlParams.get('error');
      const errorDescription = urlParams.get('error_description');
      
      // Handle authorization errors
      if (error) {
        console.error('Authorization error:', error, errorDescription);
        throw new Error(`Authorization failed: ${error} - ${errorDescription}`);
      }
      
      // Validate required parameters
      if (!code) {
        throw new Error('Authorization code not received');
      }
      
      if (!state) {
        throw new Error('State parameter not received');
      }
      
      // Validate state parameter (CSRF protection)
      const storedState = this.retrieveSecurely(STORAGE_KEYS.STATE);
      if (state !== storedState) {
        console.error('State mismatch:', { received: state, stored: storedState });
        throw new Error('Invalid state parameter - possible CSRF attack');
      }
      
      // Exchange authorization code for tokens
      const tokens = await this.exchangeCodeForTokens(code);
      
      // Store tokens securely
      await this.storeTokens(tokens);
      
      console.log('✅ PKCE authentication completed successfully');
      
      // Clean up authorization state
      this.cleanupAuthState();
      
      // Get and return notebooks
      return await this.getNotebooks();
      
    } catch (error) {
      console.error('Authorization callback handling failed:', error);
      this.cleanupAuthState();
      throw error;
    }
  }

  /**
   * Exchanges authorization code for access and refresh tokens
   * Supports both PKCE (client-side) and client secret (backend) flows
   */
  async exchangeCodeForTokens(authorizationCode) {
    try {
      const codeVerifier = this.retrieveSecurely(STORAGE_KEYS.CODE_VERIFIER);
      
      if (!codeVerifier) {
        throw new Error('Code verifier not found - PKCE flow was not properly initialized');
      }
      
      // Try backend token exchange first (more secure)
      if (this.backendEndpoint) {
        try {
          return await this.exchangeCodeViaBackend(authorizationCode, codeVerifier);
        } catch (backendError) {
          console.warn('Backend token exchange failed, falling back to client-side PKCE:', backendError);
        }
      }
      
      // Fallback to client-side PKCE token exchange
      return await this.exchangeCodeViaPKCE(authorizationCode, codeVerifier);
      
    } catch (error) {
      console.error('Token exchange failed:', error);
      throw new Error(`Failed to exchange authorization code: ${error.message}`);
    }
  }

  /**
   * Exchange authorization code via backend service (recommended)
   * Backend service uses client secret for enhanced security
   */
  async exchangeCodeViaBackend(authorizationCode, codeVerifier) {
    console.log('🔄 Exchanging authorization code via backend service');
    
    const backendRequest = {
      code: authorizationCode,
      codeVerifier: codeVerifier,
      redirectUri: this.config.redirectUri
    };
    
    const response = await fetch(`${this.backendEndpoint}/api/auth/exchange-code`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify(backendRequest)
    });
    
    const responseData = await response.json();
    
    if (!response.ok) {
      console.error('Backend token exchange failed:', responseData);
      throw new Error(`Backend token exchange failed: ${responseData.error} - ${responseData.error_description}`);
    }
    
    // Validate token response
    if (!responseData.access_token) {
      throw new Error('Access token not received from backend');
    }
    
    console.log('✅ Tokens received from backend service');
    
    return {
      accessToken: responseData.access_token,
      refreshToken: responseData.refresh_token,
      expiresIn: responseData.expires_in,
      tokenType: responseData.token_type,
      scope: responseData.scope
    };
  }

  /**
   * Exchange authorization code via client-side PKCE (fallback)
   * Uses PKCE without client secret
   */
  async exchangeCodeViaPKCE(authorizationCode, codeVerifier) {
    console.log('🔄 Exchanging authorization code via client-side PKCE');
    
    // Prepare token exchange request
    const tokenRequest = {
      client_id: this.config.clientId,
      code: authorizationCode,
      redirect_uri: this.config.redirectUri,
      grant_type: 'authorization_code',
      code_verifier: codeVerifier
    };
    
    const response = await fetch(this.tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
      },
      body: new URLSearchParams(tokenRequest)
    });
    
    const responseData = await response.json();
    
    if (!response.ok) {
      console.error('PKCE token exchange failed:', responseData);
      throw new Error(`PKCE token exchange failed: ${responseData.error} - ${responseData.error_description}`);
    }
    
    // Validate token response
    if (!responseData.access_token) {
      throw new Error('Access token not received in PKCE response');
    }
    
    console.log('✅ Tokens received via PKCE flow');
    
    return {
      accessToken: responseData.access_token,
      refreshToken: responseData.refresh_token,
      expiresIn: responseData.expires_in,
      tokenType: responseData.token_type,
      scope: responseData.scope
    };
  }

  /**
   * Refreshes an expired access token using refresh token
   */
  async refreshAccessToken() {
    try {
      const refreshToken = this.retrieveSecurely(STORAGE_KEYS.REFRESH_TOKEN);
      
      if (!refreshToken) {
        throw new Error('Refresh token not available');
      }
      
      const refreshRequest = {
        client_id: this.config.clientId,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope: this.config.scopes.join(' ')
      };
      
      console.log('🔄 Refreshing access token');
      
      const response = await fetch(this.tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Accept': 'application/json'
        },
        body: new URLSearchParams(refreshRequest)
      });
      
      const responseData = await response.json();
      
      if (!response.ok) {
        console.error('Token refresh failed:', responseData);
        
        // If refresh token is invalid, clear all auth data
        if (responseData.error === 'invalid_grant') {
          console.log('Refresh token expired, clearing auth data');
          this.clearAuthData();
        }
        
        throw new Error(`Token refresh failed: ${responseData.error} - ${responseData.error_description}`);
      }
      
      const tokens = {
        accessToken: responseData.access_token,
        refreshToken: responseData.refresh_token || refreshToken, // Use new refresh token if provided
        expiresIn: responseData.expires_in,
        tokenType: responseData.token_type,
        scope: responseData.scope
      };
      
      await this.storeTokens(tokens);
      
      console.log('✅ Access token refreshed successfully');
      
    } catch (error) {
      console.error('Token refresh failed:', error);
      throw error;
    }
  }

  /**
   * Makes authenticated requests to Microsoft Graph API
   */
  async getNotebooks() {
    try {
      const accessToken = this.retrieveSecurely(STORAGE_KEYS.ACCESS_TOKEN);
      
      if (!accessToken) {
        throw new Error('Access token not available');
      }
      
      console.log('📚 Fetching OneNote notebooks from Microsoft Graph');
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      });
      
      if (!response.ok) {
        const errorData = await response.text();
        console.error('Graph API request failed:', response.status, errorData);
        
        // Handle token expiration
        if (response.status === 401) {
          console.log('🔄 Token expired, attempting refresh');
          await this.refreshAccessToken();
          // Retry request with new token
          return await this.getNotebooks();
        }
        
        throw new Error(`Graph API request failed: ${response.status} ${response.statusText}`);
      }
      
      const data = await response.json();
      
      if (data.value && data.value.length > 0) {
        const notebooks = data.value.map(notebook => ({
          id: notebook.id,
          name: notebook.displayName,
          displayName: notebook.displayName,
          createdDateTime: notebook.createdDateTime,
          lastModifiedDateTime: notebook.lastModifiedDateTime,
          isDefault: notebook.isDefault || false,
          sectionsUrl: notebook.sectionsUrl,
          sectionGroupsUrl: notebook.sectionGroupsUrl,
          links: notebook.links
        }));
        
        console.log(`✅ Successfully retrieved ${notebooks.length} OneNote notebooks`);
        return notebooks;
      } else {
        console.log('📭 No OneNote notebooks found');
        return [];
      }
      
    } catch (error) {
      console.error('Failed to get notebooks:', error);
      throw error;
    }
  }

  /**
   * SSO fallback for Office Add-ins environment
   */
  async trySSoFallback() {
    return new Promise((resolve, reject) => {
      try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.auth) {
          throw new Error('Office.js SSO not available');
        }
        
        console.log('🔄 Attempting SSO fallback');
        
        Office.context.auth.getAccessTokenAsync({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true
        }, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            try {
              // Store SSO token temporarily
              const ssoToken = result.value;
              
              // For SSO tokens to work with Graph API, you typically need a backend
              // that can exchange the SSO token for a Graph API token
              // For now, we'll return mock data
              console.log('⚠️ SSO token received but backend exchange not implemented');
              resolve(this.getMockNotebooks());
              
            } catch (error) {
              console.error('SSO token processing failed:', error);
              reject(error);
            }
          } else {
            console.error('SSO authentication failed:', result.error);
            reject(new Error(`SSO failed: ${result.error?.message || 'Unknown error'}`));
          }
        });
        
      } catch (error) {
        console.error('SSO fallback failed:', error);
        reject(error);
      }
    });
  }

  /**
   * Token validation and management
   */
  async hasValidToken() {
    const accessToken = this.retrieveSecurely(STORAGE_KEYS.ACCESS_TOKEN);
    const expiresAt = this.retrieveSecurely(STORAGE_KEYS.TOKEN_EXPIRES);
    
    if (!accessToken || !expiresAt) {
      return false;
    }
    
    const now = Date.now();
    const expirationTime = parseInt(expiresAt, 10);
    
    // Add 5 minute buffer before expiration
    const bufferTime = 5 * 60 * 1000;
    
    return now < (expirationTime - bufferTime);
  }

  async canRefreshToken() {
    const refreshToken = this.retrieveSecurely(STORAGE_KEYS.REFRESH_TOKEN);
    return !!refreshToken;
  }

  /**
   * Get current access token
   * @returns {string|null} The access token or null if not available
   */
  getAccessToken() {
    return this.retrieveSecurely(STORAGE_KEYS.ACCESS_TOKEN);
  }

  /**
   * Secure token storage
   */
  async storeTokens(tokens) {
    const expiresAt = Date.now() + (tokens.expiresIn * 1000);
    
    this.storeSecurely(STORAGE_KEYS.ACCESS_TOKEN, tokens.accessToken);
    this.storeSecurely(STORAGE_KEYS.TOKEN_EXPIRES, expiresAt.toString());
    
    if (tokens.refreshToken) {
      this.storeSecurely(STORAGE_KEYS.REFRESH_TOKEN, tokens.refreshToken);
    }
    
    console.log(`✅ Tokens stored securely, expires at: ${new Date(expiresAt).toISOString()}`);
  }

  /**
   * Secure storage operations
   */
  storeSecurely(key, value) {
    try {
      // Use configured storage preference
      if (STORAGE_CONFIG.useSessionStorage && typeof sessionStorage !== 'undefined') {
        sessionStorage.setItem(key, value);
      } else if (typeof localStorage !== 'undefined') {
        localStorage.setItem(key, value);
      } else {
        console.warn('No secure storage available');
      }
    } catch (error) {
      console.error('Failed to store data securely:', error);
    }
  }

  retrieveSecurely(key) {
    try {
      if (STORAGE_CONFIG.useSessionStorage && typeof sessionStorage !== 'undefined') {
        return sessionStorage.getItem(key);
      } else if (typeof localStorage !== 'undefined') {
        return localStorage.getItem(key);
      }
      return null;
    } catch (error) {
      console.error('Failed to retrieve data securely:', error);
      return null;
    }
  }

  /**
   * Cleanup operations
   */
  cleanupAuthState() {
    this.removeSecurely(STORAGE_KEYS.CODE_VERIFIER);
    this.removeSecurely(STORAGE_KEYS.STATE);
  }

  clearAuthData() {
    Object.values(STORAGE_KEYS).forEach(key => {
      this.removeSecurely(key);
    });
    console.log('🧹 Authentication data cleared');
  }

  removeSecurely(key) {
    try {
      if (STORAGE_CONFIG.useSessionStorage && typeof sessionStorage !== 'undefined') {
        sessionStorage.removeItem(key);
      }
      if (typeof localStorage !== 'undefined') {
        localStorage.removeItem(key);
      }
    } catch (error) {
      console.error('Failed to remove data:', error);
    }
  }

  /**
   * Logout functionality
   */
  async logout() {
    try {
      // Clear local auth data
      this.clearAuthData();
      
      // Build logout URL
      const logoutUrl = `${this.config.authority}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(this.config.postLogoutRedirectUri)}`;
      
      console.log('👋 Logging out and redirecting to:', logoutUrl);
      
      // Redirect to logout endpoint
      window.location.href = logoutUrl;
      
    } catch (error) {
      console.error('Logout failed:', error);
      throw error;
    }
  }

  /**
   * Mock data for development and testing
   */
  getMockNotebooks() {
    console.log('🔧 Generating mock OneNote notebooks (PKCE fallback)');
    
    return [
      {
        id: 'pkce-mock-1',
        name: 'PKCE Test Notebook',
        displayName: 'PKCE Test Notebook',
        isDefault: true,
        createdDateTime: new Date(Date.now() - 86400000 * 7).toISOString(),
        lastModifiedDateTime: new Date().toISOString(),
        sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/pkce-mock-1/sections',
        sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/pkce-mock-1/sectionGroups',
        links: {
          oneNoteClientUrl: { href: 'onenote:https://mock-pkce/test-notebook' },
          oneNoteWebUrl: { href: 'https://mock-onenote-web/test-notebook' }
        }
      },
      {
        id: 'pkce-mock-2',
        name: 'Secure Authentication Demo',
        displayName: 'Secure Authentication Demo',
        isDefault: false,
        createdDateTime: new Date(Date.now() - 86400000 * 14).toISOString(),
        lastModifiedDateTime: new Date(Date.now() - 86400000 * 2).toISOString(),
        sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/pkce-mock-2/sections',
        sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/pkce-mock-2/sectionGroups',
        links: {
          oneNoteClientUrl: { href: 'onenote:https://mock-pkce/auth-demo' },
          oneNoteWebUrl: { href: 'https://mock-onenote-web/auth-demo' }
        }
      }
    ];
  }
}

// Export a singleton instance for easy use
export const pkceAuth = new PKCEAuthenticator();

// Legacy compatibility function
export async function authenticateAndGetNotebooks() {
  return await pkceAuth.authenticateAndGetNotebooks();
}
