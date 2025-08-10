/*
 * Azure AD Configuration for PKCE Authentication
 * 
 * This file contains the Azure Active Directory configuration required for
 * OAuth 2.0 Authorization Code Flow with PKCE (Proof Key for Code Exchange).
 * 
 * Setup Instructions:
 * 1. Register your application in Azure AD: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps
 * 2. Configure the following settings:
 *    - Platform: Single-page application (SPA)  
 *    - Redirect URIs: Add your callback URL(s)
 *    - API permissions: Microsoft Graph (Notes.Read, Notes.ReadWrite, User.Read)
 * 3. Copy the Application (client) ID to this configuration
 * 4. Update the redirect URI to match your hosting environment
 * 
 * Security Notes:
 * - No client secret is required for PKCE flow
 * - This configuration can be safely included in client-side code
 * - The client ID is not considered sensitive information
 * - All security is handled by PKCE code challenge/verifier pairs
 */

// Azure AD App Registration Configuration
export const AZURE_AD_CONFIG = {
  // Your Azure AD Application (client) ID
  // Replace this with your actual client ID from Azure AD App Registration
  clientId: ' a73f5240-e06c-43a3-8328-1fbd80766263',
  
  // Azure AD Authority URL
  // Use 'common' for multi-tenant, or your specific tenant ID for single tenant
  authority: 'https://login.microsoftonline.com/common',
  
  // Microsoft Graph API permissions required
  // These must be configured in your Azure AD App Registration
  scopes: [
    'https://graph.microsoft.com/Notes.Read',        // Read OneNote notebooks and sections
    'https://graph.microsoft.com/Notes.ReadWrite',   // Create and modify OneNote pages
    'https://graph.microsoft.com/User.Read',         // Read user profile information
    'profile',                                       // Basic profile information
    'openid',                                        // OpenID Connect sign-in
    'email'                                          // User's email address
  ],
  
  // OAuth 2.0 Redirect URIs
  // These must match exactly with what's configured in Azure AD
  redirectUri: getRedirectUri(),
  
  // Post-logout redirect URI (optional)
  postLogoutRedirectUri: getPostLogoutRedirectUri()
};

// PKCE Configuration
export const PKCE_CONFIG = {
  // Code challenge method (always S256 for security)
  codeChallengeMethod: 'S256',
  
  // Response type (always 'code' for authorization code flow)
  responseType: 'code',
  
  // Response mode (query or fragment)
  responseMode: 'query',
  
  // Prompt behavior
  // - 'select_account': Always show account selection
  // - 'login': Force fresh authentication
  // - 'consent': Force fresh consent
  // - 'none': Silent authentication only
  prompt: 'select_account'
};

// Token Storage Configuration
export const STORAGE_CONFIG = {
  // Use sessionStorage for better security (cleared when tab closes)
  useSessionStorage: true,
  
  // Token storage keys
  keys: {
    ACCESS_TOKEN: 'pkce_access_token',
    REFRESH_TOKEN: 'pkce_refresh_token',
    TOKEN_EXPIRES: 'pkce_token_expires',
    CODE_VERIFIER: 'pkce_code_verifier',
    STATE: 'pkce_oauth_state',
    USER_INFO: 'pkce_user_info'
  }
};

// Development vs Production Configuration
export const ENVIRONMENT_CONFIG = {
  // Set to true for development mode with additional logging
  isDevelopment: process.env.NODE_ENV === 'development',
  
  // Enable mock data fallback in development
  enableMockData: true,
  
  // API endpoints
  graphApiEndpoint: 'https://graph.microsoft.com/v1.0',
  
  // Timeout settings (in milliseconds)
  authTimeout: 30000,      // 30 seconds for auth operations
  apiTimeout: 10000        // 10 seconds for API calls
};

/**
 * Determine the appropriate redirect URI based on environment
 */
function getRedirectUri() {
  if (typeof window !== 'undefined') {
    // Browser environment - use current origin
    const origin = window.location.origin;
    return `${origin}/src/auth/callback.html`;
  } else {
    // Default for development
    return 'http://localhost:3000/src/auth/callback.html';
  }
}

/**
 * Determine the post-logout redirect URI
 */
function getPostLogoutRedirectUri() {
  if (typeof window !== 'undefined') {
    return window.location.origin;
  } else {
    return 'http://localhost:3000';
  }
}

/**
 * Validate the configuration at runtime
 */
export function validateConfig() {
  const errors = [];
  
  // Check required fields
  if (!AZURE_AD_CONFIG.clientId || AZURE_AD_CONFIG.clientId === '12345678-1234-1234-1234-123456789abc') {
    errors.push('CLIENT_ID must be set to your Azure AD Application ID');
  }
  
  if (!AZURE_AD_CONFIG.authority) {
    errors.push('AUTHORITY must be set to your Azure AD authority URL');
  }
  
  if (!AZURE_AD_CONFIG.scopes || AZURE_AD_CONFIG.scopes.length === 0) {
    errors.push('SCOPES must include at least one Microsoft Graph permission');
  }
  
  if (!AZURE_AD_CONFIG.redirectUri) {
    errors.push('REDIRECT_URI must be configured');
  }
  
  // Validate redirect URI format
  try {
    new URL(AZURE_AD_CONFIG.redirectUri);
  } catch (e) {
    errors.push('REDIRECT_URI must be a valid URL');
  }
  
  if (errors.length > 0) {
    console.error('❌ Azure AD Configuration Errors:', errors);
    throw new Error(`Configuration validation failed: ${errors.join(', ')}`);
  }
  
  console.log('✅ Azure AD configuration validated successfully');
  return true;
}

/**
 * Get configuration for current environment
 */
export function getConfig() {
  // Validate configuration
  validateConfig();
  
  return {
    azureAd: AZURE_AD_CONFIG,
    pkce: PKCE_CONFIG,
    storage: STORAGE_CONFIG,
    environment: ENVIRONMENT_CONFIG
  };
}

// Export individual configs for convenience
export { AZURE_AD_CONFIG as azureConfig };
export { PKCE_CONFIG as pkceConfig };
export { STORAGE_CONFIG as storageConfig };
export { ENVIRONMENT_CONFIG as envConfig };
