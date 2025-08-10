/* eslint-disable no-unused-vars */
/* global process */

/**
 * Environment Configuration Loader
 * 
 * This module loads configuration from environment variables and provides
 * fallback defaults for development. In production, all sensitive values
 * should be provided via environment variables.
 * 
 * Security Note: Client secrets should never be exposed in client-side code.
 * This configuration is intended for backend services only.
 */

/**
 * Load environment configuration
 * In browser environments, these would typically come from build-time injection
 * or server-side rendering context
 */
function loadEnvironmentConfig() {
  // Check if we're in a Node.js environment
  const isNode = typeof process !== 'undefined' && process.env;
  
  // For browser environments, these would be injected at build time
  const browserConfig = {
    clientId: 'a73f5240-e06c-43a3-8328-1fbd80766263',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'https://localhost:3000/src/auth/callback',
    postLogoutRedirectUri: 'https://localhost:3000',
    backendServiceUrl: 'https://your-backend-service.azurewebsites.net',
    scopes: ['https://graph.microsoft.com/Notes.Read', 'https://graph.microsoft.com/Notes.ReadWrite', 'https://graph.microsoft.com/Mail.Read', 'https://graph.microsoft.com/User.Read'],
    useSessionStorage: true,
    debugAuth: true
  };
  
  // Server-side configuration (Node.js with access to process.env)
  const serverConfig = isNode ? {
    clientId: process.env.CLIENT_ID || browserConfig.clientId,
    clientSecret: process.env.CLIENT_SECRET || '',
    tenantId: process.env.TENANT_ID || 'common',
    authority: process.env.AUTHORITY || browserConfig.authority,
    tokenEndpoint: process.env.TOKEN_ENDPOINT || `${browserConfig.authority}/oauth2/v2.0/token`,
    authEndpoint: process.env.AUTH_ENDPOINT || `${browserConfig.authority}/oauth2/v2.0/authorize`,
    redirectUri: process.env.REDIRECT_URI || browserConfig.redirectUri,
    postLogoutRedirectUri: process.env.POST_LOGOUT_REDIRECT_URI || browserConfig.postLogoutRedirectUri,
    backendServiceUrl: process.env.BACKEND_SERVICE_URL || browserConfig.backendServiceUrl,
    scopes: (process.env.GRAPH_SCOPES || '').split(' ').filter(Boolean) || browserConfig.scopes,
    useSessionStorage: process.env.USE_SESSION_STORAGE === 'true' || browserConfig.useSessionStorage,
    debugAuth: process.env.DEBUG_AUTH === 'true' || browserConfig.debugAuth,
    nodeEnv: process.env.NODE_ENV || 'development'
  } : {};
  
  // Merge configurations with server config taking precedence
  return { ...browserConfig, ...serverConfig };
}

/**
 * Get authentication configuration
 */
export function getAuthConfig() {
  const envConfig = loadEnvironmentConfig();
  
  return {
    azureAd: {
      clientId: envConfig.clientId,
      clientSecret: envConfig.clientSecret, // Only available server-side
      authority: envConfig.authority,
      redirectUri: envConfig.redirectUri,
      postLogoutRedirectUri: envConfig.postLogoutRedirectUri,
      scopes: Array.isArray(envConfig.scopes) ? envConfig.scopes : [envConfig.scopes],
      tenantId: envConfig.tenantId || 'common'
    },
    endpoints: {
      token: envConfig.tokenEndpoint || `${envConfig.authority}/oauth2/v2.0/token`,
      auth: envConfig.authEndpoint || `${envConfig.authority}/oauth2/v2.0/authorize`,
      backend: envConfig.backendServiceUrl
    },
    pkce: {
      responseType: 'code',
      codeChallengeMethod: 'S256',
      responseMode: 'query',
      prompt: 'select_account'
    },
    storage: {
      useSessionStorage: envConfig.useSessionStorage,
      keys: {
        ACCESS_TOKEN: 'outlook2onenote_access_token',
        REFRESH_TOKEN: 'outlook2onenote_refresh_token',
        TOKEN_EXPIRES: 'outlook2onenote_token_expires',
        CODE_VERIFIER: 'outlook2onenote_code_verifier',
        STATE: 'outlook2onenote_state',
        USER_INFO: 'outlook2onenote_user_info'
      }
    },
    debug: envConfig.debugAuth || false
  };
}

/**
 * Get backend service configuration (server-side only)
 * This is used by backend services that need the client secret
 */
export function getBackendConfig() {
  const envConfig = loadEnvironmentConfig();
  
  // Only return client secret if we're in a server environment
  if (typeof process === 'undefined' || !process.env) {
    console.warn('Backend config requested in browser environment - client secret not available');
    return null;
  }
  
  return {
    clientId: envConfig.clientId,
    clientSecret: envConfig.clientSecret,
    authority: envConfig.authority,
    tokenEndpoint: envConfig.tokenEndpoint,
    scopes: envConfig.scopes,
    tenantId: envConfig.tenantId
  };
}

/**
 * Validate that required environment variables are set
 */
export function validateEnvironmentConfig() {
  const config = getAuthConfig();
  const errors = [];
  
  if (!config.azureAd.clientId) {
    errors.push('CLIENT_ID is required');
  }
  
  if (!config.azureAd.authority) {
    errors.push('AUTHORITY is required');
  }
  
  if (!config.azureAd.redirectUri) {
    errors.push('REDIRECT_URI is required');
  }
  
  if (!config.azureAd.scopes || config.azureAd.scopes.length === 0) {
    errors.push('At least one scope is required');
  }
  
  // For backend services, client secret is required
  const backendConfig = getBackendConfig();
  if (backendConfig && !backendConfig.clientSecret) {
    console.warn('CLIENT_SECRET not set - backend token exchange will not work');
  }
  
  if (errors.length > 0) {
    throw new Error(`Environment configuration errors: ${errors.join(', ')}`);
  }
  
  if (config.debug) {
    console.log('Environment configuration validated:', {
      clientId: config.azureAd.clientId,
      authority: config.azureAd.authority,
      redirectUri: config.azureAd.redirectUri,
      scopes: config.azureAd.scopes,
      hasClientSecret: !!backendConfig?.clientSecret
    });
  }
  
  return true;
}

// Legacy export for compatibility
export function getConfig() {
  return getAuthConfig();
}
