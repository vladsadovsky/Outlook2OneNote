/* eslint-disable no-unused-vars */
/* global crypto, TextEncoder */

/**
 * Crypto Utilities for OAuth 2.0 PKCE Implementation
 * 
 * This module provides cryptographic functions required for OAuth 2.0
 * Authorization Code Flow with Proof Key for Code Exchange (PKCE).
 * 
 * PKCE enhances security by:
 * - Preventing authorization code interception attacks
 * - Eliminating need for client secrets in public clients
 * - Using dynamic code verifier/challenge pairs
 * 
 * References:
 * - RFC 7636: Proof Key for Code Exchange by OAuth Public Clients
 * - https://tools.ietf.org/html/rfc7636
 */

/**
 * Generates a cryptographically random code verifier for PKCE
 * 
 * Requirements (RFC 7636):
 * - Minimum length: 43 characters
 * - Maximum length: 128 characters  
 * - Character set: [A-Z] / [a-z] / [0-9] / "-" / "." / "_" / "~"
 * 
 * @returns {string} Base64url-encoded code verifier
 */
export function generateCodeVerifier() {
  // Generate 32 random bytes (256 bits)
  // This will result in a 43-character base64url string
  const randomBytes = new Uint8Array(32);
  
  if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
    // Browser environment
    crypto.getRandomValues(randomBytes);
  } else {
    // Fallback for environments without crypto.getRandomValues
    console.warn('crypto.getRandomValues not available, using Math.random fallback');
    for (let i = 0; i < randomBytes.length; i++) {
      randomBytes[i] = Math.floor(Math.random() * 256);
    }
  }
  
  // Convert to base64url encoding
  const codeVerifier = base64urlEncode(randomBytes);
  
  console.log(`Generated code verifier: ${codeVerifier} (length: ${codeVerifier.length})`);
  return codeVerifier;
}

/**
 * Generates a code challenge from a code verifier using SHA256
 * 
 * PKCE code challenge is created by:
 * 1. SHA256 hash of the code verifier
 * 2. Base64url encoding of the hash
 * 
 * @param {string} codeVerifier - The code verifier to hash
 * @returns {Promise<string>} Base64url-encoded code challenge
 */
export async function generateCodeChallenge(codeVerifier) {
  try {
    // Convert code verifier to bytes
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    
    // Create SHA256 hash
    let hash;
    if (typeof crypto !== 'undefined' && crypto.subtle) {
      // Browser environment with Web Crypto API
      hash = await crypto.subtle.digest('SHA-256', data);
    } else {
      // Fallback - use a simple hash function (not cryptographically secure)
      console.warn('Web Crypto API not available, using fallback hash');
      hash = await fallbackSha256(codeVerifier);
    }
    
    // Convert hash to base64url
    const codeChallenge = base64urlEncode(new Uint8Array(hash));
    
    console.log(`Generated code challenge: ${codeChallenge}`);
    return codeChallenge;
    
  } catch (error) {
    console.error('Error generating code challenge:', error);
    throw new Error(`Failed to generate code challenge: ${error.message}`);
  }
}

/**
 * Encodes data to base64url format (RFC 4648 Section 5)
 * 
 * Base64url encoding:
 * - Uses URL-safe alphabet: A-Z, a-z, 0-9, -, _
 * - Removes padding characters (=)
 * - Safe for use in URLs and form data
 * 
 * @param {Uint8Array} data - Data to encode
 * @returns {string} Base64url-encoded string
 */
function base64urlEncode(data) {
  // Convert Uint8Array to base64
  let base64 = '';
  
  if (typeof btoa !== 'undefined') {
    // Browser environment
    const binary = String.fromCharCode(...data);
    base64 = btoa(binary);
  } else {
    // Fallback for Node.js-like environments
    base64 = Buffer.from(data).toString('base64');
  }
  
  // Convert to base64url by replacing URL-unsafe characters
  return base64
    .replace(/\+/g, '-')  // Replace + with -
    .replace(/\//g, '_')  // Replace / with _
    .replace(/=/g, '');   // Remove padding
}

/**
 * Fallback SHA256 implementation for environments without Web Crypto API
 * This is NOT cryptographically secure and should only be used for testing
 * 
 * @param {string} message - Message to hash
 * @returns {Promise<ArrayBuffer>} Hash as ArrayBuffer
 */
async function fallbackSha256(message) {
  console.warn('Using insecure fallback SHA256 - this should not be used in production');
  
  // This is a simple hash for fallback only
  // In production, you should ensure Web Crypto API is available
  let hash = 0;
  for (let i = 0; i < message.length; i++) {
    const char = message.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  
  // Convert to ArrayBuffer-like structure
  const buffer = new ArrayBuffer(32); // 256 bits
  const view = new DataView(buffer);
  
  // Fill buffer with hash-derived data
  for (let i = 0; i < 8; i++) {
    view.setUint32(i * 4, hash + i, false);
  }
  
  return buffer;
}

/**
 * Validates a code verifier according to PKCE specifications
 * 
 * @param {string} codeVerifier - Code verifier to validate
 * @returns {boolean} True if valid, false otherwise
 */
export function validateCodeVerifier(codeVerifier) {
  if (!codeVerifier || typeof codeVerifier !== 'string') {
    return false;
  }
  
  // Check length requirements (43-128 characters)
  if (codeVerifier.length < 43 || codeVerifier.length > 128) {
    console.error(`Invalid code verifier length: ${codeVerifier.length} (must be 43-128)`);
    return false;
  }
  
  // Check character set: [A-Z] / [a-z] / [0-9] / "-" / "." / "_" / "~"
  const validChars = /^[A-Za-z0-9\-._~]+$/;
  if (!validChars.test(codeVerifier)) {
    console.error('Invalid code verifier characters (must be [A-Za-z0-9\\-._~])');
    return false;
  }
  
  return true;
}

/**
 * Generates a random state parameter for OAuth security
 * 
 * The state parameter prevents CSRF attacks by:
 * - Being included in authorization requests
 * - Being returned by the authorization server
 * - Being validated by the client
 * 
 * @returns {string} Random state parameter
 */
export function generateState() {
  const randomBytes = new Uint8Array(16);
  
  if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
    crypto.getRandomValues(randomBytes);
  } else {
    for (let i = 0; i < randomBytes.length; i++) {
      randomBytes[i] = Math.floor(Math.random() * 256);
    }
  }
  
  const state = base64urlEncode(randomBytes);
  console.log(`Generated state parameter: ${state}`);
  return state;
}

/**
 * Utility function to check if cryptographic APIs are available
 * 
 * @returns {object} Availability of crypto APIs
 */
export function checkCryptoSupport() {
  const support = {
    webCrypto: typeof crypto !== 'undefined' && typeof crypto.subtle !== 'undefined',
    getRandomValues: typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function',
    textEncoder: typeof TextEncoder !== 'undefined',
    btoa: typeof btoa !== 'undefined'
  };
  
  console.log('Crypto API support:', support);
  
  if (!support.webCrypto) {
    console.warn('Web Crypto API not available - PKCE security may be reduced');
  }
  
  if (!support.getRandomValues) {
    console.warn('crypto.getRandomValues not available - using Math.random fallback');
  }
  
  return support;
}
