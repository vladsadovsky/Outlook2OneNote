/*
 * PKCE Authentication Test Script
 * 
 * This script validates the new OAuth 2.0 PKCE authentication system
 * and ensures all components are working correctly.
 * 
 * Run this test to verify:
 * - Configuration validation
 * - Crypto utilities functionality
 * - PKCE authentication flow
 * - Token management
 * - Microsoft Graph API integration
 * 
 * Usage:
 * - Include this file in your HTML or run with Node.js
 * - Check browser console for detailed test results
 * - All tests should pass for production readiness
 */

// Import modules to test
import { generateCodeVerifier, generateCodeChallenge, generateState, validateCodeVerifier, checkCryptoSupport } from '../common/crypto-utils.js';
import { PKCEAuthenticator, pkceAuth } from '../src/auth/pkce-auth.js';
import { getConfig, validateConfig } from '../src/auth/auth-config.js';
import { authenticateAndGetNotebooks, checkPlatformSupport } from '../src/auth/graphapi-auth.js';

// Test results
const testResults = {
  passed: 0,
  failed: 0,
  errors: []
};

// Test utilities
function assert(condition, message) {
  if (condition) {
    console.log(`✅ ${message}`);
    testResults.passed++;
    return true;
  } else {
    console.error(`❌ ${message}`);
    testResults.failed++;
    testResults.errors.push(message);
    return false;
  }
}

function asyncTest(testName, testFn) {
  return async function() {
    try {
      console.log(`\n🔍 Running test: ${testName}`);
      await testFn();
      console.log(`✅ Test passed: ${testName}`);
    } catch (error) {
      console.error(`❌ Test failed: ${testName}`, error);
      testResults.failed++;
      testResults.errors.push(`${testName}: ${error.message}`);
    }
  };
}

// Test Suite
const tests = [
  asyncTest('Configuration Validation', async () => {
    // Test configuration loading
    const config = getConfig();
    assert(config.azureAd, 'Azure AD config loaded');
    assert(config.pkce, 'PKCE config loaded');
    assert(config.storage, 'Storage config loaded');
    assert(config.environment, 'Environment config loaded');
    
    // Test configuration validation
    try {
      validateConfig();
      console.log('⚠️ Configuration validation passed - make sure to update clientId with real Azure AD app ID');
    } catch (error) {
      if (error.message.includes('CLIENT_ID must be set')) {
        console.log('ℹ️ Configuration validation correctly detected placeholder client ID');
      } else {
        throw error;
      }
    }
  }),

  asyncTest('Crypto Utilities', async () => {
    // Test crypto support detection
    const cryptoSupport = checkCryptoSupport();
    assert(typeof cryptoSupport === 'object', 'Crypto support check returns object');
    console.log('Crypto support:', cryptoSupport);
    
    // Test code verifier generation
    const codeVerifier = generateCodeVerifier();
    assert(typeof codeVerifier === 'string', 'Code verifier is string');
    assert(codeVerifier.length >= 43, 'Code verifier meets minimum length requirement');
    assert(codeVerifier.length <= 128, 'Code verifier meets maximum length requirement');
    assert(validateCodeVerifier(codeVerifier), 'Generated code verifier is valid');
    
    // Test code challenge generation
    const codeChallenge = await generateCodeChallenge(codeVerifier);
    assert(typeof codeChallenge === 'string', 'Code challenge is string');
    assert(codeChallenge.length > 0, 'Code challenge is not empty');
    assert(codeChallenge !== codeVerifier, 'Code challenge differs from verifier');
    
    // Test state generation
    const state = generateState();
    assert(typeof state === 'string', 'State is string');
    assert(state.length > 0, 'State is not empty');
    
    // Test multiple generations produce different values
    const codeVerifier2 = generateCodeVerifier();
    const state2 = generateState();
    assert(codeVerifier !== codeVerifier2, 'Multiple code verifiers are unique');
    assert(state !== state2, 'Multiple states are unique');
  }),

  asyncTest('PKCE Authenticator Initialization', async () => {
    // Test authenticator creation
    const authenticator = new PKCEAuthenticator();
    assert(authenticator instanceof PKCEAuthenticator, 'Authenticator instance created');
    assert(typeof authenticator.config === 'object', 'Authenticator has config');
    assert(typeof authenticator.cryptoSupport === 'object', 'Authenticator has crypto support info');
    
    // Test configuration properties
    assert(typeof authenticator.config.clientId === 'string', 'Client ID is configured');
    assert(typeof authenticator.config.authority === 'string', 'Authority is configured');
    assert(Array.isArray(authenticator.config.scopes), 'Scopes are configured as array');
    assert(authenticator.config.scopes.length > 0, 'At least one scope configured');
    assert(typeof authenticator.config.redirectUri === 'string', 'Redirect URI is configured');
    
    // Test endpoint URLs
    assert(authenticator.tokenEndpoint.includes('/oauth2/v2.0/token'), 'Token endpoint is correct');
    assert(authenticator.authEndpoint.includes('/oauth2/v2.0/authorize'), 'Auth endpoint is correct');
  }),

  asyncTest('Platform Support Detection', async () => {
    // Test platform support check
    const platformSupport = checkPlatformSupport();
    assert(typeof platformSupport === 'object', 'Platform support returns object');
    assert(typeof platformSupport.supportsPKCE === 'boolean', 'PKCE support detected');
    assert(typeof platformSupport.hasOffice === 'boolean', 'Office.js detection works');
    
    console.log('Platform support details:', platformSupport);
  }),

  asyncTest('Token Storage Operations', async () => {
    const authenticator = new PKCEAuthenticator();
    
    // Test secure storage operations
    const testKey = 'test_key';
    const testValue = 'test_value_' + Date.now();
    
    // Store and retrieve
    authenticator.storeSecurely(testKey, testValue);
    const retrievedValue = authenticator.retrieveSecurely(testKey);
    assert(retrievedValue === testValue, 'Storage and retrieval works');
    
    // Remove
    authenticator.removeSecurely(testKey);
    const removedValue = authenticator.retrieveSecurely(testKey);
    assert(removedValue === null, 'Storage removal works');
  }),

  asyncTest('Authorization URL Building', async () => {
    const authenticator = new PKCEAuthenticator();
    
    // Generate PKCE parameters
    const codeVerifier = generateCodeVerifier();
    const codeChallenge = await generateCodeChallenge(codeVerifier);
    const state = generateState();
    
    // Build authorization URL
    const authUrl = authenticator.buildAuthorizationUrl(codeChallenge, state);
    
    assert(typeof authUrl === 'string', 'Authorization URL is string');
    assert(authUrl.startsWith('https://'), 'Authorization URL uses HTTPS');
    assert(authUrl.includes('client_id='), 'URL contains client ID');
    assert(authUrl.includes('code_challenge='), 'URL contains code challenge');
    assert(authUrl.includes('state='), 'URL contains state');
    assert(authUrl.includes('response_type=code'), 'URL contains correct response type');
    assert(authUrl.includes('code_challenge_method=S256'), 'URL contains correct challenge method');
    
    console.log('Sample authorization URL:', authUrl);
  }),

  asyncTest('Token Validation Logic', async () => {
    const authenticator = new PKCEAuthenticator();
    
    // Test with no tokens
    const hasValidToken1 = await authenticator.hasValidToken();
    assert(hasValidToken1 === false, 'Reports no valid token when none stored');
    
    // Test refresh token availability
    const canRefresh1 = await authenticator.canRefreshToken();
    assert(canRefresh1 === false, 'Reports cannot refresh when no refresh token');
    
    // Test with expired token
    const expiredTime = Date.now() - 3600000; // 1 hour ago
    authenticator.storeSecurely('pkce_access_token', 'fake_token');
    authenticator.storeSecurely('pkce_token_expires', expiredTime.toString());
    
    const hasValidToken2 = await authenticator.hasValidToken();
    assert(hasValidToken2 === false, 'Reports invalid token when expired');
    
    // Test with valid token
    const futureTime = Date.now() + 3600000; // 1 hour from now
    authenticator.storeSecurely('pkce_token_expires', futureTime.toString());
    
    const hasValidToken3 = await authenticator.hasValidToken();
    assert(hasValidToken3 === true, 'Reports valid token when not expired');
    
    // Cleanup test data
    authenticator.clearAuthData();
  }),

  asyncTest('Mock Data Generation', async () => {
    const authenticator = new PKCEAuthenticator();
    
    // Test mock notebook generation
    const mockNotebooks = authenticator.getMockNotebooks();
    assert(Array.isArray(mockNotebooks), 'Mock notebooks is array');
    assert(mockNotebooks.length > 0, 'Mock notebooks contains items');
    
    // Validate notebook structure
    const notebook = mockNotebooks[0];
    assert(typeof notebook.id === 'string', 'Mock notebook has ID');
    assert(typeof notebook.name === 'string', 'Mock notebook has name');
    assert(typeof notebook.displayName === 'string', 'Mock notebook has display name');
    assert(typeof notebook.isDefault === 'boolean', 'Mock notebook has isDefault flag');
    assert(typeof notebook.links === 'object', 'Mock notebook has links object');
  }),

  asyncTest('Authentication Integration Test', async () => {
    // Test the main authentication function
    // This will likely return mock data unless real Azure AD is configured
    try {
      const notebooks = await authenticateAndGetNotebooks();
      
      if (notebooks === null) {
        console.log('ℹ️ Authentication flow started (would redirect in real scenario)');
      } else if (Array.isArray(notebooks)) {
        assert(notebooks.length > 0, 'Authentication returned notebooks');
        console.log(`Retrieved ${notebooks.length} notebooks (likely mock data)`);
      } else {
        throw new Error('Unexpected return type from authentication');
      }
    } catch (error) {
      if (error.message.includes('CLIENT_ID must be set')) {
        console.log('ℹ️ Authentication correctly detected configuration requirement');
      } else {
        throw error;
      }
    }
  })
];

// Run all tests
async function runAllTests() {
  console.log('🚀 Starting PKCE Authentication Test Suite\n');
  console.log('=' .repeat(60));
  
  const startTime = Date.now();
  
  for (const test of tests) {
    await test();
  }
  
  const endTime = Date.now();
  const duration = endTime - startTime;
  
  console.log('\n' + '='.repeat(60));
  console.log('📊 Test Results Summary');
  console.log('='.repeat(60));
  console.log(`✅ Passed: ${testResults.passed}`);
  console.log(`❌ Failed: ${testResults.failed}`);
  console.log(`⏱️ Duration: ${duration}ms`);
  
  if (testResults.failed > 0) {
    console.log('\n❌ Failed Tests:');
    testResults.errors.forEach(error => console.log(`   • ${error}`));
  }
  
  if (testResults.failed === 0) {
    console.log('\n🎉 All tests passed! PKCE authentication system is ready.');
    console.log('\n📝 Next Steps:');
    console.log('   1. Update auth-config.js with your Azure AD client ID');
    console.log('   2. Configure redirect URI in Azure AD App Registration');
    console.log('   3. Grant necessary API permissions in Azure AD');
    console.log('   4. Test with real authentication flow');
  } else {
    console.log('\n🔧 Some tests failed. Please review the errors above.');
  }
  
  return testResults.failed === 0;
}

// Export for use in other contexts
export { runAllTests, testResults };

// Auto-run if this script is loaded directly
if (typeof document !== 'undefined') {
  document.addEventListener('DOMContentLoaded', () => {
    runAllTests().catch(error => {
      console.error('Test suite encountered an error:', error);
    });
  });
} else if (typeof process !== 'undefined' && process.env.NODE_ENV === 'test') {
  runAllTests().then(success => {
    process.exit(success ? 0 : 1);
  });
}
