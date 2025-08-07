/* eslint-disable no-unused-vars */
/* global Office */
/* eslint-disable no-console */

/*
 * Microsoft Graph API Authentication Service
 * 
 * This module handles authentication for Microsoft Graph API access in Office Add-ins.
 * Implements multiple authentication methods with fallback strategies to handle 
 * platform-specific limitations and error scenarios.
 * 
 * Error 13012 Resolution:
 * - "API is not supported in this platform" occurs when getAccessTokenAsync is not available
 * - This module provides fallback authentication methods for different Office platforms
 * 
 * Authentication Methods (in order of priority):
 * 1. SSO Authentication (getAccessTokenAsync) - most reliable when available
 * 2. REST API Token (getCallbackTokenAsync with isRest: true)
 * 3. Exchange Token (getCallbackTokenAsync with isRest: false)
 * 4. Mock data (for testing/development)
 * 
 * Dependencies:
 * - Office.js
 * - Microsoft Graph API
 */

// Check platform capabilities and available authentication methods
export function checkPlatformSupport() {
  const support = {
    hasOffice: typeof Office !== 'undefined',
    hasAuth: false,
    hasMailbox: false,
    platform: 'unknown',
    version: 'unknown'
  };
  
  try {
    if (Office && Office.context) {
      support.hasAuth = !!(Office.context.auth && Office.context.auth.getAccessTokenAsync);
      support.hasMailbox = !!(Office.context.mailbox && Office.context.mailbox.getCallbackTokenAsync);
      
      if (Office.context.host) {
        support.version = Office.context.host.version || 'unknown';
      }
    }
    
    // Determine platform
    if (Office.context.platform) {
      support.platform = Office.context.platform.toString();
    }
    
    console.log("Platform capabilities:", support);
    
  } catch (error) {
    console.error("Error checking platform support:", error);
  }
  
  return support;
}

// Main authentication function with multi-method fallback
export async function authenticateAndGetNotebooks() {
  console.log("Starting Microsoft Graph API authentication for OneNote notebooks...");
  
  const platformSupport = checkPlatformSupport();
  
  if (!platformSupport.hasOffice) {
    console.warn("Office.js not available - using mock data");
    return getMockNotebooks();
  }
  
  // Try authentication methods in order of preference
  const authMethods = [
    { name: 'SSO', method: trySSoAuthentication, available: platformSupport.hasAuth },
    { name: 'REST', method: tryRestApiToken, available: platformSupport.hasMailbox },
    { name: 'Exchange', method: tryExchangeToken, available: platformSupport.hasMailbox }
  ];
  
  for (const authMethod of authMethods) {
    if (!authMethod.available) {
      console.log(`${authMethod.name} authentication not available on this platform`);
      continue;
    }
    
    try {
      console.log(`Trying ${authMethod.name} authentication...`);
      const notebooks = await authMethod.method();
      
      if (notebooks && notebooks.length > 0) {
        console.log(`✅ Successfully authenticated using ${authMethod.name} method`);
        return notebooks;
      } else {
        console.log(`${authMethod.name} authentication returned no notebooks`);
      }
    } catch (error) {
      console.error(`${authMethod.name} authentication failed:`, error);
      
      // For the specific 401 InvalidAuthenticationToken error, provide explanation
      if (error.message && error.message.includes('InvalidAuthenticationToken')) {
        console.log(`📝 Note: ${authMethod.name} token has wrong audience for Graph API - this is expected`);
      }
    }
  }
  
  // If all methods fail, return mock data
  console.log("⚠️ All authentication methods failed - using mock data for testing");
  return getMockNotebooks();
}

// Method 1: Try SSO Authentication with Backend Token Exchange
export async function trySSoAuthentication() {
  return new Promise((resolve, reject) => {
    try {
      // Add timeout to prevent hanging
      const timeout = setTimeout(() => {
        console.log("SSO authentication timeout");
        reject(new Error("SSO authentication timeout"));
      }, 10000);
      
      Office.context.auth.getAccessTokenAsync({ 
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      }, async (tokenResult) => {
        clearTimeout(timeout);
        
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          try {
            console.log("SSO token obtained, calling backend for token exchange...");
            
            // Call your backend service to exchange SSO token for Graph API token
            const backendUrl = 'https://your-backend-service.azurewebsites.net/api/auth/exchange-token'; // Replace with your backend URL
            
            const backendResponse = await fetch(backendUrl, {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify({ ssoToken: tokenResult.value })
            });
            
            if (!backendResponse.ok) {
              throw new Error(`Backend token exchange failed: ${backendResponse.status}`);
            }
            
            const notebooks = await backendResponse.json();
            
            if (notebooks.value && notebooks.value.length > 0) {
              // Map the response to a consistent format
              const formattedNotebooks = notebooks.value.map(notebook => ({
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
              
              console.log(`✅ Successfully got ${formattedNotebooks.length} real OneNote notebooks via backend`);
              resolve(formattedNotebooks);
            } else {
              console.log("No OneNote notebooks found via backend");
              resolve([]);
            }
            
          } catch (backendError) {
            console.error("Backend token exchange failed:", backendError);
            console.log("📝 To get real notebooks, you need to implement a backend token exchange service");
            reject(backendError);
          }
        } else {
          console.error("SSO authentication failed:", {
            error: tokenResult.error,
            errorCode: tokenResult.error?.code,
            errorMessage: tokenResult.error?.message
          });
          reject(new Error(`SSO failed: ${tokenResult.error?.message || 'Unknown error'}`));
        }
      });
      
    } catch (error) {
      console.error("Exception in SSO authentication:", error);
      reject(error);
    }
  });
}

// Method 2: Try REST API Token
export async function tryRestApiToken() {
  return new Promise((resolve, reject) => {
    try {
      const timeout = setTimeout(() => {
        console.log("REST API token timeout");
        reject(new Error("REST API token timeout"));
      }, 8000);
      
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (tokenResult) => {
        clearTimeout(timeout);
        
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          try {
            console.log("REST API token obtained successfully");
            const notebooks = await makeGraphApiRequest(tokenResult.value, 'REST');
            resolve(notebooks);
          } catch (graphError) {
            console.error("Graph API call failed with REST token:", graphError);
            reject(graphError);
          }
        } else {
          console.error("REST API token failed:", tokenResult.error);
          reject(new Error(`REST API token failed: ${tokenResult.error?.message || 'Unknown error'}`));
        }
      });
      
    } catch (error) {
      console.error("Exception in REST API authentication:", error);
      reject(error);
    }
  });
}

// Method 3: Try Exchange Token (fallback)
export async function tryExchangeToken() {
  return new Promise((resolve, reject) => {
    try {
      const timeout = setTimeout(() => {
        console.log("Exchange token timeout");
        reject(new Error("Exchange token timeout"));
      }, 8000);
      
      Office.context.mailbox.getCallbackTokenAsync({ isRest: false }, async (tokenResult) => {
        clearTimeout(timeout);
        
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Exchange token obtained, but Graph API may not work with this token type");
          // Exchange tokens typically don't work with Graph API, but we can try
          try {
            const notebooks = await makeGraphApiRequest(tokenResult.value, 'Exchange');
            resolve(notebooks);
          } catch (graphError) {
            console.error("Graph API call failed with Exchange token (expected):", graphError);
            reject(graphError);
          }
        } else {
          console.error("Exchange token failed:", tokenResult.error);
          reject(new Error(`Exchange token failed: ${tokenResult.error?.message || 'Unknown error'}`));
        }
      });
      
    } catch (error) {
      console.error("Exception in Exchange authentication:", error);
      reject(error);
    }
  });
}

// Make Microsoft Graph API request with proper error handling
export async function makeGraphApiRequest(accessToken, authMethod) {
  try {
    console.log(`Making Graph API request using ${authMethod} token`);
    
    const graphUrl = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks';
    
    const response = await fetch(graphUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    
    console.log(`Graph API response status: ${response.status}`);
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error(`Graph API request error (${authMethod}): Error: Graph API request failed: ${response.status}  - ${errorText}`);
      
      // If we get InvalidAuthenticationToken error, this means the token audience is wrong
      if (response.status === 401 && errorText.includes('InvalidAuthenticationToken')) {
        console.log(`Token from ${authMethod} method is not compatible with Graph API - this is expected for Office.js tokens`);
        throw new Error(`Token audience mismatch for ${authMethod} method - Office.js tokens cannot be used directly with Graph API`);
      }
      
      throw new Error(`Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`);
    }
    
    const data = await response.json();
    console.log("OneNote notebooks response:", data);
    
    if (data.value && data.value.length > 0) {
      // Map the response to a consistent format
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
      
      console.log(`Found ${notebooks.length} OneNote notebooks using ${authMethod} method`);
      return notebooks;
    } else {
      console.log("No OneNote notebooks found");
      return [];
    }
    
  } catch (error) {
    console.error(`Graph API request failed with ${authMethod} token:`, error);
    throw error;
  }
}

// Generate mock notebook data for testing
export function getMockNotebooks() {
  console.log("🔧 Generating mock OneNote notebooks for testing...");
  
  const mockNotebooks = [
    {
      id: 'mock-notebook-1',
      name: 'Work Notes (Mock Data)',
      displayName: 'Work Notes (Mock Data)',
      isDefault: false,
      createdDateTime: new Date(Date.now() - 86400000 * 30).toISOString(), // 30 days ago
      lastModifiedDateTime: new Date(Date.now() - 86400000 * 5).toISOString(), // 5 days ago
      sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-1/sections',
      sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-1/sectionGroups',
      links: {
        oneNoteClientUrl: {
          href: 'onenote:https://mock-link/work-notes'
        },
        oneNoteWebUrl: {
          href: 'https://mock-onenote-web-url/work-notes'
        }
      }
    },
    {
      id: 'mock-notebook-2',
      name: 'Personal Notes (Mock Data)', 
      displayName: 'Personal Notes (Mock Data)',
      isDefault: true,
      createdDateTime: new Date(Date.now() - 86400000 * 60).toISOString(), // 60 days ago
      lastModifiedDateTime: new Date(Date.now() - 86400000 * 1).toISOString(), // 1 day ago
      sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-2/sections',
      sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-2/sectionGroups',
      links: {
        oneNoteClientUrl: {
          href: 'onenote:https://mock-link/personal-notes'
        },
        oneNoteWebUrl: {
          href: 'https://mock-onenote-web-url/personal-notes'
        }
      }
    },
    {
      id: 'mock-notebook-3',
      name: 'Project Planning (Mock Data)',
      displayName: 'Project Planning (Mock Data)',
      isDefault: false,
      createdDateTime: new Date(Date.now() - 86400000 * 14).toISOString(), // 14 days ago
      lastModifiedDateTime: new Date(Date.now() - 86400000 * 2).toISOString(), // 2 days ago
      sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-3/sections',
      sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/mock-notebook-3/sectionGroups',
      links: {
        oneNoteClientUrl: {
          href: 'onenote:https://mock-link/project-planning'
        },
        oneNoteWebUrl: {
          href: 'https://mock-onenote-web-url/project-planning'
        }
      }
    }
  ];
  
  console.log(`✅ Generated ${mockNotebooks.length} mock notebooks for testing`);
  return mockNotebooks;
}

// Legacy fallback function for compatibility
export async function tryFallbackNotebookRetrieval() {
  console.log("📱 Fallback notebook retrieval - using mock data");
  return getMockNotebooks();
}
