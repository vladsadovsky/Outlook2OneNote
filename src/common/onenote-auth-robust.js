/* eslint-disable no-unused-vars */
/* global Office */
/* eslint-disable no-console */

/*
 * OneNote Service Module - onenote-service.js
 * 
 * This module contains all OneNote-related functionality including:
 * - Getting OneNote notebooks via Microsoft Graph API
 * - Notebook selection UI and popup handling
 * - OneNote page creation and content export
 * - Authentication and token management for Graph API
 * 
 * Dependencies:
 * - Office.js
 * - Microsoft Graph API
 */

import { 
  getSelectedNotebook, 
  setSelectedNotebook, 
  clearSelectedNotebook 
} from '../common/app-state.js';

// Function to get OneNote notebooks using Microsoft Graph API
export async function getOneNoteNotebooks() {
  try {
    console.log("Getting OneNote notebooks...");
    
    // Check platform support first
    const platformSupport = checkPlatformSupport();
    console.log("Platform support check:", platformSupport);
    
    // Try different authentication methods in order of preference
    let notebooks = null;
    
    // Method 1: Try SSO (getAccessTokenAsync) - only if supported
    if (platformSupport.ssoSupported) {
      console.log("Trying SSO authentication...");
      notebooks = await trySSoAuthentication();
    }
    
    // Method 2: Try REST API callback token if SSO failed
    if (!notebooks) {
      console.log("Trying REST API callback token...");
      notebooks = await tryRestApiToken();
    }
    
    // Method 3: Try Exchange callback token if REST failed
    if (!notebooks) {
      console.log("Trying Exchange callback token...");
      notebooks = await tryExchangeToken();
    }
    
    // Method 4: Mock data for development/testing
    if (!notebooks) {
      console.log("All authentication methods failed, using mock data...");
      notebooks = getMockNotebooks();
    }
    
    return notebooks;
    
  } catch (error) {
    console.error("Error in getOneNoteNotebooks:", error);
    console.log("Falling back to mock data due to error");
    return getMockNotebooks();
  }
}

// Check what authentication methods are supported on current platform
function checkPlatformSupport() {
  const support = {
    ssoSupported: false,
    restApiSupported: false,
    exchangeTokenSupported: false,
    platform: 'unknown',
    hostType: Office.context.host.toString()
  };
  
  try {
    // Check for SSO support
    if (Office.context.auth && typeof Office.context.auth.getAccessTokenAsync === 'function') {
      support.ssoSupported = true;
    }
    
    // Check for REST API support
    if (Office.context.mailbox && typeof Office.context.mailbox.getCallbackTokenAsync === 'function') {
      support.restApiSupported = true;
      support.exchangeTokenSupported = true;
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

// Method 1: Try SSO Authentication (getAccessTokenAsync)
async function trySSoAuthentication() {
  return new Promise((resolve) => {
    try {
      // Add timeout to prevent hanging
      const timeout = setTimeout(() => {
        console.log("SSO authentication timeout");
        resolve(null);
      }, 10000);
      
      Office.context.auth.getAccessTokenAsync({ 
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      }, async (tokenResult) => {
        clearTimeout(timeout);
        
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          try {
            console.log("SSO token obtained successfully");
            const notebooks = await makeGraphApiRequest(tokenResult.value, 'SSO');
            resolve(notebooks);
          } catch (graphError) {
            console.error("Graph API call failed with SSO token:", graphError);
            resolve(null);
          }
        } else {
          console.error("SSO authentication failed:", {
            error: tokenResult.error,
            errorCode: tokenResult.error?.code,
            errorMessage: tokenResult.error?.message
          });
          resolve(null);
        }
      });
      
    } catch (error) {
      console.error("Exception in SSO authentication:", error);
      resolve(null);
    }
  });
}

// Method 2: Try REST API Token
async function tryRestApiToken() {
  return new Promise((resolve) => {
    try {
      const timeout = setTimeout(() => {
        console.log("REST API token timeout");
        resolve(null);
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
            resolve(null);
          }
        } else {
          console.error("REST API token failed:", tokenResult.error);
          resolve(null);
        }
      });
      
    } catch (error) {
      console.error("Exception in REST API authentication:", error);
      resolve(null);
    }
  });
}

// Method 3: Try Exchange Token (fallback)
async function tryExchangeToken() {
  return new Promise((resolve) => {
    try {
      const timeout = setTimeout(() => {
        console.log("Exchange token timeout");
        resolve(null);
      }, 8000);
      
      Office.context.mailbox.getCallbackTokenAsync({ isRest: false }, async (tokenResult) => {
        clearTimeout(timeout);
        
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Exchange token obtained, but Graph API may not work with this token type");
          // Exchange tokens typically don't work with Graph API, but we can try
          resolve(null);
        } else {
          console.error("Exchange token failed:", tokenResult.error);
          resolve(null);
        }
      });
      
    } catch (error) {
      console.error("Exception in Exchange authentication:", error);
      resolve(null);
    }
  });
}

// Make Microsoft Graph API request
async function makeGraphApiRequest(accessToken, authMethod) {
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
      throw new Error(`Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`);
    }
    
    const data = await response.json();
    console.log(`OneNote notebooks response (${authMethod}):`, data);
    
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
        links: notebook.links,
        authMethod: authMethod
      }));
      
      console.log(`Found ${notebooks.length} OneNote notebooks using ${authMethod}`);
      return notebooks;
    } else {
      console.log(`No OneNote notebooks found using ${authMethod}`);
      return null;
    }
    
  } catch (error) {
    console.error(`Graph API request error (${authMethod}):`, error);
    throw error;
  }
}

// Get mock notebooks for development/testing
function getMockNotebooks() {
  console.log("Providing mock OneNote notebooks for development/testing");
  return [
    {
      id: 'mock-notebook-dev-1',
      name: 'Development Test Notebook',
      displayName: 'Development Test Notebook',
      isDefault: true,
      createdDateTime: new Date().toISOString(),
      lastModifiedDateTime: new Date().toISOString(),
      authMethod: 'mock',
      sectionsUrl: 'mock://sections',
      sectionGroupsUrl: 'mock://sectiongroups'
    },
    {
      id: 'mock-notebook-dev-2',
      name: 'Email Export Test Notebook',
      displayName: 'Email Export Test Notebook',
      isDefault: false,
      createdDateTime: new Date(Date.now() - 24*60*60*1000).toISOString(), // Yesterday
      lastModifiedDateTime: new Date().toISOString(),
      authMethod: 'mock',
      sectionsUrl: 'mock://sections',
      sectionGroupsUrl: 'mock://sectiongroups'
    }
  ];
}

// Function to show notebook selection popup
export function showNotebookPopup(notebooks, onNotebookSelected) {
  console.log("Showing notebook selection popup with", notebooks.length, "notebooks");
  
  const insertAt = document.getElementById("item-subject");
  
  // Clear any existing content
  insertAt.innerHTML = "";
  
  // Create header
  const header = document.createElement("h3");
  header.appendChild(document.createTextNode("Select a OneNote Notebook:"));
  insertAt.appendChild(header);
  
  // Show authentication method info if available
  if (notebooks.length > 0 && notebooks[0].authMethod) {
    const authInfo = document.createElement("div");
    authInfo.style.fontSize = "12px";
    authInfo.style.color = "#666";
    authInfo.style.marginBottom = "10px";
    authInfo.style.fontStyle = "italic";
    
    const authMethod = notebooks[0].authMethod;
    let authText = "";
    
    switch (authMethod) {
      case 'SSO':
        authText = "✓ Connected via Single Sign-On";
        authInfo.style.color = "#28a745";
        break;
      case 'REST':
        authText = "✓ Connected via REST API";
        authInfo.style.color = "#28a745";
        break;
      case 'mock':
        authText = "⚠ Using mock data for development/testing";
        authInfo.style.color = "#fd7e14";
        break;
      default:
        authText = "✓ Connected";
        authInfo.style.color = "#28a745";
    }
    
    authInfo.appendChild(document.createTextNode(authText));
    insertAt.appendChild(authInfo);
  }
  
  if (notebooks.length === 0) {
    const noNotebooks = document.createElement("div");
    noNotebooks.style.color = "#dc3545";
    noNotebooks.appendChild(document.createTextNode("No OneNote notebooks found. Please create a notebook in OneNote first."));
    insertAt.appendChild(noNotebooks);
    return;
  }
  
  // Create notebook list
  const notebookList = document.createElement("div");
  notebookList.style.marginTop = "10px";
  notebookList.style.maxHeight = "300px";
  notebookList.style.overflowY = "auto";
  notebookList.style.border = "1px solid #ccc";
  notebookList.style.borderRadius = "4px";
  notebookList.style.padding = "10px";
  
  notebooks.forEach((notebook, index) => {
    // Create notebook item
    const notebookItem = document.createElement("div");
    notebookItem.style.padding = "8px";
    notebookItem.style.margin = "4px 0";
    notebookItem.style.border = "1px solid #ddd";
    notebookItem.style.borderRadius = "4px";
    notebookItem.style.cursor = "pointer";
    notebookItem.style.backgroundColor = "#f9f9f9";
    
    // Highlight default notebook
    if (notebook.isDefault) {
      notebookItem.style.backgroundColor = "#e6f3ff";
      notebookItem.style.borderColor = "#4a90e2";
    }
    
    // Highlight mock notebooks differently
    if (notebook.authMethod === 'mock') {
      notebookItem.style.backgroundColor = "#fff3cd";
      notebookItem.style.borderColor = "#ffc107";
    }
    
    // Add hover effect
    notebookItem.addEventListener("mouseenter", () => {
      if (notebook.authMethod === 'mock') {
        notebookItem.style.backgroundColor = "#ffeaa7";
      } else {
        notebookItem.style.backgroundColor = notebook.isDefault ? "#d1e9ff" : "#f0f0f0";
      }
    });
    
    notebookItem.addEventListener("mouseleave", () => {
      if (notebook.authMethod === 'mock') {
        notebookItem.style.backgroundColor = "#fff3cd";
      } else {
        notebookItem.style.backgroundColor = notebook.isDefault ? "#e6f3ff" : "#f9f9f9";
      }
    });
    
    // Create notebook content
    const notebookContent = document.createElement("div");
    
    const notebookTitle = document.createElement("div");
    notebookTitle.style.fontWeight = "bold";
    notebookTitle.style.marginBottom = "4px";
    notebookTitle.appendChild(document.createTextNode(notebook.displayName || notebook.name));
    
    // Add labels
    const labelsContainer = document.createElement("div");
    labelsContainer.style.marginTop = "4px";
    
    if (notebook.isDefault) {
      const defaultLabel = document.createElement("span");
      defaultLabel.style.fontSize = "11px";
      defaultLabel.style.color = "#4a90e2";
      defaultLabel.style.marginRight = "8px";
      defaultLabel.style.padding = "2px 6px";
      defaultLabel.style.backgroundColor = "#e3f2fd";
      defaultLabel.style.borderRadius = "3px";
      defaultLabel.appendChild(document.createTextNode("Default"));
      labelsContainer.appendChild(defaultLabel);
    }
    
    if (notebook.authMethod === 'mock') {
      const mockLabel = document.createElement("span");
      mockLabel.style.fontSize = "11px";
      mockLabel.style.color = "#856404";
      mockLabel.style.marginRight = "8px";
      mockLabel.style.padding = "2px 6px";
      mockLabel.style.backgroundColor = "#fff3cd";
      mockLabel.style.borderRadius = "3px";
      mockLabel.appendChild(document.createTextNode("Test Data"));
      labelsContainer.appendChild(mockLabel);
    }
    
    notebookContent.appendChild(notebookTitle);
    notebookContent.appendChild(labelsContainer);
    
    // Add creation date if available
    if (notebook.createdDateTime) {
      const dateInfo = document.createElement("div");
      dateInfo.style.fontSize = "12px";
      dateInfo.style.color = "#666";
      dateInfo.style.marginTop = "4px";
      const createdDate = new Date(notebook.createdDateTime);
      dateInfo.appendChild(document.createTextNode(`Created: ${createdDate.toLocaleDateString()}`));
      notebookContent.appendChild(dateInfo);
    }
    
    notebookItem.appendChild(notebookContent);
    
    // Add click handler
    notebookItem.addEventListener("click", () => {
      console.log("Notebook selected:", notebook);
      
      // Store selected notebook in state
      setSelectedNotebook(notebook);
      
      // Highlight selected item
      const allItems = notebookList.querySelectorAll("div[data-notebook-item]");
      allItems.forEach(item => {
        const itemNotebook = notebooks.find(nb => nb.id === item.dataset.notebookId);
        if (itemNotebook) {
          if (itemNotebook.authMethod === 'mock') {
            item.style.backgroundColor = "#fff3cd";
            item.style.borderColor = "#ffc107";
          } else {
            item.style.backgroundColor = itemNotebook.isDefault ? "#e6f3ff" : "#f9f9f9";
            item.style.borderColor = itemNotebook.isDefault ? "#4a90e2" : "#ddd";
          }
        }
      });
      
      notebookItem.style.backgroundColor = "#d4edda";
      notebookItem.style.borderColor = "#28a745";
      
      // Call the callback with selected notebook
      if (onNotebookSelected) {
        onNotebookSelected(notebook);
      }
    });
    
    // Add data attributes for selection tracking
    notebookItem.setAttribute("data-notebook-item", "true");
    notebookItem.setAttribute("data-notebook-id", notebook.id);
    notebookItem.setAttribute("data-is-default", notebook.isDefault ? "true" : "false");
    
    notebookList.appendChild(notebookItem);
  });
  
  insertAt.appendChild(notebookList);
  
  // Add instructions
  const instructions = document.createElement("div");
  instructions.style.marginTop = "10px";
  instructions.style.fontSize = "12px";
  instructions.style.color = "#666";
  instructions.style.fontStyle = "italic";
  instructions.appendChild(document.createTextNode("Click on a notebook to select it for email export."));
  insertAt.appendChild(instructions);
}

// Default callback for notebook selection
export function onNotebookSelected(notebook) {
  console.log("Notebook selected (default callback):", notebook);
  
  const insertAt = document.getElementById("item-subject");
  insertAt.innerHTML = "";
  insertAt.appendChild(document.createTextNode(`Selected: ${notebook.displayName || notebook.name}`));
  insertAt.appendChild(document.createElement("br"));
  
  if (notebook.authMethod === 'mock') {
    const warning = document.createElement("div");
    warning.style.color = "#856404";
    warning.style.fontSize = "12px";
    warning.style.marginTop = "5px";
    warning.appendChild(document.createTextNode("Note: Using test data. Actual OneNote integration requires authentication."));
    insertAt.appendChild(warning);
  } else {
    insertAt.appendChild(document.createTextNode("Notebook selection complete. Ready to export emails."));
  }
}

// Helper function to export conversation to OneNote
export async function exportConversationToOneNote(conversationData, notebook, insertAt) {
  try {
    console.log("Exporting to OneNote notebook:", notebook);
    
    // Sort conversation data by date
    conversationData.sort((a, b) => {
      const dateA = new Date(a.DateTimeReceived || a.DateTimeSent || 0);
      const dateB = new Date(b.DateTimeReceived || b.DateTimeSent || 0);
      return dateA - dateB;
    });
    
    // Create page title based on first email subject
    const pageTitle = conversationData.length > 0 ? 
      `Email Thread: ${conversationData[0].Subject || 'No Subject'}` : 
      'Email Thread Export';
    
    // Build OneNote page content
    let pageContent = `
      <html>
        <head>
          <title>${pageTitle}</title>
        </head>
        <body>
          <h1>${pageTitle}</h1>
          <p><strong>Exported on:</strong> ${new Date().toLocaleString()}</p>
          <p><strong>Total emails:</strong> ${conversationData.length}</p>
          <p><strong>Target notebook:</strong> ${notebook.displayName || notebook.name}</p>
          <hr />
    `;
    
    // Add each email to the page content
    conversationData.forEach((email, index) => {
      pageContent += `
        <div style="margin-bottom: 20px; padding: 10px; border: 1px solid #ccc; border-radius: 4px;">
          <h3>Email ${index + 1}</h3>
          <p><strong>Subject:</strong> ${email.Subject || 'No Subject'}</p>
          <p><strong>From:</strong> ${email.Sender || 'Unknown'}</p>
          <p><strong>Date:</strong> ${email.DateTimeReceived || email.DateTimeSent || 'Unknown'}</p>
          <hr />
          <div style="margin-top: 10px;">
            ${(email.Body || 'No content available').replace(/\n/g, '<br>')}
          </div>
        </div>
      `;
    });
    
    pageContent += `
        </body>
      </html>
    `;
    
    // Show export status
    insertAt.appendChild(document.createTextNode("✓ Conversation data prepared for export"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`✓ Page title: ${pageTitle}`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`✓ ${conversationData.length} emails ready for export`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    if (notebook.authMethod === 'mock') {
      insertAt.appendChild(document.createTextNode("⚠ Mock notebook selected - no actual OneNote page will be created."));
    } else {
      insertAt.appendChild(document.createTextNode("Note: OneNote API integration pending - data is prepared but not yet sent to OneNote."));
    }
    
    console.log("OneNote page content prepared:", pageContent.substring(0, 500) + "...");
    
  } catch (error) {
    console.error("Error exporting to OneNote:", error);
    throw error;
  }
}

// Helper function to export single email to OneNote
export async function exportSingleEmailToOneNote(item, notebook, insertAt) {
  try {
    const pageTitle = `Email: ${item.subject || "No Subject"}`;
    
    insertAt.appendChild(document.createTextNode("✓ Single email prepared for export"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`✓ Page title: ${pageTitle}`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    if (notebook.authMethod === 'mock') {
      insertAt.appendChild(document.createTextNode("⚠ Mock notebook selected - no actual OneNote page will be created."));
    } else {
      insertAt.appendChild(document.createTextNode("Note: OneNote API integration pending - single email export ready."));
    }
    
    console.log("Single email export prepared for:", pageTitle);
    
  } catch (error) {
    console.error("Error exporting single email:", error);
    throw error;
  }
}

// Re-export the state management functions for convenience
export { getSelectedNotebook, setSelectedNotebook, clearSelectedNotebook };