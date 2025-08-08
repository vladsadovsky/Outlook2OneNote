/* eslint-disable no-unused-vars */
/* global Office */
/* eslint-disable no-console */

/*
 * OneNote Service Module - onenote-service.js (Modernized with PKCE)
 * 
 * This module contains all OneNote-related functionality including:
 * - Getting OneNote notebooks via Microsoft Graph API with PKCE authentication
 * - Notebook selection UI and popup handling
 * - OneNote page creation and content export
 * - Modern authentication with automatic token refresh
 * 
 * Authentication Improvements:
 * - Primary PKCE OAuth 2.0 Authorization Code Flow
 * - SSO fallback for Office Add-ins
 * - Secure token storage and automatic refresh
 * - No client secrets required
 * 
 * Dependencies:
 * - Office.js
 * - Microsoft Graph API
 * - PKCE Authentication Module
 */

import { 
  getSelectedNotebook, 
  setSelectedNotebook, 
  clearSelectedNotebook 
} from '../common/app-state.js';

import { 
  authenticateAndGetNotebooks,
  getMockNotebooks,
  pkceAuth,
  hasValidToken,
  refreshToken,
  handleAuthorizationCallback,
  logout
} from '../common/graphapi-auth.js';

// Function to get OneNote notebooks using modern PKCE authentication
export async function getOneNoteNotebooks() {
  try {
    console.log("🚀 Getting OneNote notebooks with modern PKCE authentication...");
    
    // Check if we have a valid token first
    const hasValid = await hasValidToken();
    console.log("Has valid token:", hasValid);
    
    // If we don't have a valid token, try to refresh or start new auth flow
    if (!hasValid) {
      console.log("🔄 No valid token, attempting authentication...");
    }
    
    // Attempt authentication - this will handle PKCE flow or fallback to SSO
    const notebooks = await authenticateAndGetNotebooks();
    
    if (notebooks === null) {
      // PKCE flow is in progress, user is being redirected
      console.log("🔄 PKCE authentication flow started, waiting for user...");
      
      // Show user-friendly message while waiting
      const insertAt = document.getElementById("item-subject");
      if (insertAt) {
        insertAt.innerHTML = "";
        insertAt.appendChild(document.createTextNode("🔐 Redirecting to Microsoft for secure authentication..."));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("Please complete the sign-in process and return here."));
      }
      
      return null; // Indicate that auth flow is in progress
    }
    
    if (notebooks && notebooks.length > 0) {
      console.log(`✅ Successfully retrieved ${notebooks.length} OneNote notebooks`);
      return notebooks;
    } else {
      console.log("📭 No notebooks found, using mock data");
      return getMockNotebooks();
    }
    
  } catch (error) {
    console.error("❌ Error in getOneNoteNotebooks:", error);
    
    // Try to provide helpful error messages
    if (error.message && error.message.includes('consent')) {
      console.log("📝 User consent may be required for OneNote access");
    }
    
    if (error.message && error.message.includes('token')) {
      console.log("🔄 Token issue detected, may need to re-authenticate");
    }
    
    // Return mock data for testing/development
    console.log("🔧 Falling back to mock data for development");
    return getMockNotebooks();
  }
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
  
  if (notebooks.length === 0) {
    insertAt.appendChild(document.createTextNode("No OneNote notebooks found. Please create a notebook in OneNote first."));
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
    
    // Add hover effect
    notebookItem.addEventListener("mouseenter", () => {
      notebookItem.style.backgroundColor = notebook.isDefault ? "#d1e9ff" : "#f0f0f0";
    });
    
    notebookItem.addEventListener("mouseleave", () => {
      notebookItem.style.backgroundColor = notebook.isDefault ? "#e6f3ff" : "#f9f9f9";
    });
    
    // Create notebook content
    const notebookContent = document.createElement("div");
    
    const notebookTitle = document.createElement("div");
    notebookTitle.style.fontWeight = "bold";
    notebookTitle.style.marginBottom = "4px";
    notebookTitle.appendChild(document.createTextNode(notebook.displayName || notebook.name));
    
    if (notebook.isDefault) {
      const defaultLabel = document.createElement("span");
      defaultLabel.style.fontSize = "12px";
      defaultLabel.style.color = "#4a90e2";
      defaultLabel.style.marginLeft = "8px";
      defaultLabel.appendChild(document.createTextNode("(Default)"));
      notebookTitle.appendChild(defaultLabel);
    }
    
    notebookContent.appendChild(notebookTitle);
    
    // Add creation date if available
    if (notebook.createdDateTime) {
      const dateInfo = document.createElement("div");
      dateInfo.style.fontSize = "12px";
      dateInfo.style.color = "#666";
      const createdDate = new Date(notebook.createdDateTime);
      dateInfo.appendChild(document.createTextNode(`Created: ${createdDate.toLocaleDateString()}`));
      notebookContent.appendChild(dateInfo);
    }
    
    notebookItem.appendChild(notebookContent);
    
    // Add click handler
    notebookItem.addEventListener("click", () => {
      console.log("Notebook selected:", notebook);
      
      // Highlight selected item
      const allItems = notebookList.querySelectorAll("div[data-notebook-item]");
      allItems.forEach(item => {
        item.style.backgroundColor = item.dataset.isDefault === "true" ? "#e6f3ff" : "#f9f9f9";
        item.style.borderColor = item.dataset.isDefault === "true" ? "#4a90e2" : "#ddd";
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
  insertAt.appendChild(document.createTextNode("Notebook selection complete."));
}

// Helper function to export conversation to OneNote
export async function exportConversationToOneNote(conversationData, notebook, insertAt) {
  try {
    console.log("Exporting to OneNote notebook:", notebook);
    
    // Sort conversation data by date
    conversationData.sort((a, b) => a.date - b.date);
    
    // Create page title based on first email subject
    const pageTitle = `Email Thread: ${conversationData[0].subject}`;
    
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
          <hr />
    `;
    
    // Add each email to the page content
    conversationData.forEach((email, index) => {
      pageContent += `
        <div style="margin-bottom: 20px; padding: 10px; border: 1px solid #ccc;">
          <h3>Email ${index + 1}</h3>
          <p><strong>Subject:</strong> ${email.subject}</p>
          <p><strong>From:</strong> ${email.senderName} ${email.senderEmail ? `(${email.senderEmail})` : ''}</p>
          <p><strong>Date:</strong> ${email.date.toLocaleString()}</p>
          <hr />
          <div style="margin-top: 10px;">
            ${email.body.replace(/\n/g, '<br>')}
          </div>
        </div>
      `;
    });
    
    pageContent += `
        </body>
      </html>
    `;
    
    // TODO: Implement actual OneNote API call here
    // For now, show what would be exported
    insertAt.appendChild(document.createTextNode("✓ Conversation data prepared for export"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`✓ Page title: ${pageTitle}`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`✓ ${conversationData.length} emails ready for export`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode("Note: OneNote API integration pending - data is prepared but not yet sent to OneNote."));
    
    console.log("OneNote page content prepared:", pageContent);
    
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
    insertAt.appendChild(document.createTextNode("Note: OneNote API integration pending - single email export ready."));
    
    console.log("Single email export prepared for:", pageTitle);
    
  } catch (error) {
    console.error("Error exporting single email:", error);
    throw error;
  }
}

// Modern authentication utility functions

/**
 * Check if user has valid authentication tokens
 */
export async function checkAuthenticationStatus() {
  try {
    const hasValid = await hasValidToken();
    console.log("Authentication status check:", hasValid ? "Valid" : "Invalid/Missing");
    return hasValid;
  } catch (error) {
    console.error("Error checking authentication status:", error);
    return false;
  }
}

/**
 * Handle the OAuth authorization callback (for PKCE flow)
 * This should be called when user returns from authorization
 */
export async function handleOAuthCallback() {
  try {
    console.log("🔄 Handling OAuth authorization callback...");
    const notebooks = await handleAuthorizationCallback();
    
    if (notebooks && notebooks.length > 0) {
      console.log("✅ OAuth callback handled successfully, got notebooks");
      return notebooks;
    } else {
      console.log("⚠️ OAuth callback completed but no notebooks returned");
      return [];
    }
  } catch (error) {
    console.error("❌ OAuth callback handling failed:", error);
    throw error;
  }
}

/**
 * Force refresh authentication tokens
 */
export async function refreshAuthenticationTokens() {
  try {
    console.log("🔄 Refreshing authentication tokens...");
    await refreshToken();
    console.log("✅ Tokens refreshed successfully");
    return true;
  } catch (error) {
    console.error("❌ Token refresh failed:", error);
    return false;
  }
}

/**
 * Clear authentication and logout user
 */
export async function logoutUser() {
  try {
    console.log("👋 Logging out user...");
    await logout();
    
    // Clear local notebook selection
    clearSelectedNotebook();
    
    // Update UI to reflect logout
    const insertAt = document.getElementById("item-subject");
    if (insertAt) {
      insertAt.innerHTML = "";
      insertAt.appendChild(document.createTextNode("👋 Logged out successfully. Click 'Choose Notebook' to sign in again."));
    }
    
    console.log("✅ User logged out successfully");
  } catch (error) {
    console.error("❌ Logout failed:", error);
    throw error;
  }
}

/**
 * Get current authentication method being used
 */
export function getCurrentAuthMethod() {
  // Check platform support and determine which auth method is active
  try {
    const support = pkceAuth.cryptoSupport || null;
    
    if (support && support.webCrypto) {
      return 'PKCE OAuth 2.0';
    } else if (typeof Office !== 'undefined' && Office.context && Office.context.auth) {
      return 'Office.js SSO';
    } else {
      return 'Mock Data (Development)';
    }
  } catch (error) {
    console.warn("Error determining auth method:", error);
    return 'Unknown';
  }
}

/**
 * Show authentication status in UI
 */
export async function showAuthenticationStatus() {
  const insertAt = document.getElementById("item-subject");
  if (!insertAt) return;
  
  try {
    const hasValid = await checkAuthenticationStatus();
    const authMethod = getCurrentAuthMethod();
    
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("🔐 Authentication Status"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    insertAt.appendChild(document.createTextNode(`Method: ${authMethod}`));
    insertAt.appendChild(document.createElement("br"));
    
    insertAt.appendChild(document.createTextNode(`Status: ${hasValid ? '✅ Valid' : '❌ Invalid/Missing'}`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    if (hasValid) {
      insertAt.appendChild(document.createTextNode("You can access OneNote notebooks."));
    } else {
      insertAt.appendChild(document.createTextNode("Click 'Choose Notebook' to authenticate."));
    }
    
  } catch (error) {
    console.error("Error showing authentication status:", error);
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("❌ Error checking authentication status: " + error.message));
  }
}
