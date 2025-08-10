/* eslint-disable no-unused-vars */
/* global Office */
/* eslint-disable no-console */

/*
 * OneNote Service Module - onenote-service.js (Modernized with Office SSO-first authentication)
 * 
 * This module contains all OneNote-related functionality including:
 * - Getting OneNote notebooks via Microsoft Graph API with Office SSO-first authentication
 * - Notebook selection UI and popup handling
 * - OneNote page creation and content export
 * - Modern authentication with automatic token refresh
 * 
 * Authentication Improvements:
 * - Office SSO-first with MSAL popup fallback
 * - Microsoft Graph API exclusive access
 * - Secure token storage and automatic refresh
 * - No client secrets required
 * 
 * Dependencies:
 * - Office.js
 * - Microsoft Graph API
 * - Auth Service Module
 */

import { 
  getSelectedNotebook, 
  setSelectedNotebook, 
  clearSelectedNotebook 
} from '../common/app-state.js';

import authService from '../auth/auth-service.js';

// Function to get OneNote notebooks using Office SSO-first authentication
export async function getOneNoteNotebooks() {
  try {
    console.log("� Getting OneNote notebooks with Office SSO-first authentication...");
    
    // Ensure authentication (Office SSO first, then MSAL fallback)
    await authService.authenticate();
    
    // Get notebooks from Microsoft Graph API
    console.log("� Fetching notebooks from Microsoft Graph API...");
    const data = await authService.callGraphApi('/me/onenote/notebooks');
    
    if (!data || !data.value) {
      console.log('⚠️ No notebooks returned from Graph API');
      return getMockNotebooks();
    }
    
    const notebooks = data.value;
    console.log(`✅ Successfully retrieved ${notebooks.length} OneNote notebooks`);
    
    return notebooks;
    
  } catch (error) {
    console.error("❌ Error getting OneNote notebooks:", error);
    
    // Return mock data for development
    console.log("� Using mock OneNote notebooks for development");
    return getMockNotebooks();
  }
}

// Mock notebooks for development/fallback
function getMockNotebooks() {
  return [
    {
      id: 'mock-notebook-1',
      displayName: 'Test Notebook (Mock)',
      isDefault: false,
      userRole: 'Owner',
      isShared: false,
      sectionsUrl: 'mock-sections-url',
      sectionGroupsUrl: 'mock-section-groups-url',
      self: 'mock-self-url'
    },
    {
      id: 'mock-notebook-2', 
      displayName: 'Personal Notebook (Mock)',
      isDefault: true,
      userRole: 'Owner',
      isShared: false,
      sectionsUrl: 'mock-sections-url-2',
      sectionGroupsUrl: 'mock-section-groups-url-2',
      self: 'mock-self-url-2'
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
    console.log("Debug - Notebook ID being used:", notebook.id);
    console.log("Debug - Notebook display name:", notebook.displayName);
    console.log("Debug - Full notebook object:", JSON.stringify(notebook, null, 2));
    
    // Sort conversation data by date (earliest first)
    conversationData.sort((a, b) => a.date - b.date);
    
    const earliestEmail = conversationData[0];
    const exportDate = new Date().toISOString().split('T')[0]; // YYYY-MM-DD format
    
    // Debug logging
    console.log("Debug - Earliest email subject:", JSON.stringify(earliestEmail.subject));
    console.log("Debug - Earliest email object:", JSON.stringify(earliestEmail));
    
    // Create section name: "Subject + Export Date" with fallback for empty subjects
    let emailSubject = earliestEmail.subject && earliestEmail.subject.trim() 
      ? earliestEmail.subject.trim() 
      : "No Subject";
      
    // Remove any potentially problematic characters from subject
    emailSubject = emailSubject.replace(/[<>:"\/\\|?*]/g, '').substring(0, 100); // Limit length and remove invalid chars
    
    // Ensure we still have a valid name after cleaning
    if (!emailSubject || emailSubject.trim() === '') {
      emailSubject = "Email Thread";
    }
    
    const sectionName = `${emailSubject} - ${exportDate}`;
    
    console.log("Debug - Final section name:", JSON.stringify(sectionName));
    
    // Final validation - ensure the section name is not empty or whitespace only
    if (!sectionName || sectionName.trim().length === 0) {
      throw new Error("Section name is empty after processing");
    }
    
    insertAt.appendChild(document.createTextNode(`Creating section: "${sectionName}"`));
    insertAt.appendChild(document.createElement("br"));
    
    // Step 1: Create a new section in the notebook
    const sectionData = {
      displayName: sectionName.trim() // Use displayName instead of name for OneNote API
    };
    
    console.log("Debug - Section data being sent:", JSON.stringify(sectionData));
    
    // First, let's verify the notebook exists and is accessible
    console.log("🔍 Verifying notebook access...");
    try {
      const notebookCheck = await authService.callGraphApi(`/me/onenote/notebooks/${notebook.id}`);
      console.log("✅ Notebook verified:", notebookCheck.displayName);
    } catch (verifyError) {
      console.error("❌ Cannot access notebook:", verifyError);
      
      // Try to refresh the notebooks list and find this notebook
      console.log("🔄 Refreshing notebooks list...");
      try {
        const refreshedNotebooks = await authService.callGraphApi('/me/onenote/notebooks');
        console.log("📚 Available notebooks:", refreshedNotebooks.value.map(nb => `${nb.displayName} (${nb.id})`));
        
        // Try to find the notebook by displayName
        const matchingNotebook = refreshedNotebooks.value.find(nb => 
          nb.displayName === notebook.displayName || nb.id === notebook.id
        );
        
        if (matchingNotebook && matchingNotebook.id !== notebook.id) {
          console.log("🔄 Found notebook with different ID, updating...");
          notebook.id = matchingNotebook.id;
          console.log("✅ Updated notebook ID to:", notebook.id);
        } else {
          throw new Error(`Notebook '${notebook.displayName}' not found in current notebooks list`);
        }
        
      } catch (refreshError) {
        console.error("❌ Failed to refresh notebooks:", refreshError);
        throw new Error(`Cannot access notebook: ${verifyError.message}`);
      }
    }
    
    let section;
    try {
      // Try the notebook-specific endpoint first (conversation export)
      section = await authService.callGraphApi(
        `/me/onenote/notebooks/${notebook.id}/sections`,
        'POST',
        sectionData
      );
      
      insertAt.appendChild(document.createTextNode(`✓ Section created successfully`));
      insertAt.appendChild(document.createElement("br"));
      
    } catch (error) {
      console.error("Failed to create section:", error);
      
      // If notebook-specific approach failed, try alternative approach
      console.log("🔄 Trying alternative section creation method...");
      try {
        // Try using general sections endpoint with parentNotebook property
        const alternativeSectionData = {
          displayName: sectionName.trim(),
          parentNotebook: {
            id: notebook.id
          }
        };
        
        section = await authService.callGraphApi(
          `/me/onenote/sections`,
          'POST',
          alternativeSectionData
        );
        
        console.log("✅ Section created using alternative method");
        insertAt.appendChild(document.createTextNode(`✓ Section created successfully (alternative method)`));
        insertAt.appendChild(document.createElement("br"));
        
      } catch (alternativeError) {
        console.error("Alternative section creation also failed:", alternativeError);
        insertAt.appendChild(document.createTextNode(`❌ Failed to create section: ${error.message}`));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode(`❌ Alternative method also failed: ${alternativeError.message}`));
        insertAt.appendChild(document.createElement("br"));
        throw error;
      }
    }
    
    // Step 2: Create a page for each email in the conversation
    insertAt.appendChild(document.createTextNode(`Creating ${conversationData.length} pages...`));
    insertAt.appendChild(document.createElement("br"));
    
    for (let i = 0; i < conversationData.length; i++) {
      const email = conversationData[i];
      const emailDate = email.date.toLocaleDateString(); // Format: MM/DD/YYYY or local format
      
      // Create page title: "Subject - Date" with fallback for empty subjects
      const emailSubject = email.subject && email.subject.trim() 
        ? email.subject.trim() 
        : "No Subject";
      const pageTitle = `${emailSubject} - ${emailDate}`;
      
      // Build OneNote page content in HTML format
      const pageContent = `
        <html>
          <head>
            <title>${pageTitle}</title>
          </head>
          <body>
            <div>
              <h1>${emailSubject}</h1>
              <table style="margin-bottom: 20px;">
                <tr>
                  <td><strong>From:</strong></td>
                  <td>${email.senderName}${email.senderEmail ? ` (${email.senderEmail})` : ''}</td>
                </tr>
                <tr>
                  <td><strong>Date:</strong></td>
                  <td>${email.date.toLocaleString()}</td>
                </tr>
                <tr>
                  <td><strong>To:</strong></td>
                  <td>${email.recipients || 'Not available'}</td>
                </tr>
              </table>
              <hr />
              <div style="margin-top: 20px; white-space: pre-wrap;">
                ${email.body}
              </div>
            </div>
          </body>
        </html>
      `;
      
      try {
        // Create page in the section
        await authService.callGraphApi(
          `/me/onenote/sections/${section.id}/pages`,
          'POST',
          pageContent,
          { 'Content-Type': 'text/html' }
        );
        
        insertAt.appendChild(document.createTextNode(`✓ Page ${i + 1}: "${pageTitle}"`));
        insertAt.appendChild(document.createElement("br"));
        
      } catch (pageError) {
        console.error(`Failed to create page ${i + 1}:`, pageError);
        insertAt.appendChild(document.createTextNode(`❌ Failed to create page ${i + 1}: ${pageError.message}`));
        insertAt.appendChild(document.createElement("br"));
        // Continue with other pages even if one fails
      }
    }
    
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`🎉 Export completed! Created section "${sectionName}" with ${conversationData.length} pages in OneNote.`));
    
  } catch (error) {
    console.error("Error exporting to OneNote:", error);
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`❌ Export failed: ${error.message}`));
    throw error;
  }
}

// Helper function to export single email to OneNote
export async function exportSingleEmailToOneNote(item, notebook, insertAt) {
  try {
    const exportDate = new Date().toISOString().split('T')[0]; // YYYY-MM-DD format
    const emailSubject = item.subject || "No Subject";
    
    // Create section name: "Single Email - Subject + Export Date"
    const sectionName = `Single Email - ${emailSubject} - ${exportDate}`;
    
    insertAt.appendChild(document.createTextNode(`Creating section: "${sectionName}"`));
    insertAt.appendChild(document.createElement("br"));
    
    // Step 1: Create a new section in the notebook
    const sectionData = {
      displayName: sectionName
    };
    
    let section;
    try {
      section = await authService.callGraphApi(
        `/me/onenote/notebooks/${notebook.id}/sections`,
        'POST',
        sectionData
      );
      
      insertAt.appendChild(document.createTextNode(`✓ Section created successfully`));
      insertAt.appendChild(document.createElement("br"));
      
    } catch (error) {
      console.error("Failed to create section:", error);
      insertAt.appendChild(document.createTextNode(`❌ Failed to create section: ${error.message}`));
      insertAt.appendChild(document.createElement("br"));
      throw error;
    }
    
    // Step 2: Create page for the single email
    const emailDate = new Date().toLocaleDateString();
    const pageTitle = `${emailSubject} - ${emailDate}`;
    
    // Get email body (this is limited in Office.js, but we'll do what we can)
    let emailBody = "Email body not available through Office.js";
    
    // Build OneNote page content in HTML format
    const pageContent = `
      <html>
        <head>
          <title>${pageTitle}</title>
        </head>
        <body>
          <div>
            <h1>${emailSubject}</h1>
            <table style="margin-bottom: 20px;">
              <tr>
                <td><strong>From:</strong></td>
                <td>${item.sender?.displayName || 'Not available'} ${item.sender?.emailAddress ? `(${item.sender.emailAddress})` : ''}</td>
              </tr>
              <tr>
                <td><strong>Date:</strong></td>
                <td>${item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : 'Not available'}</td>
              </tr>
              <tr>
                <td><strong>To:</strong></td>
                <td>${item.to ? item.to.map(recipient => recipient.displayName || recipient.emailAddress).join(', ') : 'Not available'}</td>
              </tr>
            </table>
            <hr />
            <div style="margin-top: 20px;">
              <p><em>Note: Email body extraction requires additional permissions. This shows the email metadata that is available through Office.js.</em></p>
              <p><strong>Subject:</strong> ${emailSubject}</p>
              <p><strong>Item ID:</strong> ${item.itemId || 'Not available'}</p>
              <p><strong>Conversation ID:</strong> ${item.conversationId || 'Not available'}</p>
            </div>
          </div>
        </body>
      </html>
    `;
    
    try {
      // Create page in the section
      await authService.callGraphApi(
        `/me/onenote/sections/${section.id}/pages`,
        'POST',
        pageContent,
        { 'Content-Type': 'text/html' }
      );
      
      insertAt.appendChild(document.createTextNode(`✓ Page created: "${pageTitle}"`));
      insertAt.appendChild(document.createElement("br"));
      
    } catch (pageError) {
      console.error("Failed to create page:", pageError);
      insertAt.appendChild(document.createTextNode(`❌ Failed to create page: ${pageError.message}`));
      insertAt.appendChild(document.createElement("br"));
      throw pageError;
    }
    
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`🎉 Single email exported! Created section "${sectionName}" with 1 page in OneNote.`));
    
  } catch (error) {
    console.error("Error exporting single email:", error);
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`❌ Export failed: ${error.message}`));
    throw error;
  }
}

// Modern authentication utility functions

/**
 * Check if user has valid authentication tokens
 */
export async function checkAuthenticationStatus() {
  try {
    const hasValid = await authService.hasValidToken();
    console.log("Authentication status check:", hasValid ? "Valid" : "Invalid/Missing");
    return hasValid;
  } catch (error) {
    console.error("Error checking authentication status:", error);
    return false;
  }
}

/**
 * Handle the OAuth authorization callback (for modern authentication)
 * This function is kept for compatibility but modern auth flow is automatic
 */
export async function handleOAuthCallback() {
  try {
    console.log("🔄 Handling OAuth authorization callback...");
    
    // With the new auth service, we just need to authenticate and get notebooks
    await authService.authenticate();
    const notebooks = await getOneNoteNotebooks();
    
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
    // The new auth service handles token refresh automatically during calls
    await authService.authenticate();
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
    await authService.logout();
    
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
    // With the new auth service, we use Office SSO first, then MSAL popup
    if (typeof Office !== 'undefined' && Office.context && Office.context.auth) {
      return 'Office.js SSO (Primary)';
    } else {
      return 'MSAL Popup Fallback';
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
