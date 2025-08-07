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

import { 
  authenticateAndGetNotebooks,
  getMockNotebooks
} from '../common/graphapi-auth.js';

// Function to get OneNote notebooks using Microsoft Graph API
export async function getOneNoteNotebooks() {
  try {
    console.log("Getting OneNote notebooks...");
    return await authenticateAndGetNotebooks();
  } catch (error) {
    console.error("Error in getOneNoteNotebooks:", error);
    // Return mock data for testing
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
