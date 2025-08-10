/* eslint-disable no-unused-vars */
/* global Office, __DEV__ */
/* eslint-disable no-console */
/* eslint-disable no-undef */

/* 
 * Task Pane Script - taskpane.js
  
 * This code is part of the Outlook2OneNote add-in.
 * It initializes the task pane and sets up event handlers for user interactions.
 * This code is executed when the Office environment is ready.
 * It checks if the host is Outlook and sets up the UI accordingly.
 *
 * This file contains only the main event handlers and Office.js initialization.
 * Business logic has been moved to separate service modules:
 * - onenote-service.js: OneNote notebook functionality
 * - email-service.js: Email thread and conversation functionality
 * 
 * Dependencies: 
 * - Office.js
 * - onenote-service.js
 * - email-service.js (for development features)
 * 
 * @requires Office.js
 * @author Your Name
 * @version 1.0
 */

// Import service modules
import { 
  getOneNoteNotebooks, 
  showNotebookPopup, 
  onNotebookSelected,
  exportConversationToOneNote,
  exportSingleEmailToOneNote
} from './onenote-service.js';

import {
  getSelectedNotebook,
  setSelectedNotebook,
  clearSelectedNotebook
} from '../common/app-state.js';

import { dumpThread, getConversationDataForExport } from './email-service.js';

// Office.js initialization
Office.onReady((info) => {
  console.log("Outlook2OneNote::Office.onReady");

  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Only show dumpthread button in development mode
    if (typeof __DEV__ !== 'undefined' && __DEV__) {
      document.getElementById("dumpthread").onclick = window.dumpThread;
    } else {
      // Hide dumpthread button in production
      const dumpthreadButton = document.getElementById("dumpthread");
      if (dumpthreadButton) {
        dumpthreadButton.style.display = "none";
      }
    }
    
    document.getElementById("choose").onclick = chooseNotebook;
    document.getElementById("export").onclick = exportThread;
  }
});

// Development-only features - excluded in production builds
if (__DEV__) {
  // Expose dumpThread to global scope for development
  window.dumpThread = dumpThread;
}

// Event handler: Choose OneNote notebook
export async function chooseNotebook() {
  console.log("Outlook2OneNote::taskpane::chooseNotebook()");

  // Display a popup listing available notebooks from OneNote
  const insertAt = document.getElementById("item-subject");

  try {
    // Show loading state
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("🔐 Starting secure authentication..."));
    
    const notebooks = await getOneNoteNotebooks();  
    if (notebooks && notebooks.length > 0) {
      insertAt.innerHTML = "";
      insertAt.appendChild(document.createTextNode(`Found ${notebooks.length} OneNote notebooks. Select one from the popup.`));
      
      showNotebookPopup(notebooks, (notebook) => {
        // Store the selected notebook globally
        setSelectedNotebook(notebook);
        console.log("Selected notebook stored:", notebook);
        
        // Call the original callback
        onNotebookSelected(notebook);
        
        // Update UI to show selected notebook
        insertAt.innerHTML = "";
        insertAt.appendChild(document.createTextNode(`✅ Selected Notebook: ${notebook.displayName || notebook.name || 'Unknown'}`));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("You can now use 'Export Thread' to export emails to this notebook."));
      });
    } else {
      insertAt.innerHTML = "";
      insertAt.appendChild(document.createTextNode("📭 No notebooks found or authentication cancelled."));
    } 
  } catch (error) {
    console.error("Error in chooseNotebook:", error);
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("❌ Error retrieving notebooks: " + error.message));
  }
}

// Event handler: Export email thread to OneNote
export async function exportThread() {
  console.log("Outlook2OneNote::taskpane::exportThread()");

  const insertAt = document.getElementById("item-subject");
  
  // Check if a notebook has been selected
  const selectedNotebook = getSelectedNotebook();
  if (!selectedNotebook) {
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("Please select a notebook first using 'Choose Notebook' button."));
    return;
  }

  try {
    // Get current email item
    const item = Office.context.mailbox.item;
    const conversationId = item.conversationId;
    
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("Exporting email thread to OneNote..."));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(`Target Notebook: ${selectedNotebook.displayName || selectedNotebook.name || 'Unknown'}`));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    if (conversationId) {
      // Get conversation items for export
      const conversationData = await getConversationDataForExport(conversationId);
      
      if (conversationData && conversationData.length > 0) {
        // Export to OneNote
        await exportConversationToOneNote(conversationData, selectedNotebook, insertAt);
      } else {
        insertAt.appendChild(document.createTextNode("No conversation data found to export."));
      }
    } else {
      // Export single email
      await exportSingleEmailToOneNote(item, selectedNotebook, insertAt);
    }
    
  } catch (error) {
    console.error("Error in exportThread:", error);
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("Error during export: " + error.message));
  }
}
