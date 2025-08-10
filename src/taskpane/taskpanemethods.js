/* eslint-disable no-unused-vars */
/* global Office */
/* eslint-disable no-console */

/**
 * TaskPane Methods for Outlook2OneNote Add-in with PKCE Authentication
 */

import { getOneNoteNotebooks, showNotebookPopup, onNotebookSelected, exportThread } from './onenote-service.js';
import { getSelectedNotebook, setSelectedNotebook, clearSelectedNotebook } from '../common/app-state.js';
import { pkceAuth } from '../common/pkce-auth.js';

/**
 * Choose OneNote notebook using PKCE authentication
 */
async function chooseNotebook() {
  try {
    console.log("Outlook2OneNote::taskpane::chooseNotebook()");
    
    // Update UI to show loading state
    const outputElement = document.getElementById("item-subject");
    outputElement.textContent = "🔐 Starting secure authentication...";
    
    // Use the OneNote service to get notebooks with authentication
    const notebooks = await getOneNoteNotebooks();
    
    if (notebooks && notebooks.length > 0) {
      console.log(`✅ Found ${notebooks.length} notebooks`);
      outputElement.textContent = `Found ${notebooks.length} OneNote notebooks`;
      showNotebookPopup(notebooks, (notebook) => {
        setSelectedNotebook(notebook.id, notebook.displayName);
        outputElement.textContent = `Selected: ${notebook.displayName}`;
      });
    } else {
      console.log("📭 No notebooks found");
      outputElement.textContent = "No OneNote notebooks found. Please create a notebook in OneNote first.";
    }
    
  } catch (error) {
    console.error("❌ Choose notebook failed:", error);
    const outputElement = document.getElementById("item-subject");
    outputElement.textContent = `❌ Authentication failed: ${error.message}`;
  }
}

/**
 * Export current thread to OneNote
 */
async function exportToOneNote() {
  try {
    console.log("📤 Starting export to OneNote...");
    
    const selectedNotebook = getSelectedNotebook();
    if (!selectedNotebook || !selectedNotebook.id) {
      document.getElementById("item-subject").textContent = "❌ Please select a OneNote notebook first";
      return;
    }
    
    // Update UI
    const outputElement = document.getElementById("item-subject");
    outputElement.textContent = `📤 Exporting to ${selectedNotebook.name}...`;
    
    // Use the export functionality from onenote-service
    const result = await exportThread(selectedNotebook);
    
    if (result && result.success) {
      outputElement.textContent = `✅ Successfully exported to ${selectedNotebook.name}`;
      console.log("✅ Export completed successfully");
    } else {
      outputElement.textContent = `❌ Export failed: ${result?.error || 'Unknown error'}`;
      console.error("❌ Export failed:", result);
    }
    
  } catch (error) {
    console.error("❌ Export to OneNote failed:", error);
    document.getElementById("item-subject").textContent = `❌ Export failed: ${error.message}`;
  }
}

/**
 * Initialize task pane methods when Office is ready
 */
Office.onReady(() => {
  console.log("Outlook2OneNote::Office.onReady");
  
  if (Office.context.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Attach event handlers to buttons using the correct IDs from HTML
    document.getElementById("choose").onclick = chooseNotebook;
    document.getElementById("export").onclick = exportToOneNote;
    
    console.log("✅ Task pane methods initialized");
  }
});

// Export functions for use in other modules
export { 
  chooseNotebook, 
  exportToOneNote 
};
