/* eslint-disable no-unused-vars */
/* global Office, testun, chooseNotebook, exportThread, onNotebookSelected, getOneNoteNotebooks, showNotebookPopup, __DEV__ */
/* eslint-disable no-console */
/* eslint-disable no-undef */

/* 
 * Task Pane Script  taskpane.js
  
 * This code is part of the Outlook2OneNote add-in.
 * It initializes the task pane and sets up event handlers for user interactions.
 * * This code is executed when the Office environment is ready.
 * It checks if the host is Outlook and sets up the UI accordingly.
 *
  This file contains the logic for the task pane of the Outlook2OneNote add-in.
  It handles user interactions such as test running the add-in, choosing a notebook, and exporting email threads.
  
  Note: This code assumes that you have already set up the necessary Office.js and Microsoft Graph API configurations.
  
  Dependencies: 
  - Office.js
  - Microsoft Graph API (for OneNote notebooks)
  
  Global variables: 
  - info: Contains information about the Office host environment.
  - document: The global document object for manipulating the DOM.
  
  Usage:
  - Call `dumpthread()` to display the current email subject.
  - Call `chooseNotebook()` to display a list of OneNote notebooks.
  - Call `exportThread()` to export the email thread (currently a placeholder).
  
  Note: Ensure that you have the necessary permissions and access tokens for Microsoft Graph API calls.
  
  @requires Office.js
  @requires Microsoft Graph API
  @author Your Name
  @version 1.0
*/


Office.onReady((info) => {
  console.log("Outlook2OneNote::Office.onReady ")

  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Only show dumpthread button in development mode
    if (__DEV__) {
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

// Global variable to store the selected OneNote notebook
let selectedNotebook = null;

// Helper function to get the currently selected notebook
function getSelectedNotebook() {
  return selectedNotebook;
}

// Helper function to clear the selected notebook
function clearSelectedNotebook() {
  selectedNotebook = null;
}

// Function to get OneNote notebooks using Microsoft Graph API
async function getOneNoteNotebooks() {
  try {
    console.log("Getting OneNote notebooks...");
    
    // Get access token for Microsoft Graph API
    return new Promise((resolve, reject) => {
      Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, async (tokenResult) => {
        if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
          try {
            const accessToken = tokenResult.value;
            
            // Make request to Microsoft Graph API to get OneNote notebooks
            const graphUrl = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks';
            
            const response = await fetch(graphUrl, {
              method: 'GET',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
                'Content-Type': 'application/json'
              }
            });
            
            if (!response.ok) {
              throw new Error(`Graph API request failed: ${response.status} ${response.statusText}`);
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
              
              console.log(`Found ${notebooks.length} OneNote notebooks`);
              resolve(notebooks);
            } else {
              console.log("No OneNote notebooks found");
              resolve([]);
            }
            
          } catch (fetchError) {
            console.error("Error fetching OneNote notebooks:", fetchError);
            reject(new Error(`Failed to retrieve notebooks: ${fetchError.message}`));
          }
        } else {
          console.error("Failed to get access token:", tokenResult.error);
          
          // Try fallback method with Office.js callback token
          tryFallbackNotebookRetrieval()
            .then(notebooks => resolve(notebooks))
            .catch(fallbackError => {
              reject(new Error(`Authentication failed: ${tokenResult.error?.message || 'Unknown error'}`));
            });
        }
      });
    });
    
  } catch (error) {
    console.error("Error in getOneNoteNotebooks:", error);
    throw error;
  }
}

// Fallback method to get OneNote notebooks using Office.js callback token
async function tryFallbackNotebookRetrieval() {
  return new Promise((resolve, reject) => {
    console.log("Trying fallback method to get OneNote notebooks...");
    
    // Try using Office.context.mailbox.getCallbackTokenAsync for Graph API access
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (tokenResult) => {
      if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
        try {
          const accessToken = tokenResult.value;
          const graphUrl = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks';
          
          const response = await fetch(graphUrl, {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Accept': 'application/json'
            }
          });
          
          if (!response.ok) {
            throw new Error(`Fallback Graph API request failed: ${response.status} ${response.statusText}`);
          }
          
          const data = await response.json();
          
          if (data.value && data.value.length > 0) {
            const notebooks = data.value.map(notebook => ({
              id: notebook.id,
              name: notebook.displayName,
              displayName: notebook.displayName,
              createdDateTime: notebook.createdDateTime,
              lastModifiedDateTime: notebook.lastModifiedDateTime,
              isDefault: notebook.isDefault || false
            }));
            
            console.log(`Fallback method found ${notebooks.length} OneNote notebooks`);
            resolve(notebooks);
          } else {
            resolve([]);
          }
          
        } catch (fetchError) {
          console.error("Fallback method failed:", fetchError);
          // Return mock data for testing if both methods fail
          resolve([
            {
              id: 'mock-notebook-1',
              name: 'Test Notebook (Mock Only - not persisted)',
              displayName: 'Test Notebook (Mock Only not persisted )',
              isDefault: true,
              createdDateTime: new Date().toISOString(),
              lastModifiedDateTime: new Date().toISOString()
            }
          ]);
        }
      } else {
        console.error("fallack token retrieval failed:", tokenResult.error);
        // Return mock data for testing
        resolve([
          {
            id: 'mock-notebook-2',
            name: 'Test Notebook (Mock 2 - not persisted)',
            displayName: 'Test Notebook (Mock 2 - not persisted)',
            isDefault: true,
            createdDateTime: new Date().toISOString(),
            lastModifiedDateTime: new Date().toISOString()
          }
        ]);
      }
    });
  });
}

// Function to show notebook selection popup
function showNotebookPopup(notebooks, onNotebookSelected) {
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

// Placeholder function for onNotebookSelected callback
function onNotebookSelected(notebook) {
  console.log("Notebook selected (default callback):", notebook);
  
  const insertAt = document.getElementById("item-subject");
  insertAt.innerHTML = "";
  insertAt.appendChild(document.createTextNode(`Selected: ${notebook.displayName || notebook.name}`));
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode("Notebook selection complete."));
}

// Development-only features - excluded in production builds
if (__DEV__) {
  
window.dumpThread = async function() {
  console.log("Outlook2OneNote::taskpane::dumpThread() ")

  try {
    const item = Office.context.mailbox.item;
    let insertAt = document.getElementById("item-subject");
    
    // Clear previous content
    insertAt.innerHTML = "";
    
    // Display current item subject
    let currentLabel = document.createElement("b");
    currentLabel.appendChild(document.createTextNode("Current Email Subject: "));
    insertAt.appendChild(currentLabel);
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
    insertAt.appendChild(document.createElement("br"));
    insertAt.appendChild(document.createElement("br"));
    
    // Get conversation ID to retrieve the whole thread
    const conversationId = item.conversationId;
    console.log("Conversation ID:", conversationId);
    
    if (conversationId) {
      // Display thread header
      let threadLabel = document.createElement("b");
      threadLabel.appendChild(document.createTextNode("Email Thread Subjects:"));
      insertAt.appendChild(threadLabel);
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createElement("br"));
      
      // Try multiple approaches to get conversation items
      try {
        await getConversationItems(conversationId, insertAt);
      } catch (error) {
        console.error("All conversation retrieval methods failed:", error);
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("Could not retrieve conversation items. This may be due to:"));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("1. EWS is disabled on this Exchange server"));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("2. Insufficient permissions for conversation access"));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("3. Network or authentication issues"));
      }
    } else {
      insertAt.appendChild(document.createTextNode("No conversation ID found - this might be a single email."));
    }
    
  } catch (error) {
    console.error("Error in dumpThread:", error);
    let insertAt = document.getElementById("item-subject");
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("Error: " + error.message));
  }
}

async function getConversationItems(conversationId, insertAt) {
  return new Promise((resolve, reject) => {
    // Use EWS (Exchange Web Services) to get conversation items
    // Simplified request without impersonation to avoid 403 errors
    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="message:Sender" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ConversationId Id="${conversationId}" />
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        try {
          console.log("EWS Response:", result.value);
          parseConversationResponse(result.value, insertAt);
          resolve();
        } catch (parseError) {
          console.error("Error parsing EWS response:", parseError);
          insertAt.appendChild(document.createTextNode("Error parsing conversation data."));
          reject(parseError);
        }
      } else {
        console.error("EWS request failed:", result.error);
        if (result.error?.message?.includes("403")) {
          console.log("Got 403 error, trying simple EWS method...");
          trySimpleEwsRequest(conversationId, insertAt)
            .then(() => resolve())
            .catch((simpleError) => {
              console.error("Simple EWS also failed:", simpleError);
              // Now try REST API
              tryRestApiFallback(conversationId, insertAt)
                .then(() => resolve())
                .catch((restError) => {
                  console.error("REST API also failed:", restError);
                  // Final fallback: try Office.js conversation API
                  tryOfficeConversationApi(conversationId, insertAt)
                    .then(() => resolve())
                    .catch((officeError) => {
                      console.error("All methods failed");
                      insertAt.appendChild(document.createTextNode("Could not retrieve full thread. All methods failed."));
                      insertAt.appendChild(document.createElement("br"));
                      insertAt.appendChild(document.createTextNode("EWS Error: " + (result.error?.message || "Unknown error")));
                      insertAt.appendChild(document.createElement("br"));
                      insertAt.appendChild(document.createTextNode("REST Error: " + restError.message));
                      insertAt.appendChild(document.createElement("br"));
                      insertAt.appendChild(document.createTextNode("Office API Error: " + officeError.message));
                      resolve();
                    });
                });
            });
        } else {
          // Fallback: Try using REST API instead of EWS
          tryRestApiFallback(conversationId, insertAt)
            .then(() => resolve())
            .catch((restError) => {
              console.error("REST API fallback also failed:", restError);
              // Final fallback: try Office.js conversation API
              tryOfficeConversationApi(conversationId, insertAt)
                .then(() => resolve())
                .catch((officeError) => {
                  console.error("All fallback methods failed");
                  insertAt.appendChild(document.createTextNode("Could not retrieve full thread. All methods failed."));
                  insertAt.appendChild(document.createElement("br"));
                  insertAt.appendChild(document.createTextNode("EWS Error: " + (result.error?.message || "Unknown error")));
                  insertAt.appendChild(document.createElement("br"));
                  insertAt.appendChild(document.createTextNode("REST Error: " + restError.message));
                  insertAt.appendChild(document.createElement("br"));
                  insertAt.appendChild(document.createTextNode("Office API Error: " + officeError.message));
                  resolve();
                });
            });
        }
      }
    });
  });
}

async function trySimpleEwsRequest(conversationId, insertAt) {
  return new Promise((resolve, reject) => {
    // Try a simpler EWS request without Exchange Impersonation
    const simpleEwsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ConversationId Id="${conversationId}" />
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

    console.log("Trying simple EWS request without impersonation...");
    
    Office.context.mailbox.makeEwsRequestAsync(simpleEwsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        try {
          console.log("Simple EWS Response:", result.value);
          parseConversationResponse(result.value, insertAt);
          resolve();
        } catch (parseError) {
          console.error("Error parsing simple EWS response:", parseError);
          reject(parseError);
        }
      } else {
        console.error("Simple EWS request also failed:", result.error);
        reject(new Error(result.error?.message || "Simple EWS failed"));
      }
    });
  });
}

async function tryRestApiFallback(conversationId, insertAt) {
  return new Promise((resolve, reject) => {
    // Try using Office.js REST API as fallback
    console.log("Attempting REST API fallback for conversation:", conversationId);
    
    // Use Office.context.mailbox.getCallbackTokenAsync to get token for REST calls
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (tokenResult) => {
      if (tokenResult.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = tokenResult.value;
        const restUrl = Office.context.mailbox.restUrl;
        
        // Construct REST API URL for conversation
        // Use Microsoft Graph API endpoint instead of Exchange REST API
        const conversationUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'&$select=subject,sender,receivedDateTime,sentDateTime,from&$orderby=receivedDateTime&$top=50`;
        
        fetch(conversationUrl, {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Accept': 'application/json'
          }
        })
        .then(response => {
          if (!response.ok) {
            throw new Error(`REST API request failed: ${response.status} ${response.statusText}`);
          }
          return response.json();
        })
        .then(data => {
          console.log("REST API Response:", data);
          if (data.error) {
            console.error("Graph API returned error:", data.error);
            throw new Error(`Graph API Error: ${data.error.code} - ${data.error.message}`);
          }
          parseRestApiResponse(data, insertAt);
          resolve();
        })
        .catch(error => {
          console.error("REST API request failed:", error);
          reject(error);
        });
      } else {
        console.error("Failed to get callback token:", tokenResult.error);
        reject(new Error("Could not get authentication token for REST API"));
      }
    });
  });
}

async function tryOfficeConversationApi(conversationId, insertAt) {
  return new Promise((resolve, reject) => {
    // Final fallback: use Office.js to get basic conversation info
    console.log("Attempting Office.js conversation API fallback...");
    
    try {
      // Get current item and display basic info about the conversation
      const item = Office.context.mailbox.item;
      
      insertAt.appendChild(document.createTextNode("Using Office.js fallback method:"));
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createElement("br"));
      
      // Display current email info
      const emailDiv = document.createElement("div");
      emailDiv.style.marginBottom = "10px";
      emailDiv.style.padding = "8px";
      emailDiv.style.border = "1px solid #ccc";
      emailDiv.style.borderRadius = "4px";
      
      const emailInfo = document.createElement("div");
      emailInfo.innerHTML = `
        <strong>Current Email:</strong><br>
        <strong>Subject:</strong> ${item.subject || "No subject"}<br>
        <strong>Conversation ID:</strong> ${conversationId}<br>
        <strong>Note:</strong> Full conversation details unavailable due to server restrictions
      `;
      
      emailDiv.appendChild(emailInfo);
      insertAt.appendChild(emailDiv);
      
      // Try to get additional conversation info if available
      if (item.conversationId) {
        const infoDiv = document.createElement("div");
        infoDiv.style.marginTop = "10px";
        infoDiv.style.fontStyle = "italic";
        infoDiv.appendChild(document.createTextNode(
          "This email is part of a conversation thread, but detailed thread information " +
          "cannot be retrieved due to server permissions or configuration."
        ));
        insertAt.appendChild(infoDiv);
      }
      
      resolve();
      
    } catch (error) {
      console.error("Office.js fallback also failed:", error);
      reject(error);
    }
  });
}

function parseRestApiResponse(data, insertAt) {
  try {
    const messages = data.value || [];
    
    if (messages.length === 0) {
      insertAt.appendChild(document.createTextNode("No conversation items found via REST API."));
      return;
    }
    
    // Sort messages by date
    messages.sort((a, b) => new Date(a.DateTimeReceived || a.DateTimeSent) - new Date(b.DateTimeReceived || b.DateTimeSent));
    
    messages.forEach((message, index) => {
      const subject = message.subject || "No subject";
      const senderName = message.sender?.emailAddress?.name || message.from?.emailAddress?.name || "Unknown sender";
      const dateStr = message.receivedDateTime || message.sentDateTime;
      const formattedDate = dateStr ? new Date(dateStr).toLocaleString() : "Unknown date";
      
      // Display the email info
      const emailDiv = document.createElement("div");
      emailDiv.style.marginBottom = "10px";
      emailDiv.style.padding = "8px";
      emailDiv.style.border = "1px solid #ccc";
      emailDiv.style.borderRadius = "4px";
      
      const emailInfo = document.createElement("div");
      emailInfo.innerHTML = `
        <strong>${index + 1}. Subject:</strong> ${subject}<br>
        <strong>From:</strong> ${senderName}<br>
        <strong>Date:</strong> ${formattedDate}
      `;
      
      emailDiv.appendChild(emailInfo);
      insertAt.appendChild(emailDiv);
    });
    
    const summaryDiv = document.createElement("div");
    summaryDiv.style.marginTop = "15px";
    summaryDiv.style.fontWeight = "bold";
    summaryDiv.appendChild(document.createTextNode(`Total emails in thread (via REST API): ${messages.length}`));
    insertAt.appendChild(summaryDiv);
    
  } catch (error) {
    console.error("Error parsing REST API response:", error);
    insertAt.appendChild(document.createTextNode("Error parsing REST API response: " + error.message));
  }
}

function parseConversationResponse(ewsResponse, insertAt) {
  try {
    // Parse the XML response
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(ewsResponse, "text/xml");
    
    // Look for conversation items in the response
    const conversations = xmlDoc.getElementsByTagName("t:Conversation");
    
    if (conversations.length === 0) {
      insertAt.appendChild(document.createTextNode("No conversation items found in response."));
      return;
    }
    
    let itemCount = 0;
    
    // Iterate through conversations
    for (let i = 0; i < conversations.length; i++) {
      const conversation = conversations[i];
      const conversationNodes = conversation.getElementsByTagName("t:ConversationNode");
      
      // Process each conversation node (email in the thread)
      for (let j = 0; j < conversationNodes.length; j++) {
        const node = conversationNodes[j];
        const items = node.getElementsByTagName("t:Items")[0];
        
        if (items) {
          const messageElements = items.getElementsByTagName("t:Message");
          
          for (let k = 0; k < messageElements.length; k++) {
            const message = messageElements[k];
            itemCount++;
            
            // Extract subject
            const subjectElement = message.getElementsByTagName("t:Subject")[0];
            const subject = subjectElement ? subjectElement.textContent : "No subject";
            
            // Extract sender
            const senderElement = message.getElementsByTagName("t:Sender")[0];
            let senderName = "Unknown sender";
            if (senderElement) {
              const nameElement = senderElement.getElementsByTagName("t:Name")[0];
              if (nameElement) {
                senderName = nameElement.textContent;
              }
            }
            
            // Extract date
            const dateElement = message.getElementsByTagName("t:DateTimeReceived")[0] || 
                               message.getElementsByTagName("t:DateTimeSent")[0];
            let dateStr = "Unknown date";
            if (dateElement) {
              const date = new Date(dateElement.textContent);
              dateStr = date.toLocaleString();
            }
            
            // Display the email info
            const emailDiv = document.createElement("div");
            emailDiv.style.marginBottom = "10px";
            emailDiv.style.padding = "8px";
            emailDiv.style.border = "1px solid #ccc";
            emailDiv.style.borderRadius = "4px";
            
            const emailInfo = document.createElement("div");
            emailInfo.innerHTML = `
              <strong>${itemCount}. Subject:</strong> ${subject}<br>
              <strong>From:</strong> ${senderName}<br>
              <strong>Date:</strong> ${dateStr}
            `;
            
            emailDiv.appendChild(emailInfo);
            insertAt.appendChild(emailDiv);
          }
        }
      }
    }
    
    if (itemCount === 0) {
      insertAt.appendChild(document.createTextNode("No email items found in the conversation."));
    } else {
      const summaryDiv = document.createElement("div");
      summaryDiv.style.marginTop = "15px";
      summaryDiv.style.fontWeight = "bold";
      summaryDiv.appendChild(document.createTextNode(`Total emails in thread: ${itemCount}`));
      insertAt.appendChild(summaryDiv);
    }
    
  } catch (error) {
    console.error("Error parsing conversation XML:", error);
    insertAt.appendChild(document.createTextNode("Error parsing conversation data: " + error.message));
  }
}

} // End of __DEV__ block

// Production functions - always available
export async function chooseNotebook() {

console.log("Outlook2OneNote::taskpane::chooseNotebook() ")

// Display a popup listing available notebooks from OneNote
  const insertAt = document.getElementById("item-subject");

  const notebooks = await getOneNoteNotebooks();  
  if (notebooks && notebooks.length > 0) {
    showNotebookPopup(notebooks, (notebook) => {
      // Store the selected notebook globally
      selectedNotebook = notebook;
      console.log("Selected notebook stored:", selectedNotebook);
      
      // Call the original callback
      onNotebookSelected(notebook);
      
      // Update UI to show selected notebook
      insertAt.innerHTML = "";
      insertAt.appendChild(document.createTextNode(`Selected Notebook: ${notebook.displayName || notebook.name || 'Unknown'}`));
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createTextNode("You can now use 'Export Thread' to export emails to this notebook."));
    });
  } else {
    insertAt.appendChild(document.createTextNode("No notebooks found."));
  } 
}

export async function exportThread() {

console.log("Outlook2OneNote::taskpane::exportThread() ")

  const insertAt = document.getElementById("item-subject");
  
  // Check if a notebook has been selected
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

// Helper function to get conversation data for export
async function getConversationDataForExport(conversationId) {
  return new Promise((resolve, reject) => {
    console.log("Getting conversation data for export...");
    
    // Try to get conversation data using the same methods as dumpThread
    const conversationData = [];
    
    // Use simplified EWS request
    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="message:Sender" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="item:Body" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ConversationId Id="${conversationId}" />
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        try {
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(result.value, "text/xml");
          const conversations = xmlDoc.getElementsByTagName("t:Conversation");
          
          for (let i = 0; i < conversations.length; i++) {
            const conversation = conversations[i];
            const conversationNodes = conversation.getElementsByTagName("t:ConversationNode");
            
            for (let j = 0; j < conversationNodes.length; j++) {
              const node = conversationNodes[j];
              const items = node.getElementsByTagName("t:Items")[0];
              
              if (items) {
                const messageElements = items.getElementsByTagName("t:Message");
                
                for (let k = 0; k < messageElements.length; k++) {
                  const message = messageElements[k];
                  
                  const subjectElement = message.getElementsByTagName("t:Subject")[0];
                  const subject = subjectElement ? subjectElement.textContent : "No subject";
                  
                  const senderElement = message.getElementsByTagName("t:Sender")[0];
                  let senderName = "Unknown sender";
                  let senderEmail = "";
                  if (senderElement) {
                    const nameElement = senderElement.getElementsByTagName("t:Name")[0];
                    const emailElement = senderElement.getElementsByTagName("t:EmailAddress")[0];
                    if (nameElement) senderName = nameElement.textContent;
                    if (emailElement) senderEmail = emailElement.textContent;
                  }
                  
                  const dateElement = message.getElementsByTagName("t:DateTimeReceived")[0] || 
                                     message.getElementsByTagName("t:DateTimeSent")[0];
                  let dateStr = new Date().toISOString();
                  if (dateElement) {
                    dateStr = dateElement.textContent;
                  }
                  
                  const bodyElement = message.getElementsByTagName("t:Body")[0];
                  let body = "No content available";
                  if (bodyElement) {
                    body = bodyElement.textContent;
                  }
                  
                  conversationData.push({
                    subject,
                    senderName,
                    senderEmail,
                    date: new Date(dateStr),
                    body
                  });
                }
              }
            }
          }
          
          resolve(conversationData);
        } catch (parseError) {
          console.error("Error parsing conversation for export:", parseError);
          reject(parseError);
        }
      } else {
        // Fallback: get current item only
        const item = Office.context.mailbox.item;
        resolve([{
          subject: item.subject || "No subject",
          senderName: "Current User",
          senderEmail: Office.context.mailbox.userProfile.emailAddress,
          date: new Date(),
          body: "Email body content not available"
        }]);
      }
    });
  });
}

// Helper function to export conversation to OneNote
async function exportConversationToOneNote(conversationData, notebook, insertAt) {
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
async function exportSingleEmailToOneNote(item, notebook, insertAt) {
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
