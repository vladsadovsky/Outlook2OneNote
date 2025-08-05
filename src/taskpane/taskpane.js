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
  - Call `testrun()` to display the current email subject.
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
    
    // Only show testrun button in development mode
    if (__DEV__) {
      document.getElementById("testrun").onclick = window.testRun;
    } else {
      // Hide testrun button in production
      const testrunButton = document.getElementById("testrun");
      if (testrunButton) {
        testrunButton.style.display = "none";
      }
    }
    
    document.getElementById("choose").onclick = chooseNotebook;
    document.getElementById("export").onclick = exportThread;
  }
});

// Development-only features - excluded in production builds
if (__DEV__) {
  
window.testRun = async function() {
  console.log("Outlook2OneNote::taskpane::testRun() ")

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
      
      // Get all items in the conversation using EWS
      await getConversationItems(conversationId, insertAt);
    } else {
      insertAt.appendChild(document.createTextNode("No conversation ID found - this might be a single email."));
    }
    
  } catch (error) {
    console.error("Error in testRun:", error);
    let insertAt = document.getElementById("item-subject");
    insertAt.innerHTML = "";
    insertAt.appendChild(document.createTextNode("Error: " + error.message));
  }
}

async function getConversationItems(conversationId, insertAt) {
  return new Promise((resolve, reject) => {
    // Use EWS (Exchange Web Services) to get conversation items
    const ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
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
      <m:SyncState />
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
        // Fallback: show current item info only
        insertAt.appendChild(document.createTextNode("Could not retrieve full thread. Showing current email only."));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("Reason: " + (result.error?.message || "Unknown error")));
        resolve();
      }
    });
  });
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
    showNotebookPopup(notebooks, onNotebookSelected);
  } else {
    insertAt.appendChild(document.createTextNode("No notebooks found."));
  } 
}

export async function exportThread() {

console.log("Outlook2OneNote::taskpane::exportThread() ")

  // Placeholder for exporting the email thread
  const insertAt = document.getElementById("item-subject");
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode("Export Thread functionality not implemented yet."));
}
