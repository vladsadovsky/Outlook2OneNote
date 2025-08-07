/* eslint-disable no-unused-vars */
/* global Office, __DEV__ */
/* eslint-disable no-console */

/*
 * Email Service Module - email-service.js
 * 
 * This module contains all email and conversation thread related functionality including:
 * - Email thread retrieval via EWS and REST API
 * - Conversation parsing and data extraction
 * - Email content preparation for export
 * - Fallback methods for different Exchange configurations
 * 
 * Dependencies:
 * - Office.js
 * - Exchange Web Services (EWS)
 * - Microsoft Graph API (fallback)
 */

// Main function to dump thread information (development only)
export async function dumpThread() {
  console.log("Outlook2OneNote::taskpane::dumpThread()");

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

// Helper function to get conversation data for export
export async function getConversationDataForExport(conversationId) {
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

// Function to get conversation items using multiple fallback methods
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

// Fallback method 1: Simple EWS request
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

// Fallback method 2: REST API
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

// Fallback method 3: Office.js conversation API
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

// Helper function to parse EWS conversation response
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

// Helper function to parse REST API response
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
