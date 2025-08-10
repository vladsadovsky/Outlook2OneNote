/* eslint-disable no-unused-vars */
/* global Office, __DEV__ */
/* eslint-disable no-console */

/*
 * Email Service Module - email-service.js
 * 
 * This module contains all email and conversation thread related fu    // Sort messages by date (client-side since we couldn't use $orderby in the API call)
    messages.sort((a, b) => {
      const dateA = new Date(a.receivedDateTime || 0);
      const dateB = new Date(b.receivedDateTime || 0);
      return dateA - dateB; // Ascending order (oldest first)
    });ality using
 * Microsoft Graph API exclusively with Office SSO-first authentication.
 * 
 * Key Features:
 * - Email thread retrieval via Microsoft Graph API only
 * - Conversation parsing and data extraction from Graph API responses
 * - Email content preparation for export
 * - Office SSO-first authentication with MSAL fallback
 * 
 * Dependencies:
 * - Microsoft Graph API
 * - Office SSO + MSAL Authentication Service
 */

import authService from '../common/auth-service.js';

// Main function to dump thread information (development only)
export async function dumpThread() {
  console.log("Outlook2OneNote::email-service::dumpThread()");

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
      threadLabel.appendChild(document.createTextNode("Email Thread (via Microsoft Graph API):"));
      insertAt.appendChild(threadLabel);
      insertAt.appendChild(document.createElement("br"));
      insertAt.appendChild(document.createElement("br"));
      
      // Show authentication status
      insertAt.appendChild(document.createTextNode("🔐 Checking authentication..."));
      insertAt.appendChild(document.createElement("br"));
      
      try {
        // Ensure we have authentication - Office SSO first, MSAL fallback
        console.log('🔐 Checking authentication status...');
        const isAuthenticated = authService.isAuthenticated();
        
        if (!isAuthenticated) {
          console.log('🔄 No valid authentication, starting auth flow...');
          insertAt.innerHTML = "";
          insertAt.appendChild(currentLabel);
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(threadLabel);
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createTextNode("🔐 Authenticating with Microsoft... Please complete the authentication process."));
          insertAt.appendChild(document.createElement("br"));
          
          // Trigger authentication (Office SSO first, then MSAL popup)
          await authService.authenticate();
          
          // Clear the authentication message
          insertAt.innerHTML = "";
          insertAt.appendChild(currentLabel);
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createTextNode(item.subject || "No subject"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(threadLabel);
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createElement("br"));
          insertAt.appendChild(document.createTextNode("✅ Authentication successful. Loading conversation..."));
          insertAt.appendChild(document.createElement("br"));
        }
        
        await getConversationItemsViaGraph(conversationId, insertAt);
      } catch (error) {
        console.error("Microsoft Graph API conversation retrieval failed:", error);
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("Could not retrieve conversation items via Microsoft Graph API."));
        insertAt.appendChild(document.createElement("br"));
        insertAt.appendChild(document.createTextNode("Error: " + error.message));
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

// Helper function to get conversation data for export using Microsoft Graph API
export async function getConversationDataForExport(conversationId) {
  console.log("Outlook2OneNote::email-service::getConversationDataForExport()");
  
  try {
    // Ensure we have authentication - use the same method as OneNote service
    const hasValidToken = await pkceAuth.hasValidToken();
    if (!hasValidToken) {
      console.log("🔄 No valid token, attempting authentication...");
      // Trigger authentication using the same method as OneNote service
      await pkceAuth.authenticateAndGetNotebooks();
    }
    
    const accessToken = pkceAuth.getAccessToken();
    if (!accessToken) {
      throw new Error("Failed to retrieve access token for Microsoft Graph API after authentication");
    }
    
    // Use Microsoft Graph API to get conversation messages
    const graphUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${conversationId}'&$select=subject,sender,from,receivedDateTime,sentDateTime,body,bodyPreview&$orderby=receivedDateTime asc&$top=100`;
    
    const response = await fetch(graphUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Microsoft Graph API request failed: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    
    if (data.error) {
      throw new Error(`Microsoft Graph API Error: ${data.error.code} - ${data.error.message}`);
    }
    
    const messages = data.value || [];
    
    // Transform messages to the expected format
    const conversationData = messages.map(message => {
      const senderName = message.sender?.emailAddress?.name || 
                        message.from?.emailAddress?.name || 
                        "Unknown sender";
      const senderEmail = message.sender?.emailAddress?.address || 
                         message.from?.emailAddress?.address || 
                         "";
      
      return {
        subject: message.subject || "No subject",
        senderName,
        senderEmail,
        date: new Date(message.receivedDateTime || message.sentDateTime || new Date()),
        body: message.body?.content || message.bodyPreview || "No content available",
        bodyType: message.body?.contentType || "text"
      };
    });
    
    console.log(`Retrieved ${conversationData.length} messages from conversation via Microsoft Graph API`);
    return conversationData;
    
  } catch (error) {
    console.error("Error retrieving conversation data via Microsoft Graph API:", error);
    throw error;
  }
}

// Function to get conversation items using Microsoft Graph API only
async function getConversationItemsViaGraph(conversationId, insertAt) {
  try {
    console.log("Outlook2OneNote::email-service::getConversationItemsViaGraph()");
    
    // Use Microsoft Graph API to get conversation messages via auth service
    console.log('📧 Fetching conversation messages via Graph API...');
    
    // URL encode the conversation ID to handle special characters
    const encodedConversationId = encodeURIComponent(conversationId);
    console.log('🔍 Original Conversation ID:', conversationId);
    console.log('🔍 Encoded Conversation ID:', encodedConversationId);
    
    // Try an even simpler approach - get recent messages without complex filtering
    // Microsoft Graph API has limitations on complex filters with conversationId
    console.log('🔄 Trying simplified query without conversationId filter...');
    const graphEndpoint = `/me/messages?$select=subject,from,receivedDateTime,bodyPreview,conversationId&$top=100`;
    console.log('🔗 Full Graph API URL:', `https://graph.microsoft.com/v1.0${graphEndpoint}`);
    
    console.log('📡 Making Graph API call...');
    const data = await authService.callGraphApi(graphEndpoint);
    
    if (!data || !data.value) {
      throw new Error('No conversation data received from Microsoft Graph API');
    }

    const allMessages = data.value;
    console.log(`📬 Retrieved ${allMessages.length} recent messages from Graph API`);
    
    // Filter messages by conversation ID client-side
    // Handle Base64 encoding differences: Graph API uses - where Office.js uses /
    const targetWithDashes = conversationId.replace(/\//g, '-');
    
    const messages = allMessages.filter(message => {
      const msgId = message.conversationId;
      // Direct comparison and dash-converted comparison
      return msgId === conversationId || msgId === targetWithDashes;
    });
    
    console.log(`🔍 Found ${messages.length} messages matching conversation ID`);
    
    if (messages.length === 0) {
      insertAt.appendChild(document.createTextNode(`No messages found for conversation ID: ${conversationId}. Retrieved ${allMessages.length} recent messages but none matched.`));
      return;
    }
    
    // Sort messages by date since we couldn't use $orderby in the API call
    messages.sort((a, b) => {
      const dateA = new Date(a.receivedDateTime || a.sentDateTime || 0);
      const dateB = new Date(b.receivedDateTime || b.sentDateTime || 0);
      return dateA - dateB; // Ascending order (oldest first)
    });
    
    console.log('📅 Messages sorted by date:', messages.map(m => ({
      subject: m.subject?.substring(0, 50),
      date: m.receivedDateTime || m.sentDateTime
    })));
    
    // Display each message in the conversation
    messages.forEach((message, index) => {
      const subject = message.subject || "No subject";
      const senderName = message.from?.emailAddress?.name || 
                        message.from?.emailAddress?.address ||
                        "Unknown sender";
      const dateStr = message.receivedDateTime;
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
        <strong>Date:</strong> ${formattedDate}<br>
        <strong>Preview:</strong> ${message.bodyPreview ? message.bodyPreview.substring(0, 100) + '...' : 'No preview available'}
      `;
      
      emailDiv.appendChild(emailInfo);
      insertAt.appendChild(emailDiv);
    });
    
    const summaryDiv = document.createElement("div");
    summaryDiv.style.marginTop = "15px";
    summaryDiv.style.fontWeight = "bold";
    summaryDiv.appendChild(document.createTextNode(`Total emails in thread (via Microsoft Graph API): ${messages.length}`));
    insertAt.appendChild(summaryDiv);
    
    console.log(`✅ Successfully retrieved ${messages.length} messages from conversation via Microsoft Graph API`);
    
  } catch (error) {
    console.error("❌ Error retrieving conversation via Microsoft Graph API:", error);
    
    // Display error to user
    insertAt.appendChild(document.createTextNode("❌ Error retrieving conversation via Microsoft Graph API: " + error.message));
    insertAt.appendChild(document.createElement("br"));
    
    if (error.message.includes("401") || error.message.includes("authentication")) {
      insertAt.appendChild(document.createTextNode("Please ensure you are authenticated. Click 'Choose Notebook' to authenticate."));
    } else if (error.message.includes("403")) {
      insertAt.appendChild(document.createTextNode("Access denied. Please ensure the app has the required permissions."));
    } else if (error.message.includes("429")) {
      insertAt.appendChild(document.createTextNode("Rate limit exceeded. Please try again in a few moments."));
    }
    
    throw error;
  }
}
