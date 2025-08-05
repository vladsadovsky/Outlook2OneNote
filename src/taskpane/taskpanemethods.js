const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID_HERE",
    redirectUri: "https://localhost:3000/taskpane.html"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let accessToken = null;
let selectedNotebookId = null;

async function getToken() {
  if (accessToken) return accessToken;
  const result = await msalInstance.loginPopup({
    scopes: ["User.Read", "Mail.Read", "Notes.ReadWrite"]
  });
  accessToken = result.accessToken;
  return accessToken;
}


/** * Function to handle notebook selection
 * @param notebookId {string} - The ID of the selected notebook
 * @param notebookName {string} - The name of the selected notebook
 */
function onNotebookSelected(notebookId, notebookName) {
  console.log(`Selected Notebook: ${notebookName} (ID: ${notebookId})`);
  // TODO:  implement further logic to handle the selected notebook
  // For example, you might want to store the notebookId for later use when exporting the email thread.
  const insertAt = document.getElementById("item-subject");
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(`Selected Notebook: ${notebookName} (ID: ${notebookId})`));
  // You can also call a function to proceed with the export using this notebookId. 
}


/** * Method that creates a popup (modal dialog) displaying the list of available OneNote notebooks retrieved via the Microsoft Graph API, allowing the user to select one
 * @param accessToken {string} - The access token for Microsoft Graph API
 * @returns {Promise<void>} - A promise that resolves when the popup is displayed
 * @example
 *
 * Note: 
 * This function assumes you have a function getOneNoteNotebooks(accessToken) that fetches the list of notebooks.  
 * 
 * Prerequisites
 *  - You have already authenticated and have an accessToken from MSAL
 *  - You’ve included this somewhere in taskpane.js or your command script
 *  - You're serving over HTTPS and your app is registered in Azure AD with Notes.ReadWrite
*/
export 
async function getOneNoteNotebooks(accessToken) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me/onenote/notebooks", {
    headers: {
      "Authorization": `Bearer ${accessToken}`
    }
  });
  if (!response.ok) {
    console.error("Error fetching OneNote notebooks:", response.statusText);
    return [];
  }
  const data = await response.json();
  return data.value || [];
} 

async function showNotebookPopup(accessToken) {
  const notebooks = await getOneNoteNotebooks(accessToken);
  if (!notebooks || notebooks.length === 0) {
    alert("No notebooks found.");
    return;
  }

  // Create popup container
  const popup = document.createElement("div");
  popup.style.position = "fixed";
  popup.style.top = "20%";
  popup.style.left = "35%";
  popup.style.width = "30%";
  popup.style.padding = "20px";
  popup.style.backgroundColor = "#fff";
  popup.style.border = "1px solid #ccc";
  popup.style.boxShadow = "0 0 20px rgba(0,0,0,0.2)";
  popup.style.zIndex = "10000";
  popup.style.borderRadius = "8px";

  const title = document.createElement("h3");
  title.textContent = "Select a OneNote Notebook:";
  popup.appendChild(title);

  const list = document.createElement("ul");
  list.style.listStyleType = "none";
  list.style.padding = 0;

  notebooks.forEach(nb => {
    const item = document.createElement("li");
    item.style.marginBottom = "8px";
    const btn = document.createElement("button");
    btn.textContent = nb.displayName;
    btn.style.width = "100%";
    btn.style.padding = "6px";
    btn.style.border = "1px solid #888";
    btn.style.borderRadius = "4px";
    btn.style.backgroundColor = "#f5f5f5";
    btn.onclick = () => {
      document.body.removeChild(popup);
      onNotebookSelected(nb.id, nb.displayName);
    };
    item.appendChild(btn);
    list.appendChild(item);
  });

  popup.appendChild(list);
  document.body.appendChild(popup);
}




async function chooseNotebook() {
  const token = await getToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me/onenote/notebooks", {
    headers: { Authorization: `Bearer ${token}` }
  });
  const data = await res.json();
  if (data.value && data.value.length > 0) {
    selectedNotebookId = data.value[0].id;
    document.getElementById("output").textContent = "Notebook chosen: " + data.value[0].displayName;
  } else {
    document.getElementById("output").textContent = "No notebooks found.";
  }
}


/**
 *  Function to export the current email thread to OneNote  
 * * @returns {Promise<void>} - A promise that resolves when the export is complete
 * * @description This function retrieves the current email thread and exports it to a OneNote notebook.
 * * It uses the Microsoft Graph API to fetch the conversation messages and create a new OneNote page for each message.
 * * @example
 * *  exportThread();
 * * * Note:
 *  
 * This function assumes you have already authenticated and have an access token from MSAL. 
 * You should also have a selected notebook ID stored in `selectedNotebookId` before calling this function.
 * * Ensure you have the necessary permissions (Notes.ReadWrite) in your Azure AD app registration.
 * * @requires Microsoft Graph API
 * * @author Your Name
 * * @version 1.0
 * * @see https://docs.microsoft.com/en-us/graph/api/resources/onenote?view=graph-rest-1.0
 * * @see https://docs.microsoft.com/en-us/graph/api/conversation-list-messages?view=graph-rest-1.0
 * * @see https://docs.microsoft.com/en-us/graph/api/section-post-pages?view=graph-rest-1.0
 * * @see https://docs.microsoft.com/en-us/graph/api/section-post-pages?view=graph-rest-1.0#example-2-create-a-page-in-a-section
 * * @see https://docs.microsoft.com/en-us/graph/api/section-post-pages?view=graph-rest-1.0#example-3-create-a-page-in-a-section-with-html-content
 * * @see https://docs.microsoft.com/en-us/graph/api/conversation-list-messages?view=graph-rest-1.0#example-2-get-messages-in-a-conversation
 *  
 * * This function is a placeholder and does not implement the actual export logic yet. 
 * You will need to implement the logic to fetch the conversation messages and create OneNote pages.
 * 
 */
export
async function exportThread() {
  const token = await getToken();
  Office.context.mailbox.item.getConversationIdAsync(async result => {
    const convId = result.value;
    const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=conversationId eq '${convId}'&$orderby=sentDateTime asc`;
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const data = await res.json();
    if (!selectedNotebookId) return alert("Select a notebook first.");

    const sectionRes = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${selectedNotebookId}/sections`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ displayName: "Exported Thread " + new Date().toISOString() })
    });
    const section = await sectionRes.json();

    for (const msg of data.value) {
      const html = `
        <html>
          <head><title>${msg.subject}</title></head>
          <body>
            <h1>${msg.subject}</h1>
            <p><b>From:</b> ${msg.from?.emailAddress?.name}</p>
            <p><b>Date:</b> ${msg.sentDateTime}</p>
            <hr />
            <div>${msg.body?.content || ""}</div>
          </body>
        </html>`;
      await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${section.id}/pages`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "text/html"
        },
        body: html
      });
    }

    document.getElementById("output").textContent = "Thread exported to OneNote.";
  });
}
