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
