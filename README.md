
# Outlook to OneNote Web Add-in

This project is an Office 365 Outlook Web Add-in that allows users to export an email thread to a OneNote notebook using the Microsoft Graph API.

---

## 🚀 Features

- Adds a ribbon button to Outlook Web/Desktop
- Lets users choose a OneNote notebook via Microsoft Graph
- Creates a section and adds one page per email in the selected thread

---

## ⚙️ Prerequisites

1. **Node.js** installed (https://nodejs.org/)
2. **Office 365 Outlook** (Web or Desktop version)
3. **Azure AD App Registration** with the following API permissions:
   - `User.Read`
   - `Mail.Read`
   - `Notes.ReadWrite`
4. **Trusted local HTTPS certificate**

---

## 🛠 Setup Instructions

### 1. Register an Azure AD App

Go to [https://portal.azure.com](https://portal.azure.com):
- Register a new app (e.g., "OutlookOneNoteAddin")
- Add platform: `Single-page application` → `https://localhost:3000/taskpane.html`
- Add API permissions:
  - Microsoft Graph → Delegated → `User.Read`, `Mail.Read`, `Notes.ReadWrite`
- Copy the `Application (client) ID` and replace `YOUR_CLIENT_ID_HERE` in `taskpane.js`

---

### 2. Generate HTTPS Certificate

```bash
mkdir certs
openssl req -x509 -newkey rsa:2048 -nodes -keyout certs/localhost.key -out certs/localhost.crt -days 365 -subj "/CN=localhost"
```

Install `localhost.crt` in your OS trust store (Trusted Root CA).

---

### 3. Start a Local HTTPS Server

From the `server` directory:
```bash
npm init -y
npm install express
```

Create `server.js`:
```js
const fs = require('fs');
const https = require('https');
const express = require('express');
const app = express();

app.use(express.static('public'));

https.createServer({
  key: fs.readFileSync('certs/localhost.key'),
  cert: fs.readFileSync('certs/localhost.crt')
}, app).listen(3000, () => {
  console.log('Server running at https://localhost:3000');
});
```

Run:
```bash
node server.js
```

---

### 4. Sideload the Add-in in Outlook

1. Open Outlook Web
2. Go to ⚙️ Settings > View all Outlook settings > Mail > Customize actions > Add-ins
3. Select “Upload custom add-in” → from file
4. Choose your `manifest.xml` (must match local URLs)

---

### 5. Test the Plugin

- Open an email thread
- Click “Export to OneNote” from the ribbon
- Authenticate with Microsoft account
- Select notebook → A new section is created with one page per message

---

## 📁 Project Structure

```
server/
  public/
    taskpane.html
    taskpane.js
    function.html
    assets/
      icon-16.png
      icon-32.png
      icon-80.png
manifest.xml
README.md
```

---

## 🧪 Notes

- Make sure Outlook trusts the self-signed cert
- If using Edge/Chrome, allow loading from `localhost` HTTPS
- For deployment, you must host over a public HTTPS domain with valid cert


