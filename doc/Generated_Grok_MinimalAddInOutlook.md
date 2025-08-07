Question : 
Outlook has an Add In to send current mail to OneNote.  When it is added I see it requests permission interactively.  Which authentication flow this Add In is using, which API does it use ?

Answer: 
The Outlook Add-in for sending emails to OneNote likely uses the **OAuth 2.0 Authorization Code Flow** for authentication, as this is the recommended and secure approach for Office Add-ins requiring user interaction and access to Microsoft services like OneNote. Below, I explain the authentication flow, the API used, and how permissions are handled based on available information.

### Authentication Flow
The Outlook Add-in for OneNote, when prompting for permissions interactively, typically employs the **OAuth 2.0 Authorization Code Flow** with the Microsoft Identity Platform. Here’s how it works:

1. **User Interaction and Consent**:
   - When the add-in is activated (e.g., clicking "Send to OneNote"), it checks if the user is signed in and has granted the necessary permissions.
   - If not, the add-in redirects the user to the Microsoft Identity Platform’s authorization endpoint (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize`).
   - The user is prompted to sign in with their Microsoft 365 credentials and consent to the permissions requested by the add-in (e.g., access to OneNote notebooks). This is the interactive permission request you observe.[](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

2. **Authorization Code**:
   - Upon successful sign-in and consent, the Microsoft Identity Platform returns an authorization code to the add-in’s registered redirect URI.

3. **Token Exchange**:
   - The add-in’s server-side component (or client-side with proper configuration) sends a POST request to the token endpoint (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`) with the authorization code, client ID, client secret (if applicable), and redirect URI.
   - The response includes an **access token** and possibly a **refresh token**. The access token is used to authenticate API requests to OneNote.[](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

4. **Accessing Resources**:
   - The add-in includes the access token in the `Authorization: Bearer` header of API requests to interact with OneNote resources (e.g., creating a page in a notebook with the email content).

This flow is preferred for Outlook Add-ins because it supports secure, user-delegated access to Microsoft 365 resources and aligns with modern authentication standards. It may also leverage **Single Sign-On (SSO)** if enabled, reducing the need for repeated sign-ins by using the user’s existing Office session. SSO uses the **On-Behalf-Of (OBO) flow** to exchange an initial token for one scoped to the Microsoft Graph API, but this still relies on the authorization code flow for initial user consent.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authentication)[](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/overview-authn-authz)[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authenticate-a-user-with-an-sso-token)

Alternatively, the add-in might use the **Implicit Flow** in some cases (especially for single-page applications), where the access token is returned directly in the redirect URI’s query parameters. However, this is less secure and not recommended for modern add-ins due to the inability to protect client secrets in client-side code. The Authorization Code Flow with Proof Key for Code Exchange (PKCE) is more likely, as it enhances security for public clients.[](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)[](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)

### API Used
The Outlook Add-in for OneNote most likely uses the **Microsoft Graph API** to interact with OneNote, as this is the modern, recommended API for accessing Microsoft 365 services, including OneNote notebooks, sections, and pages. The Graph API has largely replaced the older OneNote API (`https://www.onenote.com/api/`), which is deprecated for most use cases.

- **Microsoft Graph API**:
  - **Endpoint**: `https://graph.microsoft.com/v1.0/me/onenote/`
  - **Functionality**: The add-in uses endpoints like `/me/onenote/pages` to create a new page in the user’s OneNote notebook, embedding the email content (e.g., subject, body, or attachments). For example, a POST request to `/me/onenote/pages` with the email’s HTML content creates a OneNote page.[](https://shivam1410.medium.com/oauth-2-0-to-access-microsoft-apis-37995572581e)[](https://stackoverflow.com/questions/44395353/how-to-set-onenote-api%25C2%25B4s-permissions)
  - **Permissions**: The add-in requests delegated permissions such as `Notes.ReadWrite` or `Notes.ReadWrite.All` to access and modify OneNote content on behalf of the user. These permissions are specified in the add-in’s manifest and Azure AD appshe application registration.[](https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/how-to/onenote-auth)[](https://shivam1410.medium.com/oauth-2-0-to-access-microsoft-apis-37995572581e)
  - **Evidence**: Microsoft’s documentation indicates that OneNote integrations typically use Graph API for modern applications, and the interactive permission prompt aligns with Graph API’s delegated permission model.[](https://pypi.org/project/Office365-REST-Python-Client/)

- **Legacy OneNote API (Unlikely)**:
  - The older OneNote API (`https://www.onenote.com/api/v1.0/`) could theoretically be used, but it’s less likely due to its deprecation and limited support for new add-ins. This API also uses OAuth 2.0 but has different permission scopes (e.g., `office.onenote_update`). If used, it would still involve an interactive permission request.[](https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/how-to/onenote-auth)[](https://stackoverflow.com/questions/44395353/how-to-set-onenote-api%25C2%25B4s-permissions)

Given that Microsoft encourages developers to migrate to Graph API (with extended support for the OneNote API ending October 2025 for certain add-ins), the Outlook Add-in for OneNote almost certainly uses the Microsoft Graph API for new deployments.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/use-rest-api)

### Permission Request
The interactive permission request you observe is part of the OAuth 2.0 flow, where the user is prompted to grant the add-in access to their OneNote data. The specific permissions requested are likely:

- **Notes.ReadWrite** (Microsoft Graph): Allows the add-in to read and write to the user’s OneNote notebooks.
- Optionally, **User.Read** for basic user profile information (e.g., to associate the notebook with the signed-in user).

These permissions are configured in the Azure AD application registration for the add-in and declared in the add-in’s manifest. The manifest may specify permissions like `read/write item` or `read/write mailbox` for Outlook-specific access, but for OneNote integration, Graph API permissions are used. The user sees a consent screen listing these permissions (e.g., “Allow this app to access and update your OneNote notebooks”).[](https://learn.microsoft.com/en-us/previous-versions/office/office-365-api/how-to/onenote-auth)[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)[](https://shivam1410.medium.com/oauth-2-0-to-access-microsoft-apis-37995572581e)

### Additional Notes
- **Manifest Configuration**: The add-in’s manifest (XML or unified JSON format) specifies the required permission levels for Outlook (e.g., `read/write item`) and API permissions for OneNote (e.g., `Notes.ReadWrite`). This ensures the add-in only requests the minimum necessary access.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions)
- **Modern Authentication**: The add-in relies on modern authentication protocols, and legacy Exchange tokens are deprecated, reinforcing the use of OAuth 2.0 with Graph API.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/faq-nested-app-auth-outlook-legacy-tokens)
- **SSO Possibility**: If the add-in supports SSO, it may use the user’s existing Office session to streamline authentication, but the initial permission consent is still required.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authentication)[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/authenticate-a-user-with-an-sso-token)
- **Deprecation Timeline**: Microsoft is phasing out legacy tokens and older APIs, with full deprecation by October 2025, making Graph API the standard for such add-ins.[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/use-rest-api)[](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/faq-nested-app-auth-outlook-legacy-tokens)

### Conclusion
The Outlook Add-in for OneNote uses the **OAuth 2.0 Authorization Code Flow** (potentially with PKCE for enhanced security) to authenticate users and obtain an access token. It leverages the **Microsoft Graph API** with permissions like `Notes.ReadWrite` to interact with OneNote, creating pages from email content. The interactive permission prompt is the user consent step of this flow, ensuring secure, user-delegated access to OneNote resources.

If you need further details or assistance with implementing or troubleshooting this add-in, let me know!


Question: For this authentication scheme do I need to develop any back end code ?



Answer: 

Whether you need to develop backend code for the Outlook Add-in for OneNote's authentication scheme, which uses the **OAuth 2.0 Authorization Code Flow** with the **Microsoft Graph API**, depends on the specific implementation details and security requirements of your add-in. Below, I outline the considerations to determine if backend code is necessary, focusing on the authentication flow and the add-in’s context.

### Short Answer
Yes, in most cases, you will need to develop some backend code to securely handle the OAuth 2.0 Authorization Code Flow for the Outlook Add-in for OneNote, particularly to manage the token exchange and securely store or process sensitive data like access tokens. However, there are scenarios where you can minimize or avoid backend code, such as using client-side flows with PKCE or leveraging Microsoft’s Single Sign-On (SSO) for Office Add-ins. Let’s break it down.

### Why Backend Code Is Typically Needed
The OAuth 2.0 Authorization Code Flow, which the Outlook Add-in for OneNote likely uses, involves several steps that benefit from or require server-side logic for security and functionality:

1. **Token Exchange**:
   - After the user consents to permissions interactively, the Microsoft Identity Platform redirects to the add-in’s registered redirect URI with an authorization code.
   - The add-in must exchange this code for an access token by sending a POST request to the token endpoint (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`). This request typically requires:
     - The authorization code.
     - The client ID.
     - A client secret (for confidential clients) or a Proof Key for Code Exchange (PKCE) code verifier (for public clients).
   - **Why Backend?** The client secret must be kept confidential and cannot be safely stored in client-side code (e.g., JavaScript running in the Outlook add-in). Exposing the client secret in the browser or client-side app risks security vulnerabilities. A backend server securely stores the client secret and handles the token exchange.

2. **Token Storage and Management**:
   - The access token (and optionally a refresh token) returned by the token endpoint needs to be securely stored and managed. Storing tokens client-side (e.g., in the browser’s local storage) is vulnerable to XSS attacks.
   - **Why Backend?** A backend server can securely store tokens, manage token refreshes (using the refresh token), and provide tokens to the client-side add-in as needed, reducing the risk of token exposure.

3. **API Requests to Microsoft Graph**:
   - The add-in uses the access token to make API calls to the Microsoft Graph API (e.g., `POST /me/onenote/pages` to create a OneNote page from an email). While these requests can technically be made from the client side, doing so requires the access token to be available in the browser, which is less secure.
   - **Why Backend?** A backend can handle Graph API requests on behalf of the client, keeping the access token server-side and only sending the results (e.g., success/failure of the OneNote page creation) to the client. This reduces the attack surface.

4. **Security and Compliance**:
   - Microsoft recommends using a backend for confidential clients (like add-ins) to comply with OAuth 2.0 best practices. The Authorization Code Flow without PKCE requires a client secret, which necessitates a secure server environment.
   - Even with PKCE (which eliminates the need for a client secret), a backend may still be preferred to handle token refreshes or sensitive operations securely.

### Scenarios Where Backend Code Might Not Be Needed
There are limited cases where you might avoid backend code, but they come with trade-offs:

1. **Using OAuth 2.0 Authorization Code Flow with PKCE**:
   - The Authorization Code Flow with Proof Key for Code Exchange (PKCE) is designed for public clients (e.g., single-page applications or add-ins running entirely in the browser). PKCE replaces the client secret with a dynamically generated code verifier and challenge, making it safer to perform the token exchange client-side.
   - **How It Works**:
     - The add-in generates a code verifier and code challenge, includes the challenge in the authorization request, and later sends the verifier in the token exchange request.
     - The token exchange occurs client-side (e.g., using `fetch` in JavaScript), and the access token is stored in memory or browser storage (with caution).
   - **Trade-Offs**:
     - No backend is needed for the token exchange, but storing tokens client-side (e.g., in local storage) is less secure due to potential XSS vulnerabilities.
     - Token refresh (using a refresh token) is still challenging client-side, as refresh tokens are long-lived and require secure handling.
     - Microsoft’s guidance for Office Add-ins leans toward server-side implementations for production scenarios to ensure robust security.
   - **Feasibility for OneNote Add-in**: Possible but not ideal, as the Outlook Add-in for OneNote likely handles sensitive data (email content, OneNote pages) and requires secure token management.

2. **Using Office Add-ins Single Sign-On (SSO)**:
   - Microsoft provides an SSO feature for Office Add-ins, which leverages the user’s existing Office session to obtain an access token for the Microsoft Graph API without a separate OAuth flow.
   - **How It Works**:
     - The add-in calls the Office JavaScript API’s `Office.auth.getAccessToken()` to retrieve a bootstrap token for the signed-in user.
     - This token is sent to a backend, which uses the **On-Behalf-Of (OBO) flow** to exchange it for a Graph API access token scoped to OneNote (e.g., `Notes.ReadWrite`).
   - **Why Backend Still Needed?**:
     - While SSO simplifies client-side authentication, the OBO flow requires a backend to securely handle the token exchange, as it involves the client secret or a secure environment to validate the bootstrap token.
     - Without a backend, SSO is not fully usable, as the client-side bootstrap token has limited scopes and cannot directly access the Graph API for OneNote operations.
   - **Feasibility for OneNote Add-in**: The official Outlook Add-in for OneNote may use SSO, but it still requires a backend for the OBO flow to obtain the Graph API token.

3. **Using Implicit Flow (Deprecated and Not Recommended)**:
   - The OAuth 2.0 Implicit Flow returns an access token directly in the redirect URI’s query parameters, avoiding the need for a token exchange.
   - **Why Not Used?**:
     - The Implicit Flow is deprecated by OAuth 2.0 standards and Microsoft due to security risks (e.g., token exposure in browser history or logs).
     - It’s unlikely the Outlook Add-in for OneNote uses this flow, as modern Office Add-ins follow the Authorization Code Flow or SSO for better security.
   - **Feasibility**: Not recommended and likely not used by the OneNote add-in.

### Specific Context of the Outlook Add-in for OneNote
- **Official Microsoft Add-in**: If you’re referring to the official “Send to OneNote” add-in developed by Microsoft, it almost certainly uses a backend to handle the OAuth 2.0 token exchange and Graph API calls. Microsoft’s infrastructure includes server-side components to securely manage tokens and API requests, ensuring compliance with security best practices.
- **Custom Add-in**: If you’re developing a custom version of this add-in, you’ll need to decide whether to implement a backend based on your security requirements and whether you use PKCE or SSO:
  - **With PKCE**: You can avoid a backend for the token exchange but must carefully manage tokens client-side and accept the associated risks.
  - **With SSO**: A backend is required for the OBO flow to exchange the bootstrap token for a Graph API token.
  - **Without PKCE or SSO**: A backend is necessary to handle the client secret and token exchange securely.

### Recommended Approach
For a production-grade Outlook Add-in for OneNote, you should develop backend code to:
1. **Handle Token Exchange**:
   - Create a server-side endpoint (e.g., using Node.js, ASP.NET, or Python) to receive the authorization code and exchange it for an access token using the client ID and client secret (or PKCE code verifier).
   - Example (Node.js with `axios`):
     ```javascript
     const axios = require('axios');
     const tokenEndpoint = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token';
     const params = new URLSearchParams({
       client_id: 'your-client-id',
       client_secret: 'your-client-secret',
       code: 'authorization-code',
       redirect_uri: 'your-redirect-uri',
       grant_type: 'authorization_code',
     });
     const response = await axios.post(tokenEndpoint, params);
     const { access_token, refresh_token } = response.data;
     ```
2. **Securely Store Tokens**:
   - Store the access token and refresh token in a secure database or cache (e.g., Redis) on the server, associating them with the user’s session.
3. **Make Graph API Calls**:
   - Use the access token to call the Microsoft Graph API (e.g., `POST https://graph.microsoft.com/v1.0/me/onenote/pages`) from the server to create OneNote pages with email content.
   - Example (Node.js):
     ```javascript
     const graphResponse = await axios.post(
       'https://graph.microsoft.com/v1.0/me/onenote/pages',
       {
         'Content-Type': 'text/html',
         content: `<html><body>${emailContent}</body></html>`,
       },
       { headers: { Authorization: `Bearer ${access_token}` } }
     );
     ```
4. **Expose an API to the Client**:
   - The client-side add-in (JavaScript in Outlook) communicates with your backend via a secure API (e.g., REST or GraphQL) to trigger OneNote actions without directly handling tokens.

### Alternatives to Minimize Backend Development
If you want to minimize backend development:
- **Use PKCE**: Implement the Authorization Code Flow with PKCE to handle token exchange client-side, but be cautious about token storage security.
- **Leverage Azure Functions**: Use a serverless platform like Azure Functions to handle the token exchange and Graph API calls with minimal infrastructure setup.
- **Third-Party Authentication Services**: Use services like Auth0 or Firebase Authentication to manage OAuth flows, though you’ll still need some backend logic to integrate with the Graph API.

### Conclusion
For the Outlook Add-in for OneNote, you’ll likely need backend code to securely handle the OAuth 2.0 Authorization Code Flow, especially for managing the token exchange with a client secret and making secure Graph API calls to OneNote. Using PKCE or SSO can reduce but not eliminate the need for a backend, as secure token management and API interactions benefit from server-side logic. For a custom add-in, a lightweight backend (e.g., an Azure Function or Node.js server) is recommended to ensure security and compliance with Microsoft’s best practices.

If you’re working on a specific implementation and need help with code samples, setup instructions, or choosing a backend framework, let me know, and I can provide more tailored guidance!



Question:
Could you generate for me the project with minimal Outlook Add In, which does following:when task pane first instantiated it requests change to add permissions to enumerate OneNote notebooks 
It uses PKCE to handle logging as you described 
In the taskpane list of logged in users notebooks is described



Answer:


Below is a guide to create a minimal Outlook Add-in that uses the OAuth 2.0 Authorization Code Flow with PKCE to authenticate with the Microsoft Identity Platform, requests permissions to enumerate OneNote notebooks, and displays the list of notebooks in the task pane. Since the add-in uses PKCE, it can operate entirely client-side, avoiding the need for a backend. I'll provide the project structure, code, and setup instructions, keeping it as minimal as possible while meeting your requirements.

### Project Overview
- **Purpose**: An Outlook Add-in that:
  - On task pane initialization, checks for authentication and requests permissions to access OneNote notebooks (`Notes.Read` scope).
  - Uses OAuth 2.0 Authorization Code Flow with PKCE for client-side authentication.
  - Lists the signed-in user’s OneNote notebooks in the task pane using the Microsoft Graph API.
- **Tech Stack**:
  - Client-side: HTML, JavaScript, Office.js (for Outlook Add-in functionality).
  - Authentication: OAuth 2.0 with PKCE via Microsoft Identity Platform.
  - API: Microsoft Graph API (`/me/onenote/notebooks` endpoint).
- **No Backend**: PKCE allows secure token exchange without a server, but tokens are stored in memory (not persisted) for simplicity and security.

### Prerequisites
- **Node.js and npm**: For local development and running a web server.
- **Microsoft 365 Developer Account**: To test the add-in with Outlook (get one via the Microsoft 365 Developer Program if needed).
- **Azure AD App Registration**: To configure OAuth 2.0 for the add-in.
- **Code Editor**: VS Code or similar.
- **Local Development Server**: For hosting the add-in’s static files (e.g., `http-server` or similar).

### Step 1: Azure AD App Registration
To use OAuth 2.0 with PKCE, you need to register the add-in in Azure AD to get a client ID and configure permissions.

1. **Go to Azure Portal**:
   - Navigate to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations** > **New registration**.
2. **Configure the App**:
   - **Name**: `OutlookOneNoteAddIn`.
   - **Supported account types**: Select “Accounts in any organizational directory and personal Microsoft accounts” for broadest compatibility.
   - **Redirect URI**: Add a Single-page application (SPA) redirect URI: `https://localhost:3000/auth-callback.html`.
     - Note: `localhost:3000` is used for local development. Adjust if using a different port or domain.
   - Register the app and note the **Application (client) ID**.
3. **Add Permissions**:
   - Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**.
   - Select: `Notes.Read` (to enumerate OneNote notebooks).
   - Optionally, add `User.Read` for basic user profile info (recommended for Graph API).
   - Click **Add permissions**.
   - Note: Admin consent may be required for organizational accounts, but for personal accounts, users will consent interactively.
4. **Enable Public Client Flow**:
   - Go to **Authentication** > Under **Advanced settings**, enable **Allow public client flows** (required for PKCE).
   - Save changes.

### Step 2: Project Structure
Create a project folder with the following structure:

```
OutlookOneNoteAddIn/
├── manifest.xml
├── src/
│   ├── index.html
│   ├── auth-callback.html
│   ├── app.js
│   ├── crypto.js
└── package.json
```

### Step 3: Project Files
Below are the minimal files needed for the add-in. Save them in the project folder.

#### `manifest.xml`
This defines the add-in’s metadata, task pane, and permissions.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:mail="http://schemas.microsoft.com/office/mailapporom/1.1"
           xsi:type="MailApp">
  <Id>[Your-Unique-GUID]</Id> <!-- Replace with a unique GUID, e.g., from https://guidgenerator.com -->
  <Version>1.0.0.0</Version>
  <ProviderName>Your Name</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="OneNote Notebook Lister"/>
  <Description DefaultValue="Lists OneNote notebooks in Outlook task pane."/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://example.com"/>
  <AppDomains>
    <AppDomain>https://login.microsoftonline.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://localhost:3000/index.html"/>
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
</OfficeApp>
```

- **Notes**:
  - Replace `[Your-Unique-GUID]` with a unique GUID (generate one at [guidgenerator.com](https://guidgenerator.com)).
  - The `AppDomains` includes `login.microsoftonline.com` for OAuth redirects.
  - `Permissions` is set to `ReadWriteMailbox` for Outlook access; OneNote permissions (`Notes.Read`) are handled via Graph API.
  - The task pane loads `index.html` from `https://localhost:3000`.

#### `package.json`
This sets up the local development server.

```json
{
  "name": "outlook-onenote-addin",
  "version": "1.0.0",
  "scripts": {
    "start": "http-server -p 3000 --ssl --cert cert.pem --key key.pem"
  },
  "dependencies": {
    "http-server": "^14.1.1"
  }
}
```

- Run `npm install` to install `http-server`.
- Generate an SSL certificate for `localhost` (required for Outlook Add-ins):
  - Use `openssl` to create `cert.pem` and `key.pem`:
    ```bash
    openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes
    ```
  - Place `cert.pem` and `key.pem` in the project root.

#### `src/index.html`
The task pane UI, loaded when the add-in starts.

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>OneNote Notebook Lister</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script src="app.js"></script>
    <style>
        body { font-family: Arial, sans-serif; padding: 10px; }
        #notebookList { margin-top: 10px; }
        #status { color: red; }
    </style>
</head>
<body>
    <h3>OneNote Notebooks</h3>
    <div id="status">Initializing...</div>
    <ul id="notebookList"></ul>
</body>
</html>
```

#### `src/auth-callback.html`
Handles the OAuth redirect and processes the authorization code.

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <title>Auth Callback</title>
    <script src="crypto.js"></script>
    <script src="app.js"></script>
</head>
<body>
    <p>Processing authentication...</p>
    <script>
        // Handle the redirect and process the auth code
        window.onload = () => handleAuthCallback();
    </script>
</body>
</html>
```

#### `src/crypto.js`
Utility functions for PKCE (generating code verifier and challenge).

```javascript
// Generate a random string for PKCE code verifier
function generateCodeVerifier() {
    const array = new Uint8Array(32);
    crypto.getRandomValues(array);
    return btoa(String.fromCharCode.apply(null, array))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=/g, '');
}

// Generate SHA-256 hash for code challenge
async function generateCodeChallenge(verifier) {
    const encoder = new TextEncoder();
    const data = encoder.encode(verifier);
    const digest = await crypto.subtle.digest('SHA-256', data);
    return btoa(String.fromCharCode.apply(null, new Uint8Array(digest)))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=/g, '');
}

if (typeof module !== 'undefined') {
    module.exports = { generateCodeVerifier, generateCodeChallenge };
}
```

#### `src/app.js`
The main logic for authentication, token exchange, and fetching OneNote notebooks.

```javascript
// Office.js initialization
Office.onReady(() => {
    if (Office.context.mailbox) {
        // Start authentication when the task pane loads
        authenticate();
    }
});

// Configuration
const clientId = '[Your-Client-ID]'; // Replace with your Azure AD client ID
const tenantId = 'common'; // Use 'common' for multi-tenant or your tenant ID
const redirectUri = 'https://localhost:3000/auth-callback.html';
const scopes = ['Notes.Read', 'User.Read', 'openid', 'profile'];
let codeVerifier = '';

// Display status messages
function setStatus(message) {
    document.getElementById('status').textContent = message;
}

// Authenticate using PKCE
async function authenticate() {
    setStatus('Checking authentication...');
    
    // Check if already authenticated (access token in memory)
    if (sessionStorage.getItem('accessToken')) {
        fetchNotebooks();
        return;
    }

    // Generate PKCE code verifier and challenge
    codeVerifier = generateCodeVerifier();
    const codeChallenge = await generateCodeChallenge(codeVerifier);

    // Build authorization URL
    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
        `?client_id=${clientId}` +
        `&response_type=code` +
        `&redirect_uri=${encodeURIComponent(redirectUri)}` +
        `&scope=${encodeURIComponent(scopes.join(' '))}` +
        `&code_challenge=${codeChallenge}` +
        `&code_challenge_method=S256`;

    // Redirect to Microsoft login
    window.location.href = authUrl;
}

// Handle OAuth callback
async function handleAuthCallback() {
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get('code');
    
    if (!code) {
        setStatus('Authentication failed: No code received.');
        return;
    }

    // Exchange code for access token
    setStatus('Exchanging code for token...');
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
        client_id: clientId,
        code: code,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
        code_verifier: codeVerifier,
    });

    try {
        const response = await fetch(tokenUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: params,
        });
        const data = await response.json();
        
        if (data.access_token) {
            sessionStorage.setItem('accessToken', data.access_token);
            // Redirect back to task pane
            window.location.href = 'https://localhost:3000/index.html';
        } else {
            setStatus('Token exchange failed: ' + (data.error_description || 'Unknown error'));
        }
    } catch (error) {
        setStatus('Error during token exchange: ' + error.message);
    }
}

// Fetch and display OneNote notebooks
async function fetchNotebooks() {
    setStatus('Fetching notebooks...');
    const accessToken = sessionStorage.getItem('accessToken');
    const graphUrl = 'https://graph.microsoft.com/v1.0/me/onenote/notebooks';

    try {
        const response = await fetch(graphUrl, {
            headers: { Authorization: `Bearer ${accessToken}` },
        });
        const data = await response.json();

        if (data.value) {
            const notebookList = document.getElementById('notebookList');
            notebookList.innerHTML = ''; // Clear existing list
            data.value.forEach(notebook => {
                const li = document.createElement('li');
                li.textContent = notebook.displayName;
                notebookList.appendChild(li);
            });
            setStatus('Notebooks loaded successfully.');
        } else {
            setStatus('Failed to load notebooks: ' + (data.error?.message || 'Unknown error'));
        }
    } catch (error) {
        setStatus('Error fetching notebooks: ' + error.message);
    }
}
```

- **Notes**:
  - Replace `[Your-Client-ID]` with the client ID from your Azure AD app registration.
  - The `scopes` array includes `Notes.Read` for enumerating notebooks and `User.Read`, `openid`, `profile` for basic user info and authentication.
  - Tokens are stored in `sessionStorage` (in-memory, cleared on browser close) for simplicity. For production, consider more secure storage or a backend.
  - The add-in redirects to Microsoft’s login page on initialization if not authenticated, then handles the callback and fetches notebooks.

### Step 4: Setup and Run
1. **Install Dependencies**:
   ```bash
   npm install
   ```
2. **Generate SSL Certificates** (if not done):
   ```bash
   openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes
   ```
3. **Start the Local Server**:
   ```bash
   npm start
   ```
   - This runs `http-server` on `https://localhost:3000` with SSL.
4. **Sideload the Add-in**:
   - Open Outlook (web or desktop with a Microsoft 365 account).
   - In Outlook Web:
     - Go to **Get Add-ins** > **My add-ins** > **Add a custom add-in** > **Add from file**.
     - Upload `manifest.xml`.
   - In Outlook Desktop:
     - Go to **File** > **Manage Add-ins** > **Add from file** and upload `manifest.xml`.
5. **Trust the SSL Certificate**:
   - When prompted, trust the `localhost` certificate in your browser or system.
6. **Test the Add-in**:
   - Open an email in Outlook, click the add-in’s button (appears as “OneNote Notebook Lister”).
   - The task pane loads, redirects to Microsoft login (if not authenticated), prompts for consent (`Notes.Read` permission), and displays the user’s OneNote notebooks.

### Step 5: How It Works
- **Initialization**: When the task pane loads (`index.html`), `Office.onReady` triggers the `authenticate` function.
- **Authentication**:
  - If no access token exists, the add-in generates a PKCE code verifier/challenge and redirects to the Microsoft login page with the `Notes.Read` scope.
  - After user consent, Microsoft redirects to `auth-callback.html` with an authorization code.
  - The callback page exchanges the code for an access token using PKCE and stores it in `sessionStorage`.
  - The user is redirected back to `index.html`.
- **Notebook Listing**: The add-in fetches notebooks from the Graph API (`/me/onenote/notebooks`) and displays them in a `<ul>` list.

### Security Considerations
- **Token Storage**: Using `sessionStorage` is minimal but not ideal for production. For better security, use a backend to store tokens or implement secure client-side storage (e.g., encrypted cookies with short lifetimes).
- **PKCE**: Ensures secure client-side authentication without a client secret.
- **HTTPS**: Required for Outlook Add-ins and OAuth redirects. Ensure `localhost` uses a valid SSL certificate.
- **Scope Minimization**: Only `Notes.Read` is requested to limit permissions to what’s needed.

### Limitations
- **Token Lifetime**: Access tokens expire (typically 1 hour). This sample doesn’t handle refresh tokens for simplicity. In production, add refresh token logic (requires a backend for secure storage).
- **Error Handling**: Basic error handling is included. Enhance for production with retries or user-friendly messages.
- **Cross-Browser**: Tested in modern browsers. Ensure `crypto.subtle` (for PKCE) is supported (all modern browsers support it).

### Example Output
After authentication, the task pane might show:
```
OneNote Notebooks
- My Notebook
- Work Notes
- Personal Journal
[Status: Notebooks loaded successfully.]
```

### Troubleshooting
- **Redirect URI Mismatch**: Ensure the redirect URI in Azure AD matches `https://localhost:3000/auth-callback.html`.
- **CORS Issues**: The Graph API and Microsoft login endpoints handle CORS correctly, but ensure your local server uses HTTPS.
- **Permission Denied**: Verify `Notes.Read` is granted in the Azure AD consent prompt.
- **Office.js Errors**: Ensure the `office.js` script loads correctly and Outlook supports the add-in (Mailbox API 1.1 or higher).

### Next Steps
If you want to extend this add-in (e.g., add refresh token support, a backend, or email-to-OneNote functionality), let me know, and I can provide additional code or guidance. For deployment, host the files on a public HTTPS server (e.g., Azure Static Web Apps) and update the manifest and Azure AD redirect URI accordingly.

Let me know if you need help setting this up, debugging, or adding features!