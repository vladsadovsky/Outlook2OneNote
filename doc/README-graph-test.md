# OneNote Graph API Test Scripts

This directory contains three different Node.js scripts to test Microsoft Graph API integration for listing OneNote notebooks. Choose the one that works best for your testing scenario.

## Test Options

### 1. Simple Test (Recommended for Quick Testing)
**File**: `simple-graph-test.js`  
**Best for**: Quick validation with minimal setup

Uses a manually obtained token from Microsoft Graph Explorer:
```bash
node simple-graph-test.js
```

**Setup**:
1. Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft account
3. Run any query (like `GET /me`)
4. Copy the access token from the "Access token" tab
5. Replace `TOKEN` variable in the script

### 2. Interactive Test (Most Realistic)
**File**: `interactive-graph-test.js`  
**Best for**: Testing browser-based OAuth flow

Opens a browser for authentication (similar to your add-in):
```bash
node interactive-graph-test.js
```

**Setup**: No additional configuration needed, uses your existing client ID.

### 3. Device Code Test (Most Complete)
**File**: `test-graph-notebooks.js`  
**Best for**: Testing MSAL authentication library

Uses device code flow for authentication:
```bash
node test-graph-notebooks.js
```

**Setup**: Requires Azure AD app configuration for device code flow.

## Prerequisites

1. **Azure AD App Registration**: You need an Azure AD application registered with the following permissions:
   - `Notes.Read` (to read OneNote notebooks)  
   - `User.Read` (basic user profile)

2. **Node.js Dependencies**: The required packages are already installed:
   ```bash
   npm install @azure/msal-node @microsoft/microsoft-graph-client
   ```

## Configuration

Before running the script, update the `CLIENT_CONFIG` in `test-graph-notebooks.js`:

```javascript
const CLIENT_CONFIG = {
  auth: {
    clientId: 'YOUR_AZURE_AD_CLIENT_ID', // Replace with your actual client ID
    authority: 'https://login.microsoftonline.com/common'
  }
};
```

The script currently uses the same client ID from your manifest.xml (`a73f5240-e06c-43a3-8328-1fbd80766263`).

## How to Run

1. **Execute the script**:
   ```bash
   node test-graph-notebooks.js
   ```

2. **Follow authentication prompts**:
   - The script will display a device code and verification URL
   - Open your browser and navigate to the URL
   - Enter the device code
   - Sign in with your Microsoft account
   - Return to the terminal

3. **View results**:
   - User information will be displayed
   - All accessible OneNote notebooks will be listed
   - Details for the first notebook will be shown

## Expected Output

```
OneNote Notebooks Test - Microsoft Graph API
==================================================
🔐 Starting authentication...

📱 Please authenticate:
1. Open your browser and go to: https://microsoft.com/devicelogin
2. Enter this code: ABC123DEF
3. Sign in with your Microsoft account
4. Return to this terminal

✅ Authentication successful!
👤 Signed in as: user@example.com

📋 Getting user information...
👤 User Details:
   Name: John Doe
   Email: john.doe@example.com
   ID: 12345678-1234-1234-1234-123456789012

📚 Fetching OneNote notebooks...

📊 Found 3 notebook(s):

1. 📓 Personal Notebook
   ID: 1-abc123...
   Created: 1/15/2024
   Modified: 8/5/2025
   Is Default: true
   Notebook URL: https://onenote.com/...

2. 📓 Work Notes
   ID: 1-def456...
   Created: 3/10/2024
   Modified: 8/1/2025
   Is Default: false
   Notebook URL: https://onenote.com/...
```

## Troubleshooting

### Common Issues:

1. **Authentication fails**: 
   - Check that your Azure AD app is properly configured
   - Ensure redirect URIs include device code flow

2. **"Forbidden" error**:
   - Verify Notes.Read permission is granted
   - Check if admin consent is required
   - Ensure OneNote is enabled for your account

3. **No notebooks found**:
   - Make sure you have OneNote notebooks in your account
   - Check permissions on existing notebooks

### Permission Setup in Azure AD:

1. Go to Azure Portal → App Registrations
2. Select your app (`a73f5240-e06c-43a3-8328-1fbd80766263`)
3. Go to "API Permissions"
4. Add permissions:
   - Microsoft Graph → Delegated permissions → `Notes.Read`
   - Microsoft Graph → Delegated permissions → `User.Read`
5. Click "Grant admin consent"

## Integration with Outlook Add-in

This test script demonstrates the same Graph API calls you can use in your Outlook add-in. Key differences for add-in integration:

1. **Authentication**: Replace device code flow with SSO token from Office.js
2. **Token acquisition**: Use `Office.auth.getAccessToken()` instead of MSAL
3. **Scope**: Ensure your manifest includes the required Graph scopes

Example integration code for your add-in:
```javascript
// In your Outlook add-in
async function getOneNoteNotebooks() {
  try {
    // Get token from Office SSO
    const token = await Office.auth.getAccessToken();
    
    // Initialize Graph client
    const graphClient = Client.init({
      authProvider: {
        getAccessToken: async () => token
      }
    });
    
    // Use same API call as in test script
    const notebooks = await graphClient.api('/me/onenote/notebooks').get();
    return notebooks.value;
    
  } catch (error) {
    console.error('Failed to get notebooks:', error);
    return [];
  }
}
```

## Files

- `test-graph-notebooks.js` - Main test script
- `package.json` - Node.js dependencies (updated automatically)
- `manifest.xml` - Contains the client ID and required scopes for the add-in
