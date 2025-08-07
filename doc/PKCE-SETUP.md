# PKCE Authentication Setup Guide

This document provides complete setup instructions for OAuth 2.0 Authorization Code Flow with PKCE (Proof Key for Code Exchange) authentication in the Outlook2OneNote add-in.

## Overview

The add-in has been modernized to use secure OAuth 2.0 PKCE authentication as the primary method for accessing Microsoft Graph API, with Office.js SSO as a fallback option.

### Authentication Methods (Priority Order):
1. **PKCE OAuth 2.0** (Primary) - Secure client-side authentication
2. **Office.js SSO** (Fallback) - For Office Add-in environments
3. **Mock Data** (Development) - For testing without authentication

## Azure AD App Registration Setup

### Step 1: Create Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps)
2. Click **"New registration"**
3. Fill in the details:
   - **Name**: `Outlook2OneNote Add-in`
   - **Supported account types**: `Accounts in any organizational directory and personal Microsoft accounts`
   - **Redirect URI**: Leave blank for now

### Step 2: Configure Platform Settings

1. In your app registration, go to **Authentication**
2. Click **"Add a platform"**
3. Select **"Single-page application (SPA)"**
4. Add redirect URIs:
   ```
   http://localhost:3000/src/auth/callback.html (for development)
   https://your-domain.com/src/auth/callback.html (for production)
   ```
5. Click **"Configure"**

### Step 3: API Permissions

1. Go to **API permissions**
2. Click **"Add a permission"**
3. Select **"Microsoft Graph"**
4. Choose **"Delegated permissions"**
5. Add the following permissions:
   - `Notes.Read` - Read OneNote notebooks
   - `Notes.ReadWrite` - Create and modify OneNote pages
   - `User.Read` - Read user profile
   - `profile` - Basic profile information
   - `openid` - OpenID Connect sign-in
   - `email` - User's email address

6. Click **"Add permissions"**
7. **Important**: Click **"Grant admin consent"** (if you have admin rights)

### Step 4: Get Application (Client) ID

1. Go to **Overview** in your app registration
2. Copy the **Application (client) ID**
3. Save this ID - you'll need it in the next section

## Code Configuration

### Step 1: Update Azure AD Configuration

Edit `src/common/auth-config.js`:

```javascript
export const AZURE_AD_CONFIG = {
  // Replace with your actual client ID from Azure AD
  clientId: 'your-actual-client-id-here',
  
  // Use 'common' for multi-tenant or your tenant ID for single-tenant
  authority: 'https://login.microsoftonline.com/common',
  
  // Permissions (should match what you configured in Azure AD)
  scopes: [
    'https://graph.microsoft.com/Notes.Read',
    'https://graph.microsoft.com/Notes.ReadWrite',
    'https://graph.microsoft.com/User.Read',
    'profile',
    'openid',
    'email'
  ],
  
  // Will be auto-detected based on your hosting environment
  redirectUri: getRedirectUri(),
  postLogoutRedirectUri: getPostLogoutRedirectUri()
};
```

### Step 2: Verify Redirect URI Configuration

The redirect URI is automatically configured based on your environment:

- **Development**: `http://localhost:3000/src/auth/callback.html`
- **Production**: `https://your-domain.com/src/auth/callback.html`

Make sure this matches exactly what you configured in Azure AD.

### Step 3: Update Office Add-in Manifest (Optional)

If you want to maintain SSO fallback capability, update your `manifest.xml`:

```xml
<WebApplicationInfo>
  <Id>your-azure-ad-client-id</Id>
  <Resource>https://graph.microsoft.com</Resource>
  <Scopes>
    <Scope>https://graph.microsoft.com/Notes.Read</Scope>
    <Scope>https://graph.microsoft.com/Notes.ReadWrite</Scope>
    <Scope>https://graph.microsoft.com/User.Read</Scope>
    <Scope>profile</Scope>
    <Scope>openid</Scope>
  </Scopes>
</WebApplicationInfo>
```

## Development Setup

### Prerequisites

- Node.js 14 or higher
- Modern web browser with ES6 modules support
- Web server that serves files with proper MIME types

### Running the Application

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Start development server**:
   ```bash
   npm run dev-server
   ```

3. **Test authentication**:
   - Open the task pane
   - Click "Choose Notebook"
   - You should be redirected to Microsoft sign-in
   - After successful authentication, you'll see available notebooks

### Development vs Production

The authentication system automatically detects the environment:

- **Development**: Uses `localhost:3000` and enables additional logging
- **Production**: Uses your actual domain and minimizes logging

## Authentication Flow

### PKCE Flow (Primary Method)

1. User clicks "Choose Notebook"
2. System generates PKCE code verifier and challenge
3. User is redirected to Microsoft authorization endpoint
4. User signs in and grants permissions
5. Microsoft redirects back with authorization code
6. System exchanges code for access token using PKCE verifier
7. Access token is used to call Microsoft Graph API
8. Notebooks are retrieved and displayed

### SSO Fallback

If PKCE fails (e.g., in certain Office environments), the system automatically falls back to Office.js SSO:

1. System detects Office.js environment
2. Uses `getAccessTokenAsync` to get SSO token
3. Attempts direct Graph API call
4. Falls back to mock data if token isn't compatible

## Security Features

### PKCE Benefits
- **No client secret required** - Safe for client-side applications
- **Dynamic code challenges** - Prevents authorization code interception
- **Cross-platform compatibility** - Works in browsers, mobile, desktop

### Token Management
- **Secure storage** - Uses `sessionStorage` by default
- **Automatic refresh** - Handles token expiration transparently
- **Proper cleanup** - Clears sensitive data on logout

### Error Handling
- **Graceful degradation** - Falls back through multiple auth methods
- **User-friendly messages** - Clear error descriptions
- **Development support** - Detailed logging in dev mode

## Troubleshooting

### Common Issues

**Issue**: "CLIENT_ID must be set to your Azure AD Application ID"
- **Solution**: Update `clientId` in `auth-config.js` with your actual Azure AD application ID

**Issue**: "Redirect URI mismatch"
- **Solution**: Ensure redirect URI in Azure AD exactly matches your hosting URL

**Issue**: "Consent required"
- **Solution**: Grant admin consent in Azure AD or have users consent individually

**Issue**: "CORS errors"
- **Solution**: Ensure your web server serves the callback page with proper headers

### Debug Mode

Set `NODE_ENV=development` to enable:
- Detailed console logging
- Mock data fallback
- Additional error information

### Testing Authentication

You can test each authentication method individually:

```javascript
// Test PKCE authentication
import { pkceAuth } from './src/common/pkce-auth.js';
await pkceAuth.authenticateAndGetNotebooks();

// Test SSO fallback
import { trySSoAuthentication } from './src/common/graphapi-auth.js';
await trySSoAuthentication();

// Check authentication status
import { checkAuthenticationStatus } from './src/taskpane/onenote-service.js';
await checkAuthenticationStatus();
```

## Migration from Legacy Authentication

The system has been updated to remove legacy Exchange authentication methods while maintaining backward compatibility:

### Removed Methods
- ❌ REST API Token (`getCallbackTokenAsync` with `isRest: true`)
- ❌ Exchange Token (`getCallbackTokenAsync` with `isRest: false`)

### New Methods
- ✅ PKCE OAuth 2.0 (Primary)
- ✅ Office.js SSO (Fallback only)
- ✅ Mock Data (Development)

### Code Changes Required
- Update `auth-config.js` with your Azure AD settings
- Ensure callback page is accessible at the redirect URI
- Test authentication flow in your environment

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure AD App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Office Add-ins Authentication](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins)
- [OAuth 2.0 PKCE Specification](https://tools.ietf.org/html/rfc7636)

## Support

If you encounter issues:
1. Check the browser console for detailed error messages
2. Verify your Azure AD configuration matches the setup guide
3. Test with mock data first to isolate authentication issues
4. Enable development mode for additional debugging information
