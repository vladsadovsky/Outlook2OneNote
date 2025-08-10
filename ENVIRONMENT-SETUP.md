# Environment Setup Instructions

## Quick Start

1. **Copy the environment template:**
   ```bash
   cp .env-template .env
   ```

2. **Update the `.env` file with your actual values:**
   - Replace `your-azure-app-client-id-here` with your actual Azure AD client ID
   - Replace `your-azure-app-client-secret-here` with your actual client secret
   - Replace `your-backend-service.azurewebsites.net` with your backend service URL (if using)
   - Adjust redirect URIs if not using localhost:3000

## Required Environment Variables

   ```env
   # Azure AD Application Configuration
   CLIENT_ID=your-azure-app-client-id-here
   CLIENT_SECRET=your-azure-app-client-secret-here
   TENANT_ID=your-tenant-id-or-common
   
   # Application URLs (adjust for your environment)
   REDIRECT_URI=https://localhost:3000/auth/callback
   POST_LOGOUT_REDIRECT_URI=https://localhost:3000
   
   # Backend Service Configuration
   BACKEND_SERVICE_URL=https://your-backend-service.azurewebsites.net
   ```

## Azure AD App Registration Setup

1. **Register a new Azure AD application:**
   - Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
   - Click "New registration"
   - Name: "Outlook2OneNote Add-in"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: `https://localhost:3000/auth/callback` (Web)

2. **Configure the application:**
   - Copy the "Application (client) ID" to your `.env` file as `CLIENT_ID`
   - Go to "Certificates & secrets" > "Client secrets" > "New client secret"
   - Copy the secret value to your `.env` file as `CLIENT_SECRET`
   - Go to "API permissions" > "Add a permission" > "Microsoft Graph" > "Delegated permissions"
   - Add: `Notes.Read`, `User.Read`
   - Click "Grant admin consent"

3. **Configure redirect URIs:**
   - Add `https://localhost:3000/auth/callback`
   - Add `https://localhost:3000` (for logout)

## Backend Service Setup (Optional)

For enhanced security with client secret validation:

1. **Deploy the backend service:**
   ```bash
   # Install dependencies
   npm install express @azure/msal-node jsonwebtoken jwks-rsa dotenv
   
   # Run the backend service
   node backend-example.js
   ```

2. **Set backend environment variables:**
   - Use the same `CLIENT_ID` and `CLIENT_SECRET`
   - Set `TENANT_ID` to your Azure AD tenant ID
   - Deploy to Azure Functions, Azure App Service, or similar

3. **Update frontend configuration:**
   - Set `BACKEND_SERVICE_URL` in your `.env` file to your deployed backend URL

## Security Notes

- ⚠️ **Never commit the `.env` file to git** - it contains sensitive secrets
- The `.env` file is already added to `.gitignore`
- In production, use secure environment variable management (Azure Key Vault, etc.)
- Client secrets should only be used in backend services, never in client-side code

## Development vs Production

### Development (Client-side only):
- Uses PKCE OAuth flow without client secret
- Tokens are exchanged directly with Azure AD
- Less secure but suitable for development/testing

### Production (With backend service):
- Uses PKCE + client secret for enhanced security
- Backend service handles sensitive operations
- Tokens are validated server-side
- Recommended for production deployments

## Testing the Authentication

1. **Start the development server:**
   ```bash
   npm run dev-server
   ```

2. **Test the authentication flow:**
   - Open https://localhost:3000
   - Click "Choose Notebook" in the Office Add-in
   - You should be redirected to Azure AD login
   - After consent, you should see your real OneNote notebooks

## Troubleshooting

- **"CLIENT_SECRET is required" error:** Make sure your `.env` file contains the client secret
- **"Redirect URI mismatch" error:** Ensure the redirect URI in Azure AD matches your `.env` file
- **"Invalid client" error:** Verify your client ID is correct in both Azure AD and `.env`
- **CORS errors:** Ensure your backend service allows your frontend origin
