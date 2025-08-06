# Authentication Setup for OneNote Integration

## Overview
The Outlook2OneNote add-in has been configured to use Microsoft Graph API for OneNote access. This requires proper Azure AD app registration and permissions.

## Manifest Changes Made

### 1. WebApplicationInfo Section Added
```xml
<WebApplicationInfo>
  <Id>a73f5240-e06c-43a3-8328-1fbd80766263</Id>
  <Resource>https://graph.microsoft.com</Resource>
  <Scopes>
    <Scope>Notes.ReadWrite</Scope>
    <Scope>Notes.Read</Scope>
    <Scope>User.Read</Scope>
  </Scopes>
</WebApplicationInfo>
```

### 2. IdentityAPI Requirement Added
- Added `<Set Name="IdentityAPI" MinVersion="1.3"/>` to both Requirements sections
- This enables modern authentication capabilities

## Azure AD App Registration Requirements

### Step 1: Create/Update Azure AD App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory > App registrations**
3. Find your app registration with ID: `a73f5240-e06c-43a3-8328-1fbd80766263`
4. If it doesn't exist, create a new one with this specific ID

### Step 2: Configure API Permissions
Add the following Microsoft Graph permissions:
- `Notes.Read` - Read user OneNote notebooks
- `Notes.ReadWrite` - Read and write user OneNote notebooks  
- `User.Read` - Read user profile (basic permission)

### Step 3: Authentication Configuration
1. **Redirect URIs**: Add the following redirect URIs:
   - `https://localhost:3000` (for development)
   - Your production domain when deployed
   
2. **Implicit Grant**: Enable the following tokens:
   - Access tokens
   - ID tokens

### Step 4: API Permissions Consent
1. Grant admin consent for the organization (if required)
2. Ensure permissions are granted and not pending

## Code Implementation Status

### ✅ Completed Features
- `getOneNoteNotebooks()` - Retrieves user's OneNote notebooks
- `showNotebookPopup()` - Interactive notebook selection UI
- Global notebook storage for export functionality
- Multiple authentication fallback methods
- Mock data fallback for testing

### 🔄 Authentication Flow
1. **Primary**: Uses `Office.context.auth.getAccessTokenAsync()` with Graph API scopes
2. **Fallback**: Uses `Office.context.mailbox.getCallbackTokenAsync()` 
3. **Testing**: Returns mock notebook data if authentication fails

### 🚀 Next Steps for Production
1. Complete Azure AD app registration with proper permissions
2. Update the App ID in manifest.xml if using a different registration
3. Test authentication flow with real OneNote data
4. Implement actual OneNote page creation API calls

## Testing the Current Implementation

### Development Testing
The current implementation includes robust fallback mechanisms:

1. **Choose Notebook** button will work immediately with mock data
2. **Export Thread** functionality is fully prepared
3. Console logging provides detailed debugging information

### Production Testing Checklist
- [ ] Azure AD app registered with correct ID
- [ ] Microsoft Graph API permissions granted
- [ ] Redirect URIs configured
- [ ] Admin consent provided (if required)
- [ ] Test with real OneNote notebooks
- [ ] Test export functionality with actual API calls

## Troubleshooting Common Issues

### Authentication Errors
- Verify Azure AD app registration ID matches manifest
- Check that API permissions are granted, not just requested
- Ensure redirect URIs include your development/production URLs

### CORS Issues
- Microsoft Graph API calls should work from Office Add-ins without CORS issues
- If problems persist, verify the app registration configuration

### Permission Errors
- Ensure `Notes.ReadWrite` permission is granted for OneNote access
- Check that the user has OneNote enabled in their Microsoft 365 subscription

## Security Considerations
- Never hardcode access tokens or secrets in client-side code
- The current implementation uses Office.js secure token methods
- Tokens are automatically managed by the Office platform
- All authentication is handled through Microsoft's secure channels

## Contact Information
For issues related to Azure AD configuration, consult your Microsoft 365 administrator or refer to the [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/).
