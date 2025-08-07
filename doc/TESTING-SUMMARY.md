# 🚀 OneNote Graph API Test Suite

## Overview
I've created **three different test scripts** to help you validate Microsoft Graph API integration for OneNote notebooks before integrating into your Outlook add-in.

## 📁 Files Created

### Test Scripts
1. **`simple-graph-test.js`** - Token-based testing (easiest)
2. **`interactive-graph-test.js`** - Browser-based auth (most realistic) 
3. **`test-graph-notebooks.js`** - Device code flow (most complete)

### Documentation
- **`README-graph-test.md`** - Detailed setup and usage instructions
- **`TESTING-SUMMARY.md`** - This file

## 🎯 Recommended Testing Approach

### Start Here: Simple Test
```bash
node simple-graph-test.js
```

**Why first?**: Fastest way to verify your Graph API permissions work.

**Steps**:
1. Visit [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in and run `GET /me`
3. Copy access token from "Access token" tab
4. Replace `TOKEN` variable in `simple-graph-test.js`
5. Run the script

### Then Try: Interactive Test
```bash
node interactive-graph-test.js
```

**Why next?**: Tests the OAuth flow your add-in will actually use.

**Features**:
- Opens browser automatically
- Uses same client ID as your add-in
- Shows realistic authentication flow
- No manual token copying needed

## 🔧 What These Scripts Test

### ✅ Graph API Capabilities
- User authentication
- OneNote notebooks listing
- Notebook sections retrieval
- Error handling
- Permission validation

### 📊 Expected Output
```
📚 Fetching OneNote notebooks...

📊 Found 3 notebook(s):

1. 📓 "Personal Notebook"
   🆔 ID: 1-abc123...
   📅 Created: 1/15/2024
   📝 Modified: 8/5/2025
   ⭐ Default: Yes

2. 📓 "Work Notes"
   🆔 ID: 1-def456...
   📅 Created: 3/10/2024
   📝 Modified: 8/1/2025
   ⭐ Default: No
```

## 🔗 Integration with Your Outlook Add-in

### Key Learnings from Tests
1. **Same API endpoints** work in your add-in
2. **Authentication differs**: Use `Office.auth.getAccessToken()` instead of manual auth
3. **Same permissions needed**: `Notes.Read` and `User.Read`
4. **Error handling**: Important for production use

### Code Translation
**Test script approach:**
```javascript
// In test script
const graphClient = Client.init({
  authProvider: {
    getAccessToken: async () => TOKEN
  }
});
const notebooks = await graphClient.api('/me/onenote/notebooks').get();
```

**Your add-in approach:**
```javascript
// In your Outlook add-in
async function getOneNoteNotebooks() {
  try {
    const token = await Office.auth.getAccessToken();
    const graphClient = Client.init({
      authProvider: {
        getAccessToken: async () => token
      }
    });
    const notebooks = await graphClient.api('/me/onenote/notebooks').get();
    return notebooks.value;
  } catch (error) {
    console.error('Failed to get notebooks:', error);
    return [];
  }
}
```

## 🛠 Dependencies Installed
The following packages are now available in your project:
- `@microsoft/microsoft-graph-client` - Microsoft Graph SDK
- `@azure/msal-node` - Authentication library

## 🐛 Common Issues & Solutions

### "Forbidden" Error
- **Cause**: Missing permissions
- **Fix**: Ensure your Azure AD app has `Notes.Read` permission granted

### "No notebooks found"
- **Cause**: User has no OneNote notebooks
- **Fix**: Create a test notebook in OneNote online

### Authentication fails
- **Cause**: Client ID or permissions issues  
- **Fix**: Verify your Azure AD app configuration

## 🎉 Success Criteria
✅ Script runs without errors  
✅ Successfully authenticates  
✅ Lists your OneNote notebooks  
✅ Shows notebook details  
✅ Handles errors gracefully  

## 🔜 Next Steps
1. **Test the scripts** to verify Graph API access
2. **Integrate learnings** into your Outlook add-in
3. **Replace authentication** with Office SSO
4. **Add error handling** for production use
5. **Consider caching** notebooks for better UX

---
*These test scripts validate the same Microsoft Graph API calls your Outlook add-in will use, ensuring your OneNote integration works before full implementation.*
