const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");

// Configure credentials
const tenantId = "85d6804d-5ae1-4a1e-9c3b-3d5355b54442";
const clientId = "a73f5240-e06c-43a3-8328-1fbd80766263";
const clientSecret = "";
const scopes = ["https://graph.microsoft.com/.default"];

// Create credential and auth provider
const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes });

// Initialize Graph client
const client = Client.initWithMiddleware({ authProvider });

// Make a Graph API call with detailed error handling
async function getUsers() {
    try {
        console.log('🔍 Testing connection and permissions...');
        console.log(`Tenant ID: ${tenantId}`);
        console.log(`Client ID: ${clientId}`);
        console.log(`Scopes: ${scopes.join(', ')}`);
        
        // First, test application info endpoint (works with client credentials)
        console.log('\n📋 Step 1: Testing /applications endpoint (client credentials compatible)...');
        try {
            const apps = await client.api("/applications").top(1).get();
            console.log('✅ Application endpoint works! Found', apps.value.length, 'applications');
        } catch (appError) {
            console.log('❌ Application endpoint failed:', appError.code, appError.message);
        }
        
        // Now try the users endpoint
        console.log('\n👥 Step 2: Testing /users endpoint...');
        const users = await client.api("/users").get();
        console.log('✅ Success! Found users:', users.value.length);
        
        // Display first few users
        users.value.slice(0, 3).forEach((user, index) => {
            console.log(`${index + 1}. ${user.displayName || user.userPrincipalName}`);
        });
        
    } catch (error) {
        console.error('\n❌ ERROR DETAILS:');
        console.error('Code:', error.code);
        console.error('Message:', error.message);
        
        if (error.code === 'Authorization_RequestDenied') {
            console.error('\n💡 SOLUTION STEPS:');
            console.error('1. Go to Azure Portal → Azure Active Directory → App registrations');
            console.error('2. Find your app:', clientId);
            console.error('3. Go to "API permissions"');
            console.error('4. Click "Add a permission" → Microsoft Graph → Application permissions');
            console.error('5. Add: User.Read.All or User.ReadBasic.All');
            console.error('6. ⚠️  CRITICAL: Click "Grant admin consent for [your organization]"');
            console.error('7. Wait a few minutes and try again');
            console.error('\nCurrent scopes:', scopes.join(', '));
        }
        
        if (error.code === 'Forbidden') {
            console.error('\n💡 PERMISSION ISSUE:');
            console.error('Your app has permissions but they may not be the right type');
            console.error('- Application permissions need admin consent');
            console.error('- Check if you have User.Read.All APPLICATION permission (not delegated)');
        }
        
        console.error('\n🔧 Debug info:');
        console.error('Full error object:', JSON.stringify(error, null, 2));
    }
}

console.log('🚀 Starting Graph API test for /users endpoint...');
getUsers();