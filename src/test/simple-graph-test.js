#!/usr/bin/env node

/**
 * Simple OneNote Notebooks Test - Token Version
 * 
 * This simplified version uses a manually obtained access token to test
 * Microsoft Graph API calls for OneNote notebooks.
 * 
 * To get an access token:
 * 1. Go to: https://developer.microsoft.com/en-us/graph/graph-explorer
 * 2. Sign in with your Microsoft account
 * 3. Try any query (like GET /me)
 * 4. Click on "Access token" tab to copy your token
 * 5. Paste it in the TOKEN variable below
 * 
 * Note: This is for testing only. In production, use proper OAuth flow.
 */

const { Client } = require('@microsoft/microsoft-graph-client');

// REPLACE THIS WITH YOUR ACTUAL TOKEN FROM GRAPH EXPLORER
const TOKEN = 'YOUR_ACCESS_TOKEN_HERE';

class SimpleOneNoteTest {
  constructor(accessToken) {
    if (!accessToken || accessToken === 'YOUR_ACCESS_TOKEN_HERE') {
      throw new Error('Please provide a valid access token. See instructions in the file.');
    }
    
    this.graphClient = Client.init({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });
  }

  async testConnection() {
    try {
      console.log('🔍 Testing Graph API connection...');
      const user = await this.graphClient.api('/me').get();
      console.log('✅ Connection successful!');
      console.log(`👤 Connected as: ${user.displayName} (${user.userPrincipalName})`);
      return true;
    } catch (error) {
      console.error('❌ Connection failed:', error.message);
      return false;
    }
  }

  async listNotebooks() {
    try {
      console.log('\n📚 Fetching OneNote notebooks...');
      
      const response = await this.graphClient
        .api('/me/onenote/notebooks')
        .get();

      const notebooks = response.value || [];

      if (notebooks.length === 0) {
        console.log('📝 No notebooks found.');
        return [];
      }

      console.log(`\n📊 Found ${notebooks.length} notebook(s):\n`);
      
      notebooks.forEach((notebook, index) => {
        console.log(`${index + 1}. 📓 "${notebook.displayName}"`);
        console.log(`   📋 ID: ${notebook.id}`);
        console.log(`   📅 Created: ${new Date(notebook.createdDateTime).toLocaleDateString()}`);
        console.log(`   📝 Last Modified: ${new Date(notebook.lastModifiedDateTime).toLocaleDateString()}`);
        console.log(`   ⭐ Default: ${notebook.isDefault ? 'Yes' : 'No'}`);
        
        if (notebook.links && notebook.links.oneNoteWebUrl) {
          console.log(`   🌐 Web URL: ${notebook.links.oneNoteWebUrl.href}`);
        }
        
        console.log(''); // Empty line
      });

      return notebooks;
      
    } catch (error) {
      console.error('❌ Failed to fetch notebooks:', error.message);
      console.error('💡 Make sure your token has Notes.Read permission');
      return [];
    }
  }

  async getNotebookSections(notebookId, notebookName) {
    try {
      console.log(`📂 Getting sections for "${notebookName}"...`);
      
      const response = await this.graphClient
        .api(`/me/onenote/notebooks/${notebookId}/sections`)
        .get();

      const sections = response.value || [];
      
      console.log(`   Found ${sections.length} section(s):`);
      sections.forEach((section, index) => {
        console.log(`   ${index + 1}. 📄 ${section.displayName}`);
      });
      
      return sections;
      
    } catch (error) {
      console.error(`❌ Failed to get sections: ${error.message}`);
      return [];
    }
  }

  async run() {
    console.log('🚀 Simple OneNote Test Starting...\n');
    
    // Test connection
    const connected = await this.testConnection();
    if (!connected) {
      console.log('\n💡 To get an access token:');
      console.log('   1. Go to: https://developer.microsoft.com/en-us/graph/graph-explorer');
      console.log('   2. Sign in with your Microsoft account');
      console.log('   3. Run any query (like GET /me)');
      console.log('   4. Copy the access token from the "Access token" tab');
      console.log('   5. Replace TOKEN in this file with your token');
      return;
    }

    // List notebooks
    const notebooks = await this.listNotebooks();
    
    // Get sections for first notebook
    if (notebooks.length > 0) {
      console.log('📂 Getting sections for first notebook...\n');
      await this.getNotebookSections(notebooks[0].id, notebooks[0].displayName);
    }

    console.log('\n✅ Test completed!');
    console.log('\n🔗 Next steps for your Outlook add-in:');
    console.log('   - These same API endpoints work in your add-in');
    console.log('   - Replace token auth with Office.auth.getAccessToken()');
    console.log('   - Add error handling for production use');
    console.log('   - Consider caching notebooks for better UX');
  }
}

// Helper to check dependencies
function checkDependencies() {
  try {
    require('@microsoft/microsoft-graph-client');
    return true;
  } catch (error) {
    console.error('❌ Missing Microsoft Graph Client. Install it with:');
    console.error('   npm install @microsoft/microsoft-graph-client\n');
    return false;
  }
}

// Main execution
async function main() {
  console.log('Simple OneNote Test - Microsoft Graph API');
  console.log('=' .repeat(45));
  
  if (!checkDependencies()) {
    process.exit(1);
  }
  
  try {
    const test = new SimpleOneNoteTest(TOKEN);
    await test.run();
  } catch (error) {
    if (error.message.includes('provide a valid access token')) {
      console.error('❌ Configuration Error:', error.message);
      console.log('\n📋 Instructions:');
      console.log('   1. Open: https://developer.microsoft.com/en-us/graph/graph-explorer');
      console.log('   2. Sign in with your Microsoft account');
      console.log('   3. Try the query: GET https://graph.microsoft.com/v1.0/me');
      console.log('   4. Go to "Access token" tab and copy the token');
      console.log('   5. Replace TOKEN variable in this file');
      console.log('   6. Run the script again');
    } else {
      console.error('❌ Error:', error.message);
    }
    process.exit(1);
  }
}

if (require.main === module) {
  main().catch(console.error);
}

module.exports = SimpleOneNoteTest;
