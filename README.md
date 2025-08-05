
# Outlook to OneNote Web Add-in

This project is---

### 5. Test the Plugin

- Open an email thread
- Click "Export to OneNote" from the ribbon
- Authenticate with Microsoft account
- Select notebook → A new section is created with one page per message

---

## 🐛 Debugging the Add-in

### Prerequisites for Debugging

1. **VS Code** with the following extensions installed:
   - JavaScript Debugger (built-in)
   - Microsoft Edge Tools for VS Code (optional but recommended)

2. **Development Server Running**:
   ```cmd
   npm run dev-server
   ```
   This starts the webpack dev server with source maps on `https://localhost:3000`

### Debugging Methods

#### Method 1: VS Code Debugging (Recommended)

1. **Start the Development Server**:
   ```cmd
   npm run dev-server
   ```

2. **Set Breakpoints**:
   - Open `src/taskpane/taskpane.js` in VS Code
   - Click in the gutter next to line numbers to set breakpoints
   - You should see red dots indicating active breakpoints

3. **Launch Debugger**:
   - Go to Run and Debug (Ctrl+Shift+D)
   - Select one of these configurations:
     - "Launch Edge against localhost (Office Add-in)"
     - "Launch Chrome against localhost (Office Add-in)"
   - Press F5 or click the green play button

4. **Attach to Outlook Web**:
   - If using the attach configuration, first open Outlook Web
   - Sideload your add-in
   - Use "Attach to Edge (Outlook Add-in)" configuration
   - VS Code will attach to the browser process

#### Method 2: Browser Developer Tools

1. **Open Outlook Web** and sideload your add-in

2. **Open Developer Tools**:
   - Press F12 or right-click → Inspect
   - Navigate to the **Sources** tab

3. **Find Your Source Files**:
   - Look for `webpack://` in the file tree
   - Navigate to `localhost:3000` → `src/taskpane/taskpane.js`
   - The original source files should be visible thanks to source maps

4. **Set Breakpoints**:
   - Click on line numbers in the source files
   - Blue dots indicate active breakpoints

### Debugging Configuration Files

The project includes pre-configured debugging settings:

**`.vscode/launch.json`**:
```json
{
  "configurations": [
    {
      "type": "msedge",
      "request": "launch",
      "name": "Launch Edge against localhost (Office Add-in)",
      "url": "https://localhost:3000/taskpane.html",
      "webRoot": "${workspaceFolder}",
      "sourceMaps": true,
      "trace": true
    },
    {
      "name": "Attach to Edge (Outlook Add-in)",
      "type": "msedge",
      "request": "attach",
      "port": 9222,
      "webRoot": "${workspaceFolder}",
      "sourceMaps": true,
      "timeout": 10000,
      "trace": true
    }
  ]
}
```

**`.vscode/settings.json`**:
```json
{
  "debug.allowBreakpointsEverywhere": true,
  "debug.javascript.unmapMissingSources": true,
  "debug.javascript.suggestPrettyPrinting": false
}
```

### Webpack Configuration for Debugging

The project uses different source map strategies:
- **Development**: `eval-source-map` (faster rebuilds, better debugging experience)
- **Production**: `source-map` (smaller files, slower builds)

### Common Debugging Issues & Solutions

#### 1. Breakpoints Show Empty Circles
**Problem**: Breakpoints appear as empty/hollow circles instead of filled red dots.

**Solutions**:
- Ensure the development server is running (`npm run dev-server`)
- Check that you're debugging the correct URL (`https://localhost:3000`)
- Verify source maps are enabled in webpack config
- Clear browser cache and restart VS Code

#### 2. Source Files Not Found
**Problem**: VS Code can't find the original source files.

**Solutions**:
- Check that `webRoot` in launch.json points to your workspace folder
- Ensure webpack is generating source maps (`devtool` configuration)
- Try refreshing the browser and reattaching the debugger

#### 3. Debugger Won't Attach to Browser
**Problem**: "Attach to Edge" configuration fails.

**Solutions**:
- Ensure Edge is launched with debugging enabled:
  ```cmd
  msedge --remote-debugging-port=9222
  ```
- Check that no other debugger is attached to the same port
- Try using Chrome instead of Edge

#### 4. Office.js Not Loading
**Problem**: Office.js APIs are undefined.

**Solutions**:
- Ensure your add-in is properly sideloaded in Outlook
- Check that the manifest.xml URLs match your development server
- Verify the add-in is running in the correct Office context

### Debugging Workflow Tips

1. **Use Console Logging**:
   ```javascript
   console.log("Outlook2OneNote::Office.onReady", info);
   console.log("Current mail item:", Office.context.mailbox.item);
   ```

2. **Check Network Tab**:
   - Monitor API calls to Microsoft Graph
   - Verify authentication tokens
   - Check for CORS issues

3. **Test in Different Environments**:
   - Outlook Web (Chrome, Edge, Firefox)
   - Outlook Desktop (if applicable)
   - Different Office 365 tenants

4. **Use Office.js Debugging**:
   ```javascript
   // Enable verbose Office.js logging
   Office.onReady((info) => {
     if (Office.context.requirements.isSetSupported('Mailbox', '1.1')) {
       console.log("Mailbox API 1.1 supported");
     }
   });
   ```

### Performance Debugging

1. **Check Bundle Size**:
   ```cmd
   npm run build
   # Check dist/ folder for bundle sizes
   ```

2. **Analyze Webpack Bundle**:
   ```cmd
   npx webpack-bundle-analyzer dist/
   ```

3. **Monitor Memory Usage**:
   - Use browser's Memory tab in DevTools
   - Check for memory leaks in long-running sessions

---

## 🔧 Conditional Compilation (Development vs Production)

This project uses webpack's `DefinePlugin` to implement conditional compilation, allowing different features and behavior between development and production builds.

### How It Works

1. **Webpack Configuration**: The `__DEV__` constant is defined based on the build mode:
   ```javascript
   new webpack.DefinePlugin({
     __DEV__: JSON.stringify(dev),  // true in development, false in production
     'process.env.NODE_ENV': JSON.stringify(dev ? 'development' : 'production')
   })
   ```

2. **Conditional Code Blocks**: Debug features are wrapped in conditional statements:
   ```javascript
   if (__DEV__) {
     // Debug-only code here - will be completely removed in production builds
     window.testRun = async function() { /* ... */ }
   }
   ```

3. **UI Elements**: The TestRun button is hidden in production:
   ```javascript
   if (__DEV__) {
     document.getElementById("testrun").onclick = window.testRun;
   } else {
     document.getElementById("testrun").style.display = "none";
   }
   ```

### Benefits

- **Smaller Production Bundles**: Debug code is completely eliminated from production builds
- **Better Performance**: No runtime checks for debug features in production
- **Secure**: Debug functionality cannot be accessed in production environment
- **Developer Experience**: Full debugging capabilities during development

### Build Modes

| Command | Mode | __DEV__ | TestRun Button | Features |
|---------|------|---------|----------------|----------|
| `npm run dev-server` | development | `true` | Visible & Functional | All debug features |
| `npm run build:dev` | development | `true` | Visible & Functional | All debug features |
| `npm run build` | production | `false` | Hidden | Production only |
| `npm run watch` | development | `true` | Visible & Functional | All debug features |

--- 365 Outlook Web Add-in that allows users to export an email thread to the OneNote notebook using the Microsoft Graph API.

---

## 🚀 Features

- Adds a ribbon button to Outlook Web/Desktop
- Lets users choose a OneNote notebook via Microsoft Graph
- Creates a section and adds one page per email in the selected thread

### 🛠 Developer Features (Debug Mode Only)

- **TestRun Function**: Available only in development builds for debugging purposes
  - Displays the current email subject and sender information
  - Retrieves and shows all emails in the currently selected thread/conversation
  - Uses Exchange Web Services (EWS) to analyze conversation structure
  - Provides detailed thread analysis including subjects, senders, and timestamps
  - **Note**: This feature is automatically excluded from production builds

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

### 3. Build and Launch

#### Development Mode (with debug features)
```cmd
npm run dev-server
```
This starts the webpack dev server in development mode with:
- Debug features enabled (TestRun button visible)
- Hot reloading for faster development
- Better source maps for debugging

#### Production Build
```cmd
npm run build
```
This creates an optimized production build with:
- Debug features excluded (TestRun button hidden)
- Minified and optimized code
- No development dependencies

#### Alternative Development Commands
```cmd
npm start          # Launches development server
npm run build:dev  # Creates development build without server
npm run watch      # Watches for changes and rebuilds automatically
``` 


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

```

---

## 🧪 Notes

- Make sure Outlook trusts the self-signed cert
- If using Edge/Chrome, allow loading from `localhost` HTTPS
- For deployment, you must host over a public HTTPS domain with valid cert


