# Outlook2OneNote v2

Export Outlook email conversation threads to OneNote, one page per email.

## Quick links

- docs/SPEC.md
- docs/DESIGN.md
- docs/TEST-PLAN.md
- docs/AZURE-APP-REG.md
- manifest.xml

## What this add-in does

- Reads the selected email thread by conversationId
- Creates a new OneNote section in a chosen notebook
- Creates one page per email, ordered chronologically
- Shows progress and a link to the created section

## First-time developer setup (personal Microsoft account)

This guide assumes no prior Azure or Office add-in setup experience. Follow every step in order.

### 1) Prerequisites

- Node.js (recent LTS is recommended)
- npm (comes with Node)
- An Outlook account (personal Microsoft account)
- OneNote access for the same account

### 2) Clone and install

```bash
npm install
```

### 3) Create and trust the local HTTPS dev certificate

The add-in runs on https://localhost:3000. Outlook requires HTTPS.

```bash
npx office-addin-dev-certs install
```

If you need to verify later:

```bash
npx office-addin-dev-certs verify
```

### 4) Register an Azure app (personal accounts only)

These steps create the App ID used by MSAL for login.

Follow the full walkthrough in docs/AZURE-APP-REG.md (Personal Microsoft accounts section).

### 5) Create the .env file

Create a file named .env.local in the repo root:

```bash
VITE_CLIENT_ID=YOUR-CLIENT-ID-HERE
VITE_ACCOUNT_TYPE_POLICY=msa-only
```

- VITE_CLIENT_ID is the Application (client) ID from Azure.
- VITE_ACCOUNT_TYPE_POLICY=msa-only forces personal account login.

### 6) Start the dev server

```bash
npm run dev
```

Vite serves the add-in at https://localhost:3000.

### 7) Sideload the add-in (Outlook Web)

This uses manifest.xml, per this repo's current workflow.

1. Open Outlook on the web: https://outlook.office.com
2. Click the gear icon -> "View all Outlook settings"
3. Go to "Mail" -> "Customize actions" or "Add-ins" (UI varies)
4. Find "Manage add-ins" or "Custom add-ins"
5. Choose "Add a custom add-in" -> "Add from file"
6. Upload manifest.xml from the repo root
7. Confirm the install

### 8) Sideload the add-in (Outlook New)

Outlook New uses the same add-ins page as Outlook Web. If you are signed into the same account:

1. Open Outlook New
2. Go to "Get Add-ins"
3. Open "My add-ins"
4. Add a custom add-in from file
5. Upload manifest.xml

### 9) Sideload the add-in (Outlook Classic)

Outlook Classic support is not guaranteed. If your Classic client supports modern web add-ins:

1. Open Outlook Classic
2. Go to File -> Manage Add-ins
3. This opens Outlook on the web add-ins page
4. Follow the Outlook Web sideload steps above

If the task pane does not appear in Classic, use Outlook Web or Outlook New for evaluation.

## Work or school accounts (Entra ID) setup

Most users will be personal accounts. If you need Entra ID support for a specific tenant, follow the separate steps below.

### 1) Register an Azure app (work or school accounts)

Follow the full walkthrough in docs/AZURE-APP-REG.md (Work or school accounts section).

### 2) Update .env.local

```bash
VITE_CLIENT_ID=YOUR-CLIENT-ID-HERE
VITE_ACCOUNT_TYPE_POLICY=entra-only
```

If you want both personal and work accounts, set VITE_ACCOUNT_TYPE_POLICY=all.

## Evaluating the add-in (manual test walk-through)

1. Ensure the dev server is running (npm run dev)
2. Open an email in Outlook (Web or New)
3. Click "Export to OneNote" in the ribbon
4. If prompted, sign in with your personal Microsoft account
5. Pick a notebook in the task pane
6. Click Export
7. Confirm:
   - A new OneNote section was created
   - Each email becomes a page
   - Pages are ordered oldest to newest

If export fails, check the error banner in the task pane and the browser console.

## Development scripts

- npm run dev: start Vite dev server
- npm run build: type-check and build
- npm run test: run unit tests (Vitest)
- npm run lint: run ESLint

## Configuration

### Environment variables

- VITE_CLIENT_ID (required)
  - Azure App Registration Application (client) ID
- VITE_ACCOUNT_TYPE_POLICY (optional)
  - all (default), msa-only, entra-only
  - For personal accounts only, set msa-only

## Troubleshooting

- Add-in not visible in ribbon:
  - Confirm manifest.xml upload succeeded
  - Ensure Outlook is in Mail Read view (reading a message)
- HTTPS errors in task pane:
  - Re-run npx office-addin-dev-certs install
  - Restart the browser/Outlook after trust changes
- Login popup blocked:
  - Allow popups for outlook.office.com
- Authentication fails:
  - Ensure VITE_CLIENT_ID matches the Azure app registration
  - Confirm redirect URI is https://localhost:3000/auth/callback.html
  - Confirm Graph permissions include Mail.Read, Notes.ReadWrite, User.Read, offline_access

## Project documentation

- docs/SPEC.md
- docs/DESIGN.md
- docs/TEST-PLAN.md
- docs/TASKS.md

## License

See LICENSE.
