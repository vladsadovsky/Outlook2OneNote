# Azure App Registration (Outlook2OneNote)

This guide covers the exact Azure app registration steps for Outlook2OneNote.

## Personal Microsoft accounts (most common)

Use this if your users sign in with a personal Microsoft account (outlook.com, hotmail.com, live.com).

### 1) Create the app registration

1. Go to https://portal.azure.com
2. Search for and open "App registrations"
3. Click "New registration"
4. Name: Outlook2OneNote (dev)
5. Supported account types:
   - Choose "Personal Microsoft accounts only"
6. Redirect URI:
   - Platform: "Single-page application (SPA)"
   - URI: https://localhost:3000/auth/callback.html
7. Click "Register"
8. Copy the "Application (client) ID" (used in .env.local as VITE_CLIENT_ID)

### 2) Add Microsoft Graph permissions

1. In your app registration, open "API permissions"
2. Click "Add a permission" -> "Microsoft Graph" -> "Delegated permissions"
3. Add these permissions:
   - Mail.Read
   - Notes.ReadWrite
   - User.Read
   - offline_access
4. Click "Add permissions"

Note: For personal accounts, admin consent is typically not required.

---

## Work or school accounts (Entra ID)

Use this if your users sign in with a work or school account.

### 1) Create the app registration

1. Go to https://portal.azure.com
2. Search for and open "App registrations"
3. Click "New registration"
4. Name: Outlook2OneNote (dev)
5. Supported account types:
   - Choose "Accounts in this organizational directory only" (single-tenant)
     or
   - "Accounts in any organizational directory" (multi-tenant)
6. Redirect URI:
   - Platform: "Single-page application (SPA)"
   - URI: https://localhost:3000/auth/callback.html
7. Click "Register"
8. Copy the "Application (client) ID" (used in .env.local as VITE_CLIENT_ID)

### 2) Add Microsoft Graph permissions

1. In your app registration, open "API permissions"
2. Click "Add a permission" -> "Microsoft Graph" -> "Delegated permissions"
3. Add these permissions:
   - Mail.Read
   - Notes.ReadWrite
   - User.Read
   - offline_access
4. Click "Add permissions"

### 3) Grant admin consent (often required)

Many tenants require admin consent for the Graph scopes above.

1. In "API permissions", click "Grant admin consent"
2. A tenant admin must approve

---

## Notes

- No client secret is required for this add-in (public client, SPA flow).
- Redirect URI must match exactly: https://localhost:3000/auth/callback.html
- For personal accounts only, set VITE_ACCOUNT_TYPE_POLICY=msa-only in .env.local
- For Entra-only deployments, set VITE_ACCOUNT_TYPE_POLICY=entra-only
