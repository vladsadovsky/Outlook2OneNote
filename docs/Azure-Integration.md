# Azure Integration Guide (Outlook2OneNote)

This document covers Azure-side setup and Graph verification for Outlook2OneNote.

## 1) What this guide is for

Use this guide to:
- Register or validate the app in Azure
- Configure Graph delegated permissions
- Align local `.env` settings with account policy
- Debug auth, mail, and OneNote access issues with Graph Explorer

---

## 2) App registration

1. Open Azure Portal: https://portal.azure.com
2. Go to **App registrations** → **New registration**
3. Name: `Outlook2OneNote (dev)`
4. Choose supported account types based on deployment:
   - Personal-only: **Personal Microsoft accounts only**
   - Entra-only: **Accounts in this organizational directory only** (single-tenant) or **any organizational directory** (multi-tenant)
   - Mixed deployments: **Accounts in any organizational directory and personal Microsoft accounts**
5. Add SPA redirect URI:
   - `https://localhost:3000/auth/callback.html`
6. Save and copy the Application (client) ID for `VITE_CLIENT_ID`

No client secret is required (SPA/public-client flow).

---

## 3) Required Microsoft Graph delegated permissions

Add these delegated permissions:
- `Mail.Read`
- `Notes.Read`
- `Notes.ReadWrite`
- `User.Read`
- `offline_access`

For enterprise tenants, admin consent is commonly required.

---

## 4) Environment alignment

Ensure `.env` matches deployment policy:
- `VITE_CLIENT_ID=<your-app-client-id>`
- `VITE_ACCOUNT_TYPE_POLICY=all | msa-only | entra-only`
- `VITE_TENANT_ID=common` (or tenant-specific value when required)

Examples:
- Personal-only deployment: `VITE_ACCOUNT_TYPE_POLICY=msa-only`
- Entra-only deployment: `VITE_ACCOUNT_TYPE_POLICY=entra-only`
- Mixed deployment: `VITE_ACCOUNT_TYPE_POLICY=all`

---

## 5) Debugging with Microsoft Graph Explorer

Use https://developer.microsoft.com/en-us/graph/graph-explorer and sign in with the same account used in add-in testing (not the sample tenant account).

Recommended request sequence:
1. `GET /v1.0/me`
   - Verifies account identity and baseline token validity.
2. `GET /v1.0/me/onenote/notebooks`
   - Verifies OneNote access and permissions.
3. `GET /v1.0/me/messages?$filter=conversationId eq '{conversationId}'&$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments&$top=51&$orderby=receivedDateTime asc`
   - Verifies thread query behavior used by export.
4. If `InefficientFilter` occurs, retry without `$orderby` to confirm query-shape limitations.

If Graph Explorer succeeds but the add-in fails, focus on local token flow, `.env` policy mismatch, or request-shape differences in code.

---

## 6) Common issue patterns

- `401 Unauthorized` in Graph Explorer:
  - Usually not signed in with your target account or missing consent.
- `30121 SharePoint license` for OneNote endpoints:
  - Verify correct account context and granted scopes; compare with Graph Explorer under same account.
- `InefficientFilter` on mail queries:
  - Use simplified query fallback (no `$orderby`) and sort client-side.
