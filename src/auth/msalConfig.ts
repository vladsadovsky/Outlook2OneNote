import type { Configuration } from '@azure/msal-browser'

export type AccountTypePolicy = 'all' | 'msa-only' | 'entra-only'

const ACCOUNT_TYPE_POLICY: AccountTypePolicy =
  (import.meta.env.VITE_ACCOUNT_TYPE_POLICY as AccountTypePolicy) ?? 'all'

const authorityMap: Record<AccountTypePolicy, string> = {
  'all': 'https://login.microsoftonline.com/common',
  'msa-only': 'https://login.microsoftonline.com/consumers',
  'entra-only': 'https://login.microsoftonline.com/organizations',
}

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID as string,
    authority: authorityMap[ACCOUNT_TYPE_POLICY],
    redirectUri: `${window.location.origin}/auth/callback.html`,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
}

export const graphScopes = [
  'https://graph.microsoft.com/Notes.Read',          // ← Added missing scope (v1 format)!
  'https://graph.microsoft.com/Notes.ReadWrite',     // ← Changed to v1 format
  'https://graph.microsoft.com/User.Read',           // ← Changed to v1 format  
  'https://graph.microsoft.com/Mail.Read',           // ← Keep for email functionality
  'offline_access',                                  // ← Keep for refresh tokens
]
