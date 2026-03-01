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
  'Mail.Read',
  'Notes.ReadWrite',
  'offline_access',
  'User.Read',
]
