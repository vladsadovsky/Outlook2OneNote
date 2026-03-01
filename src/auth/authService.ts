import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser'
import { msalConfig, graphScopes } from './msalConfig'
import { debugLog, debugError } from '@/utils/logger'

export class AuthError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'AuthError'
  }
}

const TOKEN_KEY = 'o2on_token'
const EXPIRY_KEY = 'o2on_token_expiry'
const EXPIRY_BUFFER_MS = 5 * 60 * 1000 // 5 min — refresh before actual expiry

let _msal: PublicClientApplication | null = null

async function getMsal(): Promise<PublicClientApplication> {
  if (!_msal) {
    _msal = new PublicClientApplication(msalConfig)
    await _msal.initialize()
  }
  return _msal
}

function getCachedToken(): string | null {
  const token = sessionStorage.getItem(TOKEN_KEY)
  const expiry = sessionStorage.getItem(EXPIRY_KEY)
  if (!token || !expiry) return null
  if (Date.now() > Number(expiry) - EXPIRY_BUFFER_MS) {
    sessionStorage.removeItem(TOKEN_KEY)
    sessionStorage.removeItem(EXPIRY_KEY)
    return null
  }
  return token
}

function setCachedToken(token: string, expiresOn: Date | null): void {
  const expiry = expiresOn ? expiresOn.getTime() : Date.now() + 3600_000
  sessionStorage.setItem(TOKEN_KEY, token)
  sessionStorage.setItem(EXPIRY_KEY, String(expiry))
}

// JWT payloads are base64url encoded — convert to standard base64 before atob
function parseJwtPayload(token: string): Record<string, string> {
  const base64Url = token.split('.')[1] ?? ''
  const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/')
  return JSON.parse(atob(base64)) as Record<string, string>
}

// Returns a login hint (email) from Office SSO bootstrap token, or undefined if SSO unavailable.
// Does NOT attempt to use the SSO token for Graph — that requires OBO (T-OBO).
async function getLoginHint(): Promise<string | undefined> {
  try {
    const ssoToken = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
    })
    const payload = parseJwtPayload(ssoToken)
    const hint = payload['preferred_username'] ?? payload['upn']
    debugLog('auth', 'SSO login hint', hint)
    return hint
  } catch {
    debugLog('auth', 'SSO not available — proceeding to MSAL')
    return undefined
  }
}

// Evicts the cached token. Call before a 401-retry so that the next
// getGraphToken() call acquires a fresh token from MSAL.
export function clearCachedToken(): void {
  sessionStorage.removeItem(TOKEN_KEY)
  sessionStorage.removeItem(EXPIRY_KEY)
}

// Returns a valid Graph API bearer token.
// All other modules call this — never touch MSAL or Office.auth directly.
export async function getGraphToken(): Promise<string> {
  const cached = getCachedToken()
  if (cached) {
    debugLog('auth', 'Using cached token')
    return cached
  }

  const msal = await getMsal()
  const loginHint = await getLoginHint()

  // Try silent acquisition first (avoids popup if user already signed in)
  try {
    const accounts = msal.getAllAccounts()
    const account = loginHint
      ? (accounts.find(a => a.username === loginHint) ?? accounts[0])
      : accounts[0]
    if (account) {
      const result = await msal.acquireTokenSilent({ scopes: graphScopes, account })
      debugLog('auth', 'Silent token acquired')
      setCachedToken(result.accessToken, result.expiresOn)
      return result.accessToken
    }
  } catch (e) {
    if (!(e instanceof InteractionRequiredAuthError)) {
      debugError('auth', 'Unexpected silent failure', e)
    }
    debugLog('auth', 'Silent failed — trying popup')
  }

  // Popup fallback (single popup, matches native "Save to OneNote" UX)
  try {
    const result = await msal.acquireTokenPopup({ scopes: graphScopes, loginHint })
    debugLog('auth', 'Popup token acquired')
    setCachedToken(result.accessToken, result.expiresOn)
    return result.accessToken
  } catch (e) {
    debugError('auth', 'Popup failed', e)
    throw new AuthError(
      'Sign-in failed. Please try again or contact your administrator.',
    )
  }
}
