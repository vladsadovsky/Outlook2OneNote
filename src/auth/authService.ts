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
    
    // Also log account type info for debugging
    const tid = payload['tid']  // Tenant ID
    const accountType = tid === '9188040d-6c67-4c5b-b112-36a304b66dad' ? 'Personal' : 'Work/School'
    debugLog('auth', `Account type: ${accountType} (tenant: ${tid})`)
    
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
  // Also clear localStorage (some browsers use this)
  localStorage.removeItem(TOKEN_KEY)
  localStorage.removeItem(EXPIRY_KEY)
  debugLog('auth', 'All cached tokens cleared - forcing fresh authentication')
}

// Returns a valid Graph API bearer token.
// All other modules call this — never touch MSAL or Office.auth directly.
export async function getGraphToken(): Promise<string> {
  console.log('[authService] getGraphToken called')
  debugLog('auth', 'getGraphToken called')
  
  const cached = getCachedToken()
  if (cached) {
    console.log('[authService] Using cached token')
    debugLog('auth', 'Using cached token')
    return cached
  }

  debugLog('auth', 'No cached token, acquiring new one')
  console.log('[authService] No cached token, acquiring new one')
  
  try {
    const msal = await getMsal()
    debugLog('auth', 'MSAL initialized successfully')
    
    const loginHint = await getLoginHint()
    debugLog('auth', `Login hint: ${loginHint || 'none'}`)

    // Skip silent authentication and go straight to popup
    console.log('[authService] Attempting popup authentication')
    debugLog('auth', 'Attempting popup authentication...')
    
    const result = await msal.acquireTokenPopup({ 
      scopes: graphScopes, 
      loginHint
    })
    console.log('[authService] Fresh popup token acquired successfully')
    debugLog('auth', 'Popup authentication successful!')
    
    setCachedToken(result.accessToken, result.expiresOn)
    return result.accessToken
  } catch (e) {
    debugError('auth', `Fresh auth error: ${e instanceof Error ? e.message : 'Unknown error'}`)
    throw new AuthError(
      'Authentication failed. Please try again or contact your administrator.',
    )
  }
}
