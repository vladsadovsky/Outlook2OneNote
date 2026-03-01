import { describe, it, expect, vi, beforeEach } from 'vitest'
import { PublicClientApplication } from '@azure/msal-browser'
import type { AccountInfo, AuthenticationResult } from '@azure/msal-browser'

// ---------------------------------------------------------------------------
// Hoist mock state so vi.mock factory can close over it
// ---------------------------------------------------------------------------
const mockMsal = vi.hoisted(() => ({
  initialize: vi.fn<() => Promise<void>>(),
  getAllAccounts: vi.fn<() => AccountInfo[]>(),
  acquireTokenSilent: vi.fn<() => Promise<AuthenticationResult>>(),
  acquireTokenPopup: vi.fn<() => Promise<AuthenticationResult>>(),
}))

vi.mock('@azure/msal-browser', () => ({
  PublicClientApplication: vi.fn(() => mockMsal),
  InteractionRequiredAuthError: class InteractionRequiredAuthError extends Error {
    constructor(message?: string) {
      super(message)
      this.name = 'InteractionRequiredAuthError'
    }
  },
}))

// msalConfig imports import.meta.env — mock the module to avoid Vite env dependency
vi.mock('@/auth/msalConfig', () => ({
  msalConfig: { auth: { clientId: 'test-client', authority: 'https://login.microsoftonline.com/common' } },
  graphScopes: ['Mail.Read', 'Notes.ReadWrite'],
}))

// ---------------------------------------------------------------------------
// Import after mocks are wired
// ---------------------------------------------------------------------------
import { getGraphToken, AuthError } from '@/auth/authService'
import { InteractionRequiredAuthError } from '@azure/msal-browser'

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const TOKEN_KEY = 'o2on_token'
const EXPIRY_KEY = 'o2on_token_expiry'

function makeJwt(payload: Record<string, string>): string {
  const header = btoa('{}')
  const body = btoa(JSON.stringify(payload))
  return `${header}.${body}.signature`
}

function makeAuthResult(accessToken: string, expiresInMs = 3600_000): AuthenticationResult {
  return {
    accessToken,
    expiresOn: new Date(Date.now() + expiresInMs),
  } as unknown as AuthenticationResult
}

function makeAccount(username: string): AccountInfo {
  return { username, homeAccountId: 'home', environment: 'login.microsoftonline.com', tenantId: 't', localAccountId: 'l' }
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------
let mockGetAccessToken: ReturnType<typeof vi.fn>

beforeEach(() => {
  vi.resetAllMocks()
  sessionStorage.clear()

  // vi.resetAllMocks() clears the PublicClientApplication factory — restore it
  vi.mocked(PublicClientApplication).mockImplementation(() => mockMsal as unknown as PublicClientApplication)

  // Default MSAL behaviours
  mockMsal.initialize.mockResolvedValue(undefined)
  mockMsal.getAllAccounts.mockReturnValue([])

  // Default SSO: unavailable
  mockGetAccessToken = vi.fn().mockRejectedValue(new Error('SSO unavailable'))
  ;(globalThis as Record<string, unknown>)['Office'] = {
    auth: { getAccessToken: mockGetAccessToken },
  }
})

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------
describe('getGraphToken', () => {
  it('returns cached token without calling MSAL', async () => {
    sessionStorage.setItem(TOKEN_KEY, 'cached-token')
    sessionStorage.setItem(EXPIRY_KEY, String(Date.now() + 3600_000))

    const token = await getGraphToken()

    expect(token).toBe('cached-token')
    expect(mockMsal.acquireTokenSilent).not.toHaveBeenCalled()
    expect(mockMsal.acquireTokenPopup).not.toHaveBeenCalled()
  })

  it('treats expired cached token as absent and falls through to popup', async () => {
    sessionStorage.setItem(TOKEN_KEY, 'old-token')
    // expired 10 minutes ago
    sessionStorage.setItem(EXPIRY_KEY, String(Date.now() - 10 * 60 * 1000))
    mockMsal.acquireTokenPopup.mockResolvedValue(makeAuthResult('fresh-token'))

    const token = await getGraphToken()

    expect(token).toBe('fresh-token')
  })

  it('uses silent token when SSO provides login hint and account exists', async () => {
    const jwt = makeJwt({ preferred_username: 'user@example.com' })
    mockGetAccessToken.mockResolvedValue(jwt)
    mockMsal.getAllAccounts.mockReturnValue([makeAccount('user@example.com')])
    mockMsal.acquireTokenSilent.mockResolvedValue(makeAuthResult('silent-token'))

    const token = await getGraphToken()

    expect(token).toBe('silent-token')
    expect(mockMsal.acquireTokenPopup).not.toHaveBeenCalled()
  })

  it('uses silent token when SSO is unavailable but an MSAL account exists', async () => {
    // SSO unavailable (default mock)
    mockMsal.getAllAccounts.mockReturnValue([makeAccount('other@example.com')])
    mockMsal.acquireTokenSilent.mockResolvedValue(makeAuthResult('silent-token'))

    const token = await getGraphToken()

    expect(token).toBe('silent-token')
    expect(mockMsal.acquireTokenPopup).not.toHaveBeenCalled()
  })

  it('falls back to popup when silent throws InteractionRequiredAuthError', async () => {
    const jwt = makeJwt({ preferred_username: 'user@example.com' })
    mockGetAccessToken.mockResolvedValue(jwt)
    mockMsal.getAllAccounts.mockReturnValue([makeAccount('user@example.com')])
    mockMsal.acquireTokenSilent.mockRejectedValue(
      new InteractionRequiredAuthError('interaction_required'),
    )
    mockMsal.acquireTokenPopup.mockResolvedValue(makeAuthResult('popup-token'))

    const token = await getGraphToken()

    expect(token).toBe('popup-token')
  })

  it('falls back to popup when no MSAL accounts exist', async () => {
    // SSO unavailable (default), no accounts
    mockMsal.acquireTokenPopup.mockResolvedValue(makeAuthResult('popup-token'))

    const token = await getGraphToken()

    expect(token).toBe('popup-token')
    expect(mockMsal.acquireTokenSilent).not.toHaveBeenCalled()
  })

  it('throws AuthError when popup is rejected', async () => {
    mockMsal.acquireTokenPopup.mockRejectedValue(new Error('user closed popup'))

    await expect(getGraphToken()).rejects.toThrow(AuthError)
  })

  it('caches token in sessionStorage after acquisition', async () => {
    mockMsal.acquireTokenPopup.mockResolvedValue(makeAuthResult('new-token'))

    await getGraphToken()

    expect(sessionStorage.getItem(TOKEN_KEY)).toBe('new-token')
    expect(sessionStorage.getItem(EXPIRY_KEY)).not.toBeNull()
  })

  it('passes loginHint to acquireTokenPopup when SSO provides one', async () => {
    const jwt = makeJwt({ preferred_username: 'user@example.com' })
    mockGetAccessToken.mockResolvedValue(jwt)
    // No accounts → goes straight to popup
    mockMsal.acquireTokenPopup.mockResolvedValue(makeAuthResult('popup-token'))

    await getGraphToken()

    expect(mockMsal.acquireTokenPopup).toHaveBeenCalledWith(
      expect.objectContaining({ loginHint: 'user@example.com' }),
    )
  })
})
