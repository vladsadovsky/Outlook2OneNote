import { getGraphToken, clearCachedToken } from '@/auth/authService'
import { debugLog, debugError } from '@/utils/logger'

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const GRAPH_BETA_BASE = 'https://graph.microsoft.com/beta'

export interface GraphClient {
  get<T>(path: string, params?: Record<string, string>): Promise<T>
  post<T>(path: string, body: unknown, contentType?: string): Promise<T>
}

// Internal: makes one fetch attempt with a given token.
// Returns the Response or throws on network error.
async function attempt(url: string, init: RequestInit, token: string): Promise<Response> {
  return fetch(url, {
    ...init,
    headers: {
      'Authorization': `Bearer ${token}`,
      ...(init.headers as Record<string, string> | undefined),
    },
  })
}

// Resolves an absolute or relative Graph path to a full URL with optional query params.
function resolveUrl(path: string, params?: Record<string, string>): string {
  let base: string
  if (path.startsWith('http')) {
    base = path
  } else if (path.startsWith('/beta/')) {
    base = `${GRAPH_BETA_BASE}${path.substring(5)}` // Remove '/beta' and use beta base
  } else {
    base = `${GRAPH_BASE}${path}`
  }
  
  if (!params) return base
  return `${base}?${new URLSearchParams(params)}`
}

// Core request helper: attaches Bearer token, retries once on 401, respects 429 Retry-After.
async function graphRequest(url: string, init: RequestInit): Promise<Response> {
  let token = await getGraphToken()
  let res = await attempt(url, init, token)

  // 401 — token may have expired between cache check and use; refresh once
  if (res.status === 401) {
    debugLog('graphClient', '401 — refreshing token and retrying')
    clearCachedToken()
    token = await getGraphToken()
    res = await attempt(url, init, token)
  }

  // 429 — rate limited; honour Retry-After header, retry once
  if (res.status === 429) {
    const retryAfterSec = Number(res.headers.get('Retry-After') ?? '5')
    debugLog('graphClient', `429 — waiting ${retryAfterSec}s`)
    await new Promise(resolve => setTimeout(resolve, retryAfterSec * 1000))
    res = await attempt(url, init, token)
  }

  if (!res.ok) {
    const body = await res.text().catch(() => '')
    debugError('graphClient', `HTTP ${res.status} ${res.statusText}`, 'URL:', url)
    debugError('graphClient', 'Response body:', body)
    
    // Try to parse error details if JSON
    let errorDetails = body
    try {
      const parsed = JSON.parse(body)
      if (parsed.error) {
        errorDetails = `${parsed.error.code}: ${parsed.error.message}`
      }
    } catch {
      // Not JSON, use raw body
    }
    
    throw new Error(`Graph API error ${res.status}: ${errorDetails}`)
  }

  return res
}

export function createGraphClient(): GraphClient {
  return {
    async get<T>(path: string, params?: Record<string, string>): Promise<T> {
      const url = resolveUrl(path, params)
      debugLog('graphClient', 'GET', url)
      const res = await graphRequest(url, { method: 'GET' })
      return res.json() as Promise<T>
    },

    async post<T>(path: string, body: unknown, contentType = 'application/json'): Promise<T> {
      const url = resolveUrl(path)
      debugLog('graphClient', 'POST', url)
      const res = await graphRequest(url, {
        method: 'POST',
        headers: { 'Content-Type': contentType },
        body: contentType === 'application/json'
          ? JSON.stringify(body)
          : body as BodyInit,
      })
      return res.json() as Promise<T>
    },
  }
}
