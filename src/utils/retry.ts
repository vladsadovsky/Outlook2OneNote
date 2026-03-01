import { debugLog } from './logger'

export async function withRetry<T>(
  fn: () => Promise<T>,
  maxRetries = 3,
  baseDelay = 500,
): Promise<T> {
  let attempt = 0
  while (true) {
    try {
      return await fn()
    } catch (error) {
      attempt++
      if (attempt >= maxRetries) throw error
      const delay = baseDelay * Math.pow(2, attempt - 1)
      debugLog('retry', `attempt ${attempt} in ${delay}ms...`)
      await new Promise(resolve => setTimeout(resolve, delay))
    }
  }
}
