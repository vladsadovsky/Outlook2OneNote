import { describe, it, expect, vi, afterEach } from 'vitest'
import { withRetry } from '@/utils/retry'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
})

describe('withRetry', () => {
  it('returns the result when the function succeeds on first attempt', async () => {
    const fn = vi.fn<() => Promise<string>>().mockResolvedValue('ok')
    const result = await withRetry(fn)
    expect(result).toBe('ok')
    expect(fn).toHaveBeenCalledTimes(1)
  })

  it('retries after failure and returns result on second attempt', async () => {
    vi.useFakeTimers()
    const fn = vi.fn<() => Promise<string>>()
      .mockRejectedValueOnce(new Error('fail'))
      .mockResolvedValue('ok')

    const promise = withRetry(fn, 3, 100)
    await vi.advanceTimersByTimeAsync(100)
    const result = await promise
    expect(result).toBe('ok')
    expect(fn).toHaveBeenCalledTimes(2)
  })

  it('throws after maxRetries attempts', async () => {
    vi.useFakeTimers()
    const err = new Error('persistent failure')
    const fn = vi.fn<() => Promise<string>>().mockRejectedValue(err)

    const promise = withRetry(fn, 3, 10)
    await vi.advanceTimersByTimeAsync(10000)
    await expect(promise).rejects.toThrow('persistent failure')
    expect(fn).toHaveBeenCalledTimes(3)
  })

  it('uses exponential backoff: delay doubles on each retry', async () => {
    vi.useFakeTimers()
    const delays: number[] = []

    // Spy before fake timers replace setTimeout
    vi.spyOn(globalThis, 'setTimeout').mockImplementation((callback, delay) => {
      if (typeof delay === 'number') delays.push(delay)
      // Run callback immediately so the promise chain progresses
      if (typeof callback === 'function') (callback as () => void)()
      return 0 as unknown as ReturnType<typeof setTimeout>
    })

    const fn = vi.fn<() => Promise<string>>()
      .mockRejectedValueOnce(new Error('fail 1'))
      .mockRejectedValueOnce(new Error('fail 2'))
      .mockResolvedValue('ok')

    await withRetry(fn, 3, 500)

    expect(delays[0]).toBe(500)   // attempt 1: base * 2^0
    expect(delays[1]).toBe(1000)  // attempt 2: base * 2^1
  })
})
