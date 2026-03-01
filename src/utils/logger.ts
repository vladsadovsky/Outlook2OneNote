export function debugLog(tag: string, ...args: unknown[]): void {
  if (import.meta.env.DEV) {
    console.log(`[${tag}]`, ...args)
  }
}

export function debugError(tag: string, ...args: unknown[]): void {
  if (import.meta.env.DEV) {
    console.error(`[${tag}]`, ...args)
  }
}
