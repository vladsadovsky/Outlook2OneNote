export function debugLog(tag: string, ...args: unknown[]): void {
  const timestamp = new Date().toISOString().slice(11, 23)
  const message = `[${timestamp}] [${tag}] ${args.map(arg => 
    typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)
  ).join(' ')}`
  
  // Always log to console in dev mode (visible in dev server terminal)
  if (import.meta.env.DEV) {
    console.log(message)
    // Also use console.warn with emoji for better VS Code terminal visibility
    console.warn(`🔍 DEBUG: ${message}`)
  }
}

export function debugError(tag: string, ...args: unknown[]): void {
  const timestamp = new Date().toISOString().slice(11, 23)
  const message = `[${timestamp}] [${tag}] ERROR: ${args.map(arg => 
    typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)
  ).join(' ')}`
  
  // Always log errors to console in dev mode
  if (import.meta.env.DEV) {
    console.error(message)
    // Also use console.warn with emoji for better VS Code terminal visibility
    console.warn(`🚨 ERROR: ${message}`)
  }
}
