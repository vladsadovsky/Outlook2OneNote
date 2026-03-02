import React, { useEffect, useState } from 'react'

interface DebugEntry {
  timestamp: string
  message: string
  level: 'log' | 'error' | 'warn'
}

export default function DebugPanel(): React.ReactElement {
  const [entries, setEntries] = useState<DebugEntry[]>([])
  const [isVisible, setIsVisible] = useState(true)

  useEffect(() => {
    // Override console methods to capture logs
    const originalLog = console.log
    const originalError = console.error
    const originalWarn = console.warn

    const addEntry = (level: 'log' | 'error' | 'warn', ...args: unknown[]) => {
      const timestamp = new Date().toISOString().slice(11, 23)
      const message = args.map(arg => 
        typeof arg === 'object' ? JSON.stringify(arg, null, 2) : String(arg)
      ).join(' ')
      
      setEntries(prev => [...prev.slice(-49), { timestamp, message, level }]) // Keep last 50 entries
    }

    console.log = (...args) => {
      originalLog(...args)
      addEntry('log', ...args)
    }

    console.error = (...args) => {
      originalError(...args)
      addEntry('error', ...args)
    }

    console.warn = (...args) => {
      originalWarn(...args)
      addEntry('warn', ...args)
    }

    // Cleanup on unmount
    return () => {
      console.log = originalLog
      console.error = originalError
      console.warn = originalWarn
    }
  }, [])

  if (!isVisible) {
    return (
      <button 
        onClick={() => setIsVisible(true)}
        style={{ 
          position: 'fixed', 
          bottom: '10px', 
          right: '10px', 
          zIndex: 1000,
          fontSize: '12px',
          padding: '4px 8px'
        }}
      >
        Show Debug
      </button>
    )
  }

  return (
    <div style={{
      position: 'fixed',
      bottom: '10px',
      right: '10px',
      width: '300px',
      height: '200px',
      backgroundColor: '#f5f5f5',
      border: '1px solid #ccc',
      borderRadius: '4px',
      zIndex: 1000,
      fontSize: '10px',
      fontFamily: 'monospace'
    }}>
      <div style={{ 
        backgroundColor: '#ddd', 
        padding: '4px 8px', 
        display: 'flex', 
        justifyContent: 'space-between',
        alignItems: 'center'
      }}>
        <span>Debug Console</span>
        <div>
          <button 
            onClick={() => setEntries([])}
            style={{ fontSize: '10px', marginRight: '4px', padding: '2px 4px' }}
          >
            Clear
          </button>
          <button 
            onClick={() => setIsVisible(false)}
            style={{ fontSize: '10px', padding: '2px 4px' }}
          >
            ×
          </button>
        </div>
      </div>
      <div style={{
        height: '170px',
        overflow: 'auto',
        padding: '4px',
        backgroundColor: '#fff'
      }}>
        {entries.map((entry, i) => (
          <div key={i} style={{
            marginBottom: '2px',
            color: entry.level === 'error' ? '#d73a49' : 
                   entry.level === 'warn' ? '#e36209' : '#24292e'
          }}>
            <span style={{ color: '#6a737d' }}>[{entry.timestamp}]</span> {entry.message}
          </div>
        ))}
      </div>
    </div>
  )
}