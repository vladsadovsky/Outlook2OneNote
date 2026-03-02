import React from 'react'
import { debugLog } from '@/utils/logger'

interface AppHeaderProps {
  onSettingsClick: () => void
}

export default function AppHeader({ onSettingsClick }: AppHeaderProps): React.ReactElement {
  const handleSettingsClick = () => {
    debugLog('AppHeader', 'Settings gear icon clicked')
    onSettingsClick()
  }

  return (
    <header className="app-header">
      <span className="app-header__title">Export Thread to OneNote</span>
      <button
        type="button"
        className="app-header__settings-btn"
        aria-label="Open settings"
        onClick={handleSettingsClick}
      >
        ⚙
      </button>
    </header>
  )
}
