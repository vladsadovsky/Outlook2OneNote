import React from 'react'

interface AppHeaderProps {
  onSettingsClick: () => void
}

export default function AppHeader({ onSettingsClick }: AppHeaderProps): React.ReactElement {
  return (
    <header className="app-header">
      <span className="app-header__title">Export to OneNote</span>
      <button
        type="button"
        className="app-header__settings-btn"
        aria-label="Open settings"
        onClick={onSettingsClick}
      >
        ⚙
      </button>
    </header>
  )
}
