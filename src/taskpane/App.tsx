import React, { useState } from 'react'
import type { SettingsService } from '@/settings/settingsService'
import type { MailService } from '@/services/mailService'
import type { OneNoteService } from '@/onenote/oneNoteService'
import { debugLog } from '@/utils/logger'
import type { SettingsSchema } from '@/types/settings'
import AppHeader from './components/AppHeader'
import ExportView from './views/ExportView'
import SettingsView from './views/SettingsView'
import DebugPanel from './components/DebugPanel'

interface AppProps {
  settingsService: SettingsService
  mailService: MailService
  oneNoteService: OneNoteService
}

export default function App({
  settingsService,
  mailService,
  oneNoteService,
}: AppProps): React.ReactElement {
  const [showSettings, setShowSettings] = useState(false)
  const [settings, setSettings] = useState<SettingsSchema>(() => settingsService.get())

  const handleShowSettings = () => {
    debugLog('App', 'Opening settings view')
    setShowSettings(true)
  }

  const handleCloseSettings = () => {
    debugLog('App', 'Closing settings view')
    setShowSettings(false)
  }

  const handleSettingsSaved = (next: SettingsSchema) => {
    setSettings(next)
  }

  return (
    <div className="app">
      <AppHeader onSettingsClick={handleShowSettings} />
      {showSettings ? (
        <SettingsView
          settingsService={settingsService}
          oneNoteService={oneNoteService}
          onClose={handleCloseSettings}
          onSaved={handleSettingsSaved}
        />
      ) : (
        <ExportView
          mailService={mailService}
          oneNoteService={oneNoteService}
          settingsService={settingsService}
          onOpenSettings={handleShowSettings}
        />
      )}
      {/* Debug panel - only visible in development */}
      {import.meta.env.DEV && settings.showDebugPanel && <DebugPanel />}
    </div>
  )
}
