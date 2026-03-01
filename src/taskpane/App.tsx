import React, { useState } from 'react'
import type { SettingsService } from '@/settings/settingsService'
import type { MailService } from '@/services/mailService'
import type { OneNoteService } from '@/onenote/oneNoteService'
import AppHeader from './components/AppHeader'
import ExportView from './views/ExportView'
import SettingsView from './views/SettingsView'

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

  return (
    <div className="app">
      <AppHeader onSettingsClick={() => setShowSettings(true)} />
      {showSettings ? (
        <SettingsView
          settingsService={settingsService}
          oneNoteService={oneNoteService}
          onClose={() => setShowSettings(false)}
        />
      ) : (
        <ExportView
          mailService={mailService}
          oneNoteService={oneNoteService}
          settingsService={settingsService}
          onOpenSettings={() => setShowSettings(true)}
        />
      )}
    </div>
  )
}
