import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import { createSettingsService } from '@/settings/settingsService'
import { createGraphClient } from '@/services/graphClient'
import { createMailService } from '@/services/mailService'
import { createOneNoteService } from '@/onenote/oneNoteService'
import './taskpane.css'

Office.onReady(() => {
  // Services that depend on Office APIs must be created inside onReady
  const settingsService = createSettingsService()
  const graphClient = createGraphClient()
  const mailService = createMailService(graphClient)
  const oneNoteService = createOneNoteService(graphClient)

  const root = document.getElementById('root')
  if (!root) throw new Error('Root element not found')

  ReactDOM.createRoot(root).render(
    React.createElement(App, { settingsService, mailService, oneNoteService }),
  )
})
