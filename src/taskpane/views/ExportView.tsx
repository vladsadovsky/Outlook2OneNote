import React, { useState } from 'react'
import type { MailService } from '@/services/mailService'
import type { OneNoteService } from '@/onenote/oneNoteService'
import type { SettingsService } from '@/settings/settingsService'
import type { ExportStatus } from '@/types/export'
import type { Page } from '@/onenote/types'
import { ThreadTooLargeError } from '@/services/mailService'
import { exportThread } from '../exportOrchestrator'
import ProgressBar from '../components/ProgressBar'
import ErrorBanner from '../components/ErrorBanner'

interface ExportViewProps {
  mailService: MailService
  oneNoteService: OneNoteService
  settingsService: SettingsService
  onOpenSettings: () => void
}

export default function ExportView({
  mailService,
  oneNoteService,
  settingsService,
  onOpenSettings,
}: ExportViewProps): React.ReactElement {
  const [status, setStatus] = useState<ExportStatus>('idle')
  const [progressMsg, setProgressMsg] = useState('')
  const [errorMsg, setErrorMsg] = useState('')
  const [exportedPages, setExportedPages] = useState<Page[]>([])

  const settings = settingsService.get()

  async function handleExport(): Promise<void> {
    if (!settings.notebookId) {
      onOpenSettings()
      return
    }

    setStatus('fetching')
    setErrorMsg('')
    setExportedPages([])

    try {
      const item = Office.context.mailbox.item
      if (!item) throw new Error('No email selected.')

      const conversationId = item.conversationId
      if (!conversationId) throw new Error('Conversation ID not available.')

      const result = await exportThread({
        conversationId,
        settings: settingsService.get(),
        mailService,
        oneNoteService,
        onProgress: p => {
          setProgressMsg(p.message)
          setStatus('exporting')
        },
      })

      setExportedPages(result.pages)
      setStatus('success')
    } catch (e) {
      if (e instanceof ThreadTooLargeError) {
        setErrorMsg(
          `This thread has ${e.count}+ messages, which exceeds the 50-message limit. ` +
          'Please archive or delete older messages in Outlook to reduce the thread size, then try again.',
        )
      } else {
        setErrorMsg(e instanceof Error ? e.message : 'An unexpected error occurred.')
      }
      setStatus('error')
    }
  }

  function handleRetry(): void {
    setStatus('idle')
    setErrorMsg('')
  }

  const noNotebook = !settings.notebookId

  return (
    <div className="export-view">
      {noNotebook && (
        <div className="export-view__no-notebook">
          <p>No notebook selected.</p>
          <button type="button" className="btn btn--secondary" onClick={onOpenSettings}>
            Open Settings
          </button>
        </div>
      )}

      {status === 'idle' && !noNotebook && (
        <div className="export-view__idle">
          <p className="export-view__notebook-name">
            Notebook: <strong>{settings.notebookDisplayName ?? settings.notebookId}</strong>
          </p>
          <button
            type="button"
            className="btn btn--primary export-view__export-btn"
            onClick={() => void handleExport()}
          >
            Export to OneNote
          </button>
        </div>
      )}

      {(status === 'fetching' || status === 'exporting') && (
        <ProgressBar message={progressMsg || 'Starting…'} />
      )}

      {status === 'error' && (
        <ErrorBanner message={errorMsg} onRetry={handleRetry} />
      )}

      {status === 'success' && (
        <div className="export-view__success">
          <p className="export-view__success-msg">
            ✓ Exported {exportedPages.length} message{exportedPages.length !== 1 ? 's' : ''} to OneNote.
          </p>
          <div className="export-view__links">
            {exportedPages.length > 0 && (settings.preferredLink === 'web' || settings.preferredLink === 'both') && (
              <a
                href={exportedPages[0].webUrl}
                target="_blank"
                rel="noreferrer"
                className="export-view__link"
              >
                Open in OneNote Web
              </a>
            )}
            {exportedPages.length > 0 && (settings.preferredLink === 'desktop' || settings.preferredLink === 'both') && (
              <a
                href={exportedPages[0].oneNoteClientUrl}
                target="_blank"
                rel="noreferrer"
                className="export-view__link"
              >
                Open in OneNote Desktop
              </a>
            )}
          </div>
          <button
            type="button"
            className="btn btn--secondary"
            onClick={() => setStatus('idle')}
          >
            Export another
          </button>
        </div>
      )}
    </div>
  )
}
