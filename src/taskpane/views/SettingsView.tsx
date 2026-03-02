import React, { useEffect, useRef, useState } from 'react'
import type { SettingsSchema } from '@/types/settings'
import type { SettingsService } from '@/settings/settingsService'
import type { OneNoteService } from '@/onenote/oneNoteService'
import type { Notebook } from '@/onenote/types'
import { debugLog, debugError } from '@/utils/logger'
import NotebookPicker from '../components/NotebookPicker'

interface SettingsViewProps {
  settingsService: SettingsService
  oneNoteService: OneNoteService
  onClose: () => void
  onSaved: (settings: SettingsSchema) => void
}

export default function SettingsView({
  settingsService,
  oneNoteService,
  onClose,
  onSaved,
}: SettingsViewProps): React.ReactElement {
  debugLog('SettingsView', 'Component rendering/mounting')
  
  const current = settingsService.get()

  // Draft state — not persisted until Save
  const [draft, setDraft] = useState<SettingsSchema>(current)
  const [notebooks, setNotebooks] = useState<Notebook[]>([])
  const [notebooksLoading, setNotebooksLoading] = useState(false)
  const [notebooksError, setNotebooksError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)

  const overlayRef = useRef<HTMLDivElement>(null)

  // Load notebooks on mount
  useEffect(() => {
    debugLog('SettingsView', 'useEffect triggered - calling loadNotebooks')
    void loadNotebooks()
  }, [])

  // Keyboard handling: Escape = cancel, Ctrl+Enter = save
  useEffect(() => {
    function onKeyDown(e: KeyboardEvent): void {
      if (e.key === 'Escape') onClose()
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) void handleSave()
    }
    document.addEventListener('keydown', onKeyDown)
    return () => document.removeEventListener('keydown', onKeyDown)
  })

  // Click-outside to dismiss (cancel)
  function handleOverlayClick(e: React.MouseEvent<HTMLDivElement>): void {
    if (e.target === overlayRef.current) onClose()
  }

  async function loadNotebooks(): Promise<void> {
    setNotebooksLoading(true)
    setNotebooksError(null)
    
    debugLog('SettingsView', 'Starting notebook loading...')
    
    try {
      console.log('[SettingsView] Loading notebooks...')
      const list = await oneNoteService.listNotebooks()
      console.log('[SettingsView] Notebooks loaded:', list.length)
      
      debugLog('SettingsView', `Successfully loaded ${list.length} notebooks!`)
      setNotebooks(list)
    } catch (error) {
      console.error('[SettingsView] Failed to load notebooks:', error)
      
      const errorMsg = error instanceof Error ? error.message : 'Unknown error'
      debugError('SettingsView', `Notebook loading failed: ${errorMsg}`)
      
      // Check for SharePoint license issue
      if (errorMsg.includes('SharePoint license') || errorMsg.includes('30121')) {
        setNotebooksError(`Microsoft Graph API doesn't support OneNote access for personal accounts. 
This is a Microsoft limitation. Consider using a work/school account with Microsoft 365, 
or export emails manually to OneNote for now.`)
      } else if (errorMsg.includes('404')) {
        setNotebooksError('OneNote service not found. Please ensure you have access to Microsoft OneNote.')
      } else if (errorMsg.includes('All OneNote API approaches failed')) {
        setNotebooksError(`Personal Microsoft accounts have limited OneNote API access. 
Try using a work/school account, or access OneNote directly at onenote.com.`)
      } else {
        setNotebooksError('Could not load notebooks. Check your connection and try again.')
      }
    } finally {
      setNotebooksLoading(false)
    }
  }

  function handleNotebookSelect(nb: Notebook): void {
    setDraft(d => ({ ...d, notebookId: nb.id, notebookDisplayName: nb.displayName }))
  }

  async function handleSave(): Promise<void> {
    setSaving(true)
    setSaveError(null)
    try {
      settingsService.set(draft)
      await settingsService.save()
      onSaved({ ...draft })
      onClose()
    } catch {
      setSaveError('Failed to save settings. Please try again.')
    } finally {
      setSaving(false)
    }
  }

  return (
    <div className="settings-overlay" ref={overlayRef} onClick={handleOverlayClick}>
      <div className="settings-dialog" role="dialog" aria-label="Settings" aria-modal="true">
        <h2 className="settings-dialog__title">Settings</h2>

        {/* Notebook picker */}
        {notebooksLoading ? (
          <p className="settings-dialog__loading">Loading notebooks…</p>
        ) : notebooksError ? (
          <p className="settings-dialog__error">{notebooksError}</p>
        ) : (
          <NotebookPicker
            notebooks={notebooks}
            selectedId={draft.notebookId}
            onSelect={handleNotebookSelect}
            onRefresh={() => void loadNotebooks()}
          />
        )}

        {/* Section name format */}
        <div className="settings-field">
          <label className="settings-field__label" htmlFor="section-format">
            Section name format
          </label>
          <input
            id="section-format"
            className="settings-field__input"
            type="text"
            value={draft.sectionNameFormat}
            onChange={e => setDraft(d => ({ ...d, sectionNameFormat: e.target.value }))}
          />
          <span className="settings-field__hint">Tokens: {'{subject}'} {'{date}'}</span>
        </div>

        {/* Sort order */}
        <div className="settings-field">
          <span className="settings-field__label">Email sort order</span>
          <label className="settings-field__radio">
            <input
              type="radio"
              name="sortOrder"
              value="asc"
              checked={draft.sortOrder === 'asc'}
              onChange={() => setDraft(d => ({ ...d, sortOrder: 'asc' }))}
            />
            Oldest first
          </label>
          <label className="settings-field__radio">
            <input
              type="radio"
              name="sortOrder"
              value="desc"
              checked={draft.sortOrder === 'desc'}
              onChange={() => setDraft(d => ({ ...d, sortOrder: 'desc' }))}
            />
            Newest first
          </label>
        </div>

        {/* Include attachments */}
        <div className="settings-field">
          <label className="settings-field__checkbox">
            <input
              type="checkbox"
              checked={draft.includeAttachments}
              onChange={e => setDraft(d => ({ ...d, includeAttachments: e.target.checked }))}
            />
            Include attachment list in exported pages
          </label>
        </div>

        {/* Preferred link */}
        <div className="settings-field">
          <span className="settings-field__label">Open-in link</span>
          {(['web', 'desktop', 'both'] as const).map(v => (
            <label key={v} className="settings-field__radio">
              <input
                type="radio"
                name="preferredLink"
                value={v}
                checked={draft.preferredLink === v}
                onChange={() => setDraft(d => ({ ...d, preferredLink: v }))}
              />
              {v === 'web' ? 'Web only' : v === 'desktop' ? 'Desktop app only' : 'Both'}
            </label>
          ))}
        </div>

        {/* Debug panel */}
        <div className="settings-field">
          <label className="settings-field__checkbox">
            <input
              type="checkbox"
              checked={draft.showDebugPanel}
              onChange={e => setDraft(d => ({ ...d, showDebugPanel: e.target.checked }))}
            />
            Show debug console
          </label>
        </div>

        {saveError && <p className="settings-dialog__error">{saveError}</p>}

        {/* Actions */}
        <div className="settings-dialog__actions">
          <button type="button" className="btn btn--secondary" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn btn--primary"
            onClick={() => void handleSave()}
            disabled={saving}
          >
            {saving ? 'Saving…' : 'Save'}
          </button>
        </div>
        <p className="settings-dialog__hint">Ctrl+Enter to save · Esc to cancel</p>
      </div>
    </div>
  )
}
