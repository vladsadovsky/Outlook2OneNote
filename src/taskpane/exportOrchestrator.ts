import type { MailService } from '@/services/mailService'
import type { OneNoteService } from '@/onenote/oneNoteService'
import type { SettingsSchema } from '@/types/settings'
import type { ExportResult } from '@/types/export'
import { buildPageHtml } from '@/onenote/pageBuilder'
import { formatDateForSection, formatDateTimeForPageTitle } from '@/utils/dateFormatter'
import { debugLog } from '@/utils/logger'

export interface ExportProgress {
  message: string
}

export interface ExportOptions {
  conversationId: string
  settings: SettingsSchema
  mailService: MailService
  oneNoteService: OneNoteService
  onProgress: (progress: ExportProgress) => void
}

// Replaces {subject} and {date} tokens in the section name format string.
function buildSectionName(format: string, subject: string, isoDate: string): string {
  const date = formatDateForSection(isoDate)
  const raw = format
    .replace('{subject}', subject)
    .replace('{date}', date)
  const sanitized = raw
    // OneNote/Graph name constraints: remove invalid characters
    .replace(/[\?*\/\\:<>|&#'"%~]/g, ' ')
    // Collapse repeated whitespace
    .replace(/\s+/g, ' ')
    .trim()
  const finalName = sanitized.length > 0 ? sanitized : `Email Thread ${date}`

  // Truncate to 50 chars — OneNote section name limit
  return finalName.slice(0, 50)
}

export async function exportThread(options: ExportOptions): Promise<ExportResult> {
  const { conversationId, settings, mailService, oneNoteService, onProgress } = options

  // ── 1. Fetch messages ────────────────────────────────────────────────────────
  onProgress({ message: 'Fetching messages…' })
  let messages = await mailService.getConversationMessages(conversationId)
  debugLog('orchestrator', `Fetched ${messages.length} message(s)`)

  // ── 2. Apply sort order ──────────────────────────────────────────────────────
  if (settings.sortOrder === 'desc') {
    messages = [...messages].reverse()
  }

  // ── 3. Build section name from first message ─────────────────────────────────
  const firstMsg = messages[0]
  const sectionName = buildSectionName(
    settings.sectionNameFormat,
    firstMsg?.subject ?? 'Email Thread',
    firstMsg?.receivedDateTime ?? new Date().toISOString(),
  )

  // ── 4. Create OneNote section ────────────────────────────────────────────────
  onProgress({ message: 'Creating OneNote section…' })
  if (!settings.notebookId) {
    throw new Error('No notebook selected. Open Settings to choose a notebook.')
  }
  const section = await oneNoteService.createSection(settings.notebookId, sectionName)
  debugLog('orchestrator', 'Section created', section.id)

  // ── 5. Export each message ───────────────────────────────────────────────────
  const pages: ExportResult['pages'] = []
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i]
    onProgress({ message: `Exporting message ${i + 1} of ${messages.length}…` })

    const title = `${formatDateTimeForPageTitle(msg.receivedDateTime)} — ${msg.subject}`
    const pageHtml = buildPageHtml(msg, settings.includeAttachments)
    const page = await oneNoteService.createPage(section.id, pageHtml, title)
    pages.push(page)
  }

  debugLog('orchestrator', 'Export complete', pages.length, 'pages')

  return {
    pages,
    sectionWebUrl: '',       // section URL not returned by createSection; link via pages
    sectionClientUrl: '',
  }
}
