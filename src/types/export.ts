import type { Page } from '@/onenote/types'

export interface ExportJob {
  conversationId: string
  notebookId: string
  sectionName: string
}

export type ExportStatus = 'idle' | 'fetching' | 'exporting' | 'success' | 'error'

export interface ExportResult {
  pages: Page[]
  sectionWebUrl: string
  sectionClientUrl: string
}
