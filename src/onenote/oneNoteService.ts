import type { GraphClient } from '@/services/graphClient'
import type { Notebook, Page, Section } from './types'
import { debugLog } from '@/utils/logger'

// GraphClient is injected — this module has zero runtime imports from auth/ or taskpane/.

export interface OneNoteService {
  listNotebooks(): Promise<Notebook[]>
  createSection(notebookId: string, sectionName: string): Promise<Section>
  // pageHtml is a full OneNote-compatible HTML document (built by pageBuilder)
  createPage(sectionId: string, pageHtml: string, title: string): Promise<Page>
}

// Graph API response shapes
interface GraphNotebookList {
  value: GraphNotebook[]
}
interface GraphNotebook {
  id: string
  displayName: string
  links: { oneNoteWebUrl: { href: string } }
}
interface GraphSection {
  id: string
  displayName: string
}
interface GraphPage {
  id: string
  links: {
    oneNoteWebUrl: { href: string }
    oneNoteClientUrl: { href: string }
  }
}

export function createOneNoteService(client: GraphClient): OneNoteService {
  return {
    async listNotebooks(): Promise<Notebook[]> {
      debugLog('oneNoteService', 'listNotebooks')
      const res = await client.get<GraphNotebookList>('/me/onenote/notebooks', {
        '$select': 'id,displayName,links',
        '$orderby': 'displayName asc',
      })
      return res.value.map(n => ({
        id: n.id,
        displayName: n.displayName,
        webUrl: n.links.oneNoteWebUrl.href,
      }))
    },

    async createSection(notebookId: string, sectionName: string): Promise<Section> {
      debugLog('oneNoteService', 'createSection', sectionName)
      const res = await client.post<GraphSection>(
        `/me/onenote/notebooks/${notebookId}/sections`,
        { displayName: sectionName },
      )
      return { id: res.id, displayName: res.displayName }
    },

    async createPage(sectionId: string, pageHtml: string, title: string): Promise<Page> {
      debugLog('oneNoteService', 'createPage', title)
      const res = await client.post<GraphPage>(
        `/me/onenote/sections/${sectionId}/pages`,
        pageHtml,
        'text/html',
      )
      return {
        id: res.id,
        webUrl: res.links.oneNoteWebUrl.href,
        oneNoteClientUrl: res.links.oneNoteClientUrl.href,
      }
    },
  }
}
