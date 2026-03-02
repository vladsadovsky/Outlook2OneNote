import type { GraphClient } from '@/services/graphClient'
import type { Notebook, Page, Section } from './types'
import { debugLog, debugError } from '@/utils/logger'

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
      debugLog('oneNoteService', 'listNotebooks called - starting API call')
      
      try {
        // First, let's debug what account type we're dealing with
        debugLog('oneNoteService', 'Testing Graph API access...')
        
        // Try to get basic account info first
        const userInfo = await client.get<any>('/me')
        debugLog('oneNoteService', `User account: ${userInfo.userPrincipalName} (tenant: ${userInfo.tenantId || 'personal'})`)
        
        // For personal accounts, try multiple approaches
        const approaches = [
          // Approach 1: Standard v1.0 (we expect this to fail)
          () => client.get<GraphNotebookList>('/me/onenote/notebooks', {
            '$select': 'id,displayName,links',
            '$orderby': 'displayName asc',
          }),
          
          // Approach 2: Beta API without orderby
          () => client.get<GraphNotebookList>('/beta/me/onenote/notebooks', {
            '$select': 'id,displayName,links',
          }),
          
          // Approach 3: Simple beta call without any query params
          () => client.get<GraphNotebookList>('/beta/me/onenote/notebooks'),
          
          // Approach 4: Try v1.0 without orderby
          () => client.get<GraphNotebookList>('/me/onenote/notebooks', {
            '$select': 'id,displayName,links',
          }),
          
          // Approach 5: Minimal v1.0 call
          () => client.get<GraphNotebookList>('/me/onenote/notebooks'),
        ]
        
        for (let i = 0; i < approaches.length; i++) {
          try {
            debugLog('oneNoteService', `Trying approach ${i + 1}/${approaches.length}...`)
            const res = await approaches[i]()
            debugLog('oneNoteService', `Approach ${i + 1} success - got ${res.value.length} notebooks`)
            return res.value.map(n => ({
              id: n.id,
              displayName: n.displayName,
              webUrl: n.links?.oneNoteWebUrl?.href || '#',
            }))
          } catch (error) {
            debugError('oneNoteService', `Approach ${i + 1} failed:`, error)
            if (i < approaches.length - 1) {
              debugLog('oneNoteService', `Trying next approach...`)
            }
          }
        }
        
        // If all approaches fail, throw a comprehensive error
        throw new Error('All OneNote API approaches failed. Personal Microsoft accounts may have limited Graph API access for OneNote.')
        
      } catch (error) {
        debugError('oneNoteService', 'All listNotebooks approaches failed:', error)
        throw error
      }
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
