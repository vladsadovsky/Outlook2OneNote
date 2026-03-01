import { describe, it, expect, vi } from 'vitest'
import type { GraphClient } from '@/services/graphClient'
import { createOneNoteService } from '@/onenote/oneNoteService'

function makeClient(): GraphClient {
  return {
    get: vi.fn(),
    post: vi.fn(),
  }
}

describe('createOneNoteService', () => {
  describe('listNotebooks', () => {
    it('calls /me/onenote/notebooks and maps the result', async () => {
      const client = makeClient()
      vi.mocked(client.get).mockResolvedValue({
        value: [
          { id: 'nb-1', displayName: 'My Notebook', links: { oneNoteWebUrl: { href: 'https://onenote.com/nb1' } } },
          { id: 'nb-2', displayName: 'Work', links: { oneNoteWebUrl: { href: 'https://onenote.com/nb2' } } },
        ],
      })
      const service = createOneNoteService(client)
      const notebooks = await service.listNotebooks()
      expect(notebooks).toHaveLength(2)
      expect(notebooks[0]).toEqual({ id: 'nb-1', displayName: 'My Notebook', webUrl: 'https://onenote.com/nb1' })
      expect(client.get).toHaveBeenCalledWith('/me/onenote/notebooks', expect.objectContaining({ '$select': expect.any(String) }))
    })
  })

  describe('createSection', () => {
    it('posts to /me/onenote/notebooks/{id}/sections', async () => {
      const client = makeClient()
      vi.mocked(client.post).mockResolvedValue({ id: 'sec-1', displayName: 'My Section' })
      const service = createOneNoteService(client)
      const section = await service.createSection('nb-1', 'My Section')
      expect(section).toEqual({ id: 'sec-1', displayName: 'My Section' })
      expect(client.post).toHaveBeenCalledWith(
        '/me/onenote/notebooks/nb-1/sections',
        { displayName: 'My Section' },
      )
    })
  })

  describe('createPage', () => {
    it('posts HTML to /me/onenote/sections/{id}/pages with text/html content type', async () => {
      const client = makeClient()
      vi.mocked(client.post).mockResolvedValue({
        id: 'page-1',
        links: {
          oneNoteWebUrl: { href: 'https://onenote.com/page1' },
          oneNoteClientUrl: { href: 'onenote:///page1' },
        },
      })
      const service = createOneNoteService(client)
      const pageHtml = '<html><body><p>content</p></body></html>'
      const page = await service.createPage('sec-1', pageHtml, 'Email title')
      expect(page).toEqual({
        id: 'page-1',
        webUrl: 'https://onenote.com/page1',
        oneNoteClientUrl: 'onenote:///page1',
      })
      expect(client.post).toHaveBeenCalledWith(
        '/me/onenote/sections/sec-1/pages',
        pageHtml,
        'text/html',
      )
    })
  })
})
