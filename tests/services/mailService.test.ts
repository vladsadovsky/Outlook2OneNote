import { describe, it, expect, vi, beforeEach } from 'vitest'
import type { GraphClient } from '@/services/graphClient'
import { createMailService, ThreadTooLargeError } from '@/services/mailService'

interface GraphEmailAddress { name: string; address: string }
interface GraphAttachment { name: string; size: number; contentType: string; isInline: boolean }
interface GraphMsg {
  id: string
  subject: string
  from: { emailAddress: GraphEmailAddress }
  toRecipients: { emailAddress: GraphEmailAddress }[]
  ccRecipients: { emailAddress: GraphEmailAddress }[]
  receivedDateTime: string
  body: { content: string; contentType: string }
  hasAttachments: boolean
  attachments: GraphAttachment[]
}

function makeGraphMsg(id: string, subject = 'Test'): GraphMsg {
  return {
    id,
    subject,
    from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
    toRecipients: [{ emailAddress: { name: 'Bob', address: 'bob@example.com' } }],
    ccRecipients: [],
    receivedDateTime: '2024-01-01T10:00:00Z',
    body: { content: '<p>body</p>', contentType: 'html' },
    hasAttachments: false,
    attachments: [],
  }
}

function makeClient(pages: { value: GraphMsg[]; nextLink?: string }[]): GraphClient {
  let callCount = 0
  const mockGet = vi.fn() as GraphClient['get']
  vi.mocked(mockGet).mockImplementation(async <T>() => {
    const page = pages[callCount++]
    return { value: page.value, '@odata.nextLink': page.nextLink } as T
  })
  return { get: mockGet, post: vi.fn() }
}

describe('createMailService', () => {
  let service: ReturnType<typeof createMailService>

  beforeEach(() => {
    service = createMailService(makeClient([{ value: [makeGraphMsg('1')] }]))
  })

  it('returns mapped EmailMessage objects', async () => {
    const messages = await service.getConversationMessages('conv-1')
    expect(messages).toHaveLength(1)
    expect(messages[0].id).toBe('1')
    expect(messages[0].subject).toBe('Test')
    expect(messages[0].from.address).toBe('alice@example.com')
  })

  it('follows @odata.nextLink for paged results', async () => {
    const client = makeClient([
      { value: [makeGraphMsg('1'), makeGraphMsg('2')], nextLink: 'https://graph.microsoft.com/v1.0/next' },
      { value: [makeGraphMsg('3')] },
    ])
    service = createMailService(client)
    const messages = await service.getConversationMessages('conv-1')
    expect(messages).toHaveLength(3)
    expect(client.get).toHaveBeenCalledTimes(2)
  })

  it('passes nextLink as path for subsequent pages with no extra params', async () => {
    const nextLink = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc'
    const client = makeClient([
      { value: [makeGraphMsg('1')], nextLink },
      { value: [makeGraphMsg('2')] },
    ])
    service = createMailService(client)
    await service.getConversationMessages('conv-1')
    expect(vi.mocked(client.get).mock.calls[1][0]).toBe(nextLink)
    expect(vi.mocked(client.get).mock.calls[1][1]).toBeUndefined()
  })

  it('throws ThreadTooLargeError when message count exceeds 50', async () => {
    const msgs = Array.from({ length: 51 }, (_, i) => makeGraphMsg(`msg-${i}`))
    const client = makeClient([{ value: msgs }])
    service = createMailService(client)
    await expect(service.getConversationMessages('conv-1')).rejects.toThrow(ThreadTooLargeError)
  })

  it('ThreadTooLargeError carries the count', async () => {
    const msgs = Array.from({ length: 51 }, (_, i) => makeGraphMsg(`msg-${i}`))
    const client = makeClient([{ value: msgs }])
    service = createMailService(client)
    try {
      await service.getConversationMessages('conv-1')
    } catch (e) {
      expect(e).toBeInstanceOf(ThreadTooLargeError)
      expect((e as ThreadTooLargeError).count).toBe(51)
    }
  })

  it('filters out inline attachments', async () => {
    const msg = makeGraphMsg('1')
    msg.hasAttachments = true
    const attachments = [
      { name: 'inline-img.png', size: 500, contentType: 'image/png', isInline: true },
      { name: 'report.pdf', size: 1024, contentType: 'application/pdf', isInline: false },
    ]
    const mockGet = vi.fn() as GraphClient['get']
    vi.mocked(mockGet).mockImplementation(async <T>(path: string) => {
      if (path === '/me/messages') {
        return { value: [msg] } as T
      }
      if (path === `/me/messages/${msg.id}/attachments`) {
        return { value: attachments } as T
      }
      throw new Error(`Unexpected path: ${path}`)
    })
    const client: GraphClient = { get: mockGet, post: vi.fn() }
    service = createMailService(client)
    const messages = await service.getConversationMessages('conv-1')
    expect(messages[0].attachments).toHaveLength(1)
    expect(messages[0].attachments[0].name).toBe('report.pdf')
  })

  it('maps ccRecipients correctly', async () => {
    const msg = makeGraphMsg('1')
    msg.ccRecipients = [{ emailAddress: { name: 'Carol', address: 'carol@example.com' } }]
    const client = makeClient([{ value: [msg] }])
    service = createMailService(client)
    const messages = await service.getConversationMessages('conv-1')
    expect(messages[0].ccRecipients).toHaveLength(1)
    expect(messages[0].ccRecipients[0].address).toBe('carol@example.com')
  })
})
