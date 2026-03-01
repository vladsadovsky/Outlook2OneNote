import { describe, it, expect } from 'vitest'
import { buildPageHtml } from '@/onenote/pageBuilder'
import type { EmailMessage } from '@/onenote/types'

function makeMessage(overrides: Partial<EmailMessage> = {}): EmailMessage {
  return {
    id: 'msg-1',
    subject: 'Hello World',
    from: { name: 'Alice', address: 'alice@example.com' },
    toRecipients: [{ name: 'Bob', address: 'bob@example.com' }],
    ccRecipients: [],
    receivedDateTime: '2024-06-15T10:30:00.000Z',
    bodyHtml: '<p>Test body</p>',
    attachments: [],
    ...overrides,
  }
}

describe('buildPageHtml', () => {
  it('returns a full HTML document', () => {
    const html = buildPageHtml(makeMessage(), false)
    expect(html).toContain('<!DOCTYPE html>')
    expect(html).toContain('<html>')
    expect(html).toContain('</html>')
    expect(html).toContain('<body>')
  })

  it('includes the email subject as the page title', () => {
    const html = buildPageHtml(makeMessage({ subject: 'My Subject' }), false)
    expect(html).toContain('<title>My Subject</title>')
  })

  it('includes the receivedDateTime as the created meta tag', () => {
    const html = buildPageHtml(makeMessage({ receivedDateTime: '2024-06-15T10:30:00.000Z' }), false)
    expect(html).toContain('name="created"')
    expect(html).toContain('2024-06-15T10:30:00.000Z')
  })

  it('includes from address in metadata table', () => {
    const html = buildPageHtml(makeMessage(), false)
    expect(html).toContain('Alice')
    expect(html).toContain('alice@example.com')
  })

  it('includes to recipients in metadata table', () => {
    const html = buildPageHtml(makeMessage(), false)
    expect(html).toContain('Bob')
    expect(html).toContain('bob@example.com')
  })

  it('omits CC row when no CC recipients', () => {
    const html = buildPageHtml(makeMessage({ ccRecipients: [] }), false)
    const ccMatches = html.match(/CC/g)
    expect(ccMatches).toBeNull()
  })

  it('includes CC row when CC recipients present', () => {
    const msg = makeMessage({ ccRecipients: [{ name: 'Carol', address: 'carol@example.com' }] })
    const html = buildPageHtml(msg, false)
    expect(html).toContain('CC')
    expect(html).toContain('Carol')
  })

  it('includes sanitized body HTML', () => {
    const msg = makeMessage({ bodyHtml: '<p>Safe content</p><script>evil()</script>' })
    const html = buildPageHtml(msg, false)
    expect(html).toContain('Safe content')
    expect(html).not.toContain('<script')
  })

  it('omits attachments section when includeAttachments is false', () => {
    const msg = makeMessage({
      attachments: [{ name: 'file.pdf', size: 1024, contentType: 'application/pdf' }],
    })
    const html = buildPageHtml(msg, false)
    expect(html).not.toContain('file.pdf')
  })

  it('includes attachments section when includeAttachments is true', () => {
    const msg = makeMessage({
      attachments: [{ name: 'report.pdf', size: 2048, contentType: 'application/pdf' }],
    })
    const html = buildPageHtml(msg, true)
    expect(html).toContain('report.pdf')
    expect(html).toContain('application/pdf')
  })

  it('escapes HTML entities in subject', () => {
    const msg = makeMessage({ subject: '<script>alert("xss")</script>' })
    const html = buildPageHtml(msg, false)
    expect(html).toContain('&lt;script&gt;')
    expect(html).not.toContain('<script>')
  })
})
