import { describe, it, expect } from 'vitest'
import { sanitizeHtml } from '@/onenote/htmlSanitizer'

describe('sanitizeHtml', () => {
  it('removes <script> tags and their content', () => {
    const result = sanitizeHtml('<p>Hello</p><script>alert("xss")</script><p>World</p>')
    expect(result).not.toContain('<script')
    expect(result).not.toContain('alert')
    expect(result).toContain('Hello')
    expect(result).toContain('World')
  })

  it('removes <iframe> tags', () => {
    const result = sanitizeHtml('<p>Text</p><iframe src="evil.com"></iframe>')
    expect(result).not.toContain('<iframe')
    expect(result).toContain('Text')
  })

  it('removes event handler attributes', () => {
    const result = sanitizeHtml('<p onclick="evil()">Click me</p>')
    expect(result).not.toContain('onclick')
    expect(result).toContain('Click me')
  })

  it('removes onerror and other on* attributes', () => {
    const result = sanitizeHtml('<img src="x" onerror="evil()" />')
    expect(result).not.toContain('onerror')
  })

  it('removes javascript: href attributes', () => {
    const result = sanitizeHtml('<a href="javascript:evil()">click</a>')
    expect(result).not.toContain('javascript:')
    expect(result).toContain('click')
  })

  it('strips external image src attributes', () => {
    const result = sanitizeHtml('<img src="https://evil.com/tracker.png" alt="img" />')
    expect(result).not.toContain('https://evil.com')
    expect(result).toContain('alt=')
  })

  it('preserves safe HTML content', () => {
    const result = sanitizeHtml('<p>Hello <strong>world</strong></p><ul><li>item</li></ul>')
    expect(result).toContain('<p>')
    expect(result).toContain('<strong>')
    expect(result).toContain('<ul>')
    expect(result).toContain('item')
  })

  it('removes style attributes with CSS expression()', () => {
    const result = sanitizeHtml('<p style="width:expression(alert(1))">bad</p>')
    expect(result).not.toContain('expression(')
  })

  it('removes style attributes with javascript: in CSS', () => {
    const result = sanitizeHtml('<p style="background:url(javascript:evil())">bad</p>')
    expect(result).not.toContain('javascript:')
  })

  it('handles empty string input', () => {
    expect(sanitizeHtml('')).toBe('')
  })

  it('removes <object> and <embed> tags', () => {
    const result = sanitizeHtml('<object data="flash.swf"></object><embed src="x.swf" />')
    expect(result).not.toContain('<object')
    expect(result).not.toContain('<embed')
  })
})
