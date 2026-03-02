import type { EmailMessage } from './types'
import { sanitizeHtml } from './htmlSanitizer'

// No imports from outside src/onenote/ (boundary rule — DESIGN.md 6.1).

function formatDate(isoDate: string): string {
  const d = new Date(isoDate)
  const y = d.getFullYear()
  const mo = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  const h = String(d.getHours()).padStart(2, '0')
  const min = String(d.getMinutes()).padStart(2, '0')
  return `${y}-${mo}-${day} ${h}:${min}`
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function formatAddresses(addrs: { name: string; address: string }[]): string {
  if (addrs.length === 0) return ''
  return addrs
    .map(a => (a.name ? `${escapeHtml(a.name)} &lt;${escapeHtml(a.address)}&gt;` : escapeHtml(a.address)))
    .join(', ')
}

function buildMetadataTable(message: EmailMessage): string {
  const rows = [
    ['From', formatAddresses([message.from])],
    ['To', formatAddresses(message.toRecipients)],
    ...(message.ccRecipients.length > 0 ? [['CC', formatAddresses(message.ccRecipients)]] : []),
    ['Date', escapeHtml(formatDate(message.receivedDateTime))],
    ['Subject', escapeHtml(message.subject)],
  ]
  const rowsHtml = rows
    .map(([label, value]) =>
      `<tr>` +
      `<td width="22%" style="width:22%;vertical-align:top;white-space:nowrap">` +
      `<b>${label}</b>` +
      `</td>` +
      `<td width="78%" style="width:78%;word-break:break-word">${value}</td>` +
      `</tr>`
    )
    .join('\n')
  return `
<table border="1" style="border-collapse:collapse;width:100%;table-layout:fixed">
  <colgroup>
    <col style="width:22%" />
    <col style="width:78%" />
  </colgroup>
${rowsHtml}
</table>`
}

function buildAttachmentsSection(message: EmailMessage): string {
  const real = message.attachments.filter(a => a.name)
  if (real.length === 0) return ''

  const rows = real
    .map(a => {
      const sizeKb = Math.ceil(a.size / 1024)
      return `<tr><td>${escapeHtml(a.name)}</td><td>${escapeHtml(a.contentType)}</td><td>${sizeKb} KB</td></tr>`
    })
    .join('\n')

  return `
<br />
<p><b>Attachments</b></p>
<table border="1" style="border-collapse:collapse;width:100%">
  <tr><th>Name</th><th>Type</th><th>Size</th></tr>
  ${rows}
</table>`
}

// Builds a OneNote-compatible HTML page from a generic EmailMessage.
// The returned string is the full HTML document to POST to the Graph API.
export function buildPageHtml(message: EmailMessage, includeAttachments: boolean): string {
  const title = escapeHtml(message.subject)
  const created = message.receivedDateTime
  const metadata = buildMetadataTable(message)
  const body = sanitizeHtml(message.bodyHtml)
  const attachments = includeAttachments ? buildAttachmentsSection(message) : ''

  return `<!DOCTYPE html>
<html>
<head>
  <title>${title}</title>
  <meta name="created" content="${created}" />
</head>
<body>
  <div data-id="metadata">
    ${metadata}
  </div>
  <br />
  <div data-id="body">
    ${body}
  </div>
  ${attachments}
</body>
</html>`
}
