// Generic email message — no Outlook/Graph-specific types.
// This is the contract between Outlook2OneNote and the reusable OneNote library.
// Changes to this interface are breaking changes for any consumer (e.g. llm-aggregator).

export interface EmailMessage {
  id: string
  subject: string
  from: EmailAddress
  toRecipients: EmailAddress[]
  ccRecipients: EmailAddress[]
  receivedDateTime: string        // ISO 8601
  bodyHtml: string                // raw HTML body
  attachments: AttachmentMetadata[]
}

export interface EmailAddress {
  name: string
  address: string
}

export interface AttachmentMetadata {
  name: string
  size: number
  contentType: string
}

export interface Notebook {
  id: string
  displayName: string
  webUrl: string
}

export interface Section {
  id: string
  displayName: string
}

export interface Page {
  id: string
  webUrl: string
  oneNoteClientUrl: string
}
