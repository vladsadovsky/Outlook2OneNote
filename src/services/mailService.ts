import type { GraphClient } from '@/services/graphClient'
import type { EmailMessage } from '@/onenote/types'
import { debugLog } from '@/utils/logger'

export class ThreadTooLargeError extends Error {
  constructor(public readonly count: number) {
    super(`Thread has ${count}+ messages, exceeding the 50-message limit`)
    this.name = 'ThreadTooLargeError'
  }
}

export interface MailService {
  getConversationMessages(conversationId: string): Promise<EmailMessage[]>
}

// Graph API shape for a message list page
interface GraphMessagesPage {
  value: GraphMessage[]
  '@odata.nextLink'?: string
}

// Minimal Graph Message shape — only the fields we select
interface GraphMessage {
  id: string
  subject: string
  from: { emailAddress: { name: string; address: string } }
  toRecipients: { emailAddress: { name: string; address: string } }[]
  ccRecipients: { emailAddress: { name: string; address: string } }[]
  receivedDateTime: string
  body: { content: string; contentType: string }
  hasAttachments: boolean
  attachments?: GraphAttachment[]
}

interface GraphAttachment {
  name: string
  size: number
  contentType: string
  isInline: boolean
}

const THREAD_LIMIT = 50
// Fetch one over the limit so we can detect a breach without loading the full thread
const PAGE_SIZE = 51

function mapMessage(m: GraphMessage): EmailMessage {
  return {
    id: m.id,
    subject: m.subject,
    from: { name: m.from.emailAddress.name, address: m.from.emailAddress.address },
    toRecipients: m.toRecipients.map(r => ({
      name: r.emailAddress.name,
      address: r.emailAddress.address,
    })),
    ccRecipients: m.ccRecipients.map(r => ({
      name: r.emailAddress.name,
      address: r.emailAddress.address,
    })),
    receivedDateTime: m.receivedDateTime,
    bodyHtml: m.body.content,
    attachments: (m.attachments ?? [])
      .filter(a => !a.isInline)
      .map(a => ({ name: a.name, size: a.size, contentType: a.contentType })),
  }
}

export function createMailService(client: GraphClient): MailService {
  return {
    async getConversationMessages(conversationId: string): Promise<EmailMessage[]> {
      const messages: GraphMessage[] = []

      // First page: include full params
      let path: string | undefined = '/me/messages'
      let params: Record<string, string> | undefined = {
        '$filter': `conversationId eq '${conversationId}'`,
        '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments',
        '$expand': 'attachments($select=name,size,contentType,isInline)',
        '$top': String(PAGE_SIZE),
        '$orderby': 'receivedDateTime asc',
      }

      while (path !== undefined) {
        const page: GraphMessagesPage = await client.get<GraphMessagesPage>(path, params)
        messages.push(...page.value)

        if (messages.length > THREAD_LIMIT) {
          throw new ThreadTooLargeError(messages.length)
        }

        debugLog('mailService', `Fetched ${messages.length} message(s) so far`)

        // nextLink already contains all query params — pass as path with no extra params
        path = page['@odata.nextLink']
        params = undefined
      }

      return messages.map(mapMessage)
    },
  }
}
