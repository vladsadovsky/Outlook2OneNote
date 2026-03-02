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
  getConversationMessagesSimplified(conversationId: string): Promise<EmailMessage[]>
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
  // attachments will be fetched separately if hasAttachments is true
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

function mapMessage(m: GraphMessage, attachments: GraphAttachment[] = []): EmailMessage {
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
    attachments: attachments
      .filter(a => !a.isInline)
      .map(a => ({ name: a.name, size: a.size, contentType: a.contentType })),
  }
}

export function createMailService(client: GraphClient): MailService {
  return {
    async getConversationMessages(conversationId: string): Promise<EmailMessage[]> {
      const messages: GraphMessage[] = []

      // Try the full query first, fall back to simpler query for personal accounts
      let path: string | undefined = '/me/messages'
      let params: Record<string, string> | undefined = {
        '$filter': `conversationId eq '${conversationId}'`,
        '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments',
        '$top': String(PAGE_SIZE),
        '$orderby': 'receivedDateTime asc',
      }

      try {
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
      } catch (error: any) {
        // If we get InefficientFilter error, retry with simplified query for personal accounts
        if (error?.message?.includes('InefficientFilter') || error?.message?.includes('too complex')) {
          debugLog('mailService', 'Complex query failed, trying simplified query for personal accounts...')
          return this.getConversationMessagesSimplified(conversationId)
        }
        throw error
      }

      // After fetching all messages, get attachments for those that have them
      const emailMessages: EmailMessage[] = []
      for (const message of messages) {
        let attachments: GraphAttachment[] = []
        
        if (message.hasAttachments) {
          try {
            const attachmentResponse = await client.get<{ value: GraphAttachment[] }>(
              `/me/messages/${message.id}/attachments`,
              { '$select': 'name,size,contentType,isInline' }
            )
            attachments = attachmentResponse.value
            debugLog('mailService', `Fetched ${attachments.length} attachments for message ${message.id}`)
          } catch (error) {
            debugLog('mailService', `Failed to fetch attachments for message ${message.id}:`, error)
            // Continue with empty attachments array - don't fail the entire export
          }
        }
        
        emailMessages.push(mapMessage(message, attachments))
      }

      // Sort messages by receivedDateTime (since Graph API orderby might have failed)
      emailMessages.sort((a, b) => new Date(a.receivedDateTime).getTime() - new Date(b.receivedDateTime).getTime())
      
      return emailMessages
    },

    // Simplified query method for personal accounts (fallback)
    async getConversationMessagesSimplified(conversationId: string): Promise<EmailMessage[]> {
      debugLog('mailService', 'Using simplified query for personal Microsoft accounts')
      const messages: GraphMessage[] = []

      // Simplified query: remove $orderby which causes issues with personal accounts
      let path: string | undefined = '/me/messages'
      let params: Record<string, string> | undefined = {
        '$filter': `conversationId eq '${conversationId}'`,
        '$select': 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments',
        '$top': String(PAGE_SIZE)
        // NOTE: No $orderby - we'll sort in JavaScript
      }

      while (path !== undefined) {
        const page: GraphMessagesPage = await client.get<GraphMessagesPage>(path, params)
        messages.push(...page.value)

        if (messages.length > THREAD_LIMIT) {
          throw new ThreadTooLargeError(messages.length)
        }

        debugLog('mailService', `Fetched ${messages.length} message(s) so far (simplified query)`)

        path = page['@odata.nextLink']
        params = undefined
      }

      // Get attachments for messages that have them
      const emailMessages: EmailMessage[] = []
      for (const message of messages) {
        let attachments: GraphAttachment[] = []
        
        if (message.hasAttachments) {
          try {
            const attachmentResponse = await client.get<{ value: GraphAttachment[] }>(
              `/me/messages/${message.id}/attachments`,
              { '$select': 'name,size,contentType,isInline' }
            )
            attachments = attachmentResponse.value
            debugLog('mailService', `Fetched ${attachments.length} attachments for message ${message.id}`)
          } catch (error) {
            debugLog('mailService', `Failed to fetch attachments for message ${message.id}:`, error)
          }
        }
        
        emailMessages.push(mapMessage(message, attachments))
      }

      // Sort by receivedDateTime since we can't use $orderby with personal accounts
      emailMessages.sort((a, b) => new Date(a.receivedDateTime).getTime() - new Date(b.receivedDateTime).getTime())
      debugLog('mailService', `Sorted ${emailMessages.length} messages by date`)
      
      return emailMessages
    },
  }
}
