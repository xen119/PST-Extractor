import {
  MemorySearchIndexStore,
  type SearchIndexDocument
} from '../searchIndex'
import { createMailboxCollapseCoordinator } from '../searchIndexCollapse'

function makeDocument(messageId: string, date: string, inReplyToId = ''): SearchIndexDocument {
  return {
    id: messageId,
    sourceType: 'mailbox',
    mailboxKey: 'C:/pst/case/search/mailbox.pst',
    scopePath: 'case/search',
    scopeLabel: 'Case / Search',
    fileName: 'mailbox.pst',
    mailboxName: 'mailbox.pst',
    messageId,
    descriptorId: messageId,
    folderId: 'folder',
    folderPath: 'Inbox',
    order: 1,
    messageClass: 'IPM.Note',
    kind: 'mail',
    size: 100,
    subject: 'Thread subject',
    originalSubject: 'Thread subject',
    senderName: 'Sender',
    senderEmailAddress: 'sender@example.com',
    recipientText: 'recipient@example.com',
    displayTo: 'recipient@example.com',
    displayCC: '',
    displayBCC: '',
    resolvedDisplayTo: 'recipient@example.com',
    resolvedDisplayCC: '',
    resolvedDisplayBCC: '',
    clientSubmitTime: date,
    creationTime: date,
    modificationTime: date,
    messageDeliveryTime: date,
    sortDate: date,
    sortDateMs: Date.parse(date),
    importance: 0,
    hasAttachments: false,
    isRead: true,
    isMailLike: true,
    bodySearchText: `body ${messageId}`,
    searchText: `thread ${messageId}`,
    searchTokens: ['thread', messageId],
    addressValues: ['sender@example.com', 'recipient@example.com'],
    subjectValues: ['thread subject'],
    review: { flagged: false, tags: [], createdAt: '', updatedAt: '' },
    reviewStates: [],
    reviewTagValues: [],
    threadMetadata: {
      messageId: `<${messageId}@example.com>`,
      inReplyToId: inReplyToId ? `<${inReplyToId}@example.com>` : '',
      referenceIds: inReplyToId ? [`${inReplyToId}@example.com`] : [],
      conversationId: '',
      isForward: false
    },
    updatedAt: date
  }
}

async function waitForCompletion(
  coordinator: ReturnType<typeof createMailboxCollapseCoordinator>
): Promise<void> {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    const status = await coordinator.getStatus()
    if (status.status === 'succeeded' || status.status === 'failed' || status.status === 'reindex-required') {
      return
    }
    await new Promise<void>((resolve) => setImmediate(resolve))
  }
}

describe('incremental mailbox collapse coordinator', () => {
  it('processes existing Mongo-shaped documents asynchronously and exposes final representatives', async () => {
    const store = new MemorySearchIndexStore()
    await store.replaceMailboxDocuments('C:/pst/case/search/mailbox.pst', [
      makeDocument('root', '2025-01-01T00:00:00.000Z'),
      makeDocument('reply', '2025-01-02T00:00:00.000Z', 'root')
    ])

    const coordinator = createMailboxCollapseCoordinator({ searchIndexStore: store })
    const started = await coordinator.start()
    expect(started.status).toBe('running')

    await waitForCompletion(coordinator)
    const status = await coordinator.getStatus()
    expect(status.status).toBe('succeeded')
    expect(status.percentage).toBe(100)
    expect(status.provisional).toBe(false)

    const page = await store.search({
      scope: 'all',
      sourceType: 'mailbox',
      query: 'root',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 50,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true,
      collapseProgress: status
    })
    expect(page.items.map((item) => item.messageId)).toEqual(['reply'])
  })
})

