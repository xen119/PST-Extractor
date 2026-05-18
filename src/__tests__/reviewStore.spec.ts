import {
  buildReviewContext,
  buildReviewSearchFilter,
  MongoReviewStore,
  normalizeReviewTags,
  type ReviewCollectionLike
} from '../reviewStore'
import type { ReviewRecord } from '../reviewTypes'

function makeContext() {
  return buildReviewContext('/mailboxes/team.pst', 'team.pst', {
    id: 'message-1',
    descriptorId: 'descriptor-1',
    folderId: 'folder-1',
    folderPath: 'Mailbox/Inbox',
    order: 1,
    messageClass: 'IPM.Note',
    kind: 'mail',
    subject: 'Review me',
    senderName: 'Alice',
    senderEmailAddress: 'alice@example.com',
    recipientText: 'Bob',
    displayTo: 'Bob',
    displayCC: '',
    displayBCC: '',
    resolvedDisplayTo: 'Bob <bob@example.com>',
    resolvedDisplayCC: '',
    resolvedDisplayBCC: '',
    originalSubject: 'Review me',
    clientSubmitTime: '2024-01-01T10:00:00.000Z',
    creationTime: '2024-01-01T10:00:00.000Z',
    modificationTime: '2024-01-01T10:00:00.000Z',
    messageDeliveryTime: '2024-01-01T10:00:00.000Z',
    sortDate: '2024-01-01T10:00:00.000Z',
    sortDateMs: Date.parse('2024-01-01T10:00:00.000Z'),
    importance: 1,
    hasAttachments: true,
    isRead: false,
    isMailLike: true
  })
}

function makeRecord(overrides: Partial<ReviewRecord> = {}): ReviewRecord {
  return {
    mailboxKey: '/mailboxes/team.pst',
    fileName: 'team.pst',
    messageId: 'message-1',
    descriptorId: 'descriptor-1',
    folderId: 'folder-1',
    folderPath: 'Mailbox/Inbox',
    messageClass: 'IPM.Note',
    kind: 'mail',
    isMailLike: true,
    subject: 'Review me',
    senderName: 'Alice',
    senderEmailAddress: 'alice@example.com',
    displayTo: 'Bob',
    displayCC: '',
    displayBCC: '',
    resolvedDisplayTo: 'Bob <bob@example.com>',
    resolvedDisplayCC: '',
    resolvedDisplayBCC: '',
    flagged: false,
    tags: [],
    createdAt: '2024-01-01T10:00:00.000Z',
    updatedAt: '2024-01-01T10:00:00.000Z',
    ...overrides
  }
}

function matchesClause(record: ReviewRecord, clause: Record<string, unknown>): boolean {
  return Object.entries(clause).every(([field, expected]) => {
    const value = String((record as Record<string, unknown>)[field] ?? '')
    if (expected instanceof RegExp) {
      return expected.test(value)
    }
    return true
  })
}

function matchesFilter(record: ReviewRecord, filter: Record<string, unknown>): boolean {
  if (filter.mailboxKey && record.mailboxKey !== filter.mailboxKey) {
    return false
  }

  if ('messageId' in filter) {
    if (typeof filter.messageId === 'string' && record.messageId !== filter.messageId) {
      return false
    }
    if (filter.messageId && typeof filter.messageId === 'object' && '$in' in filter.messageId) {
      const values = (filter.messageId as { $in?: string[] }).$in || []
      if (!values.includes(record.messageId)) {
        return false
      }
    }
  }

  if (filter.flagged === true && !record.flagged) {
    return false
  }

  if (
    filter['tags.0'] &&
    typeof filter['tags.0'] === 'object' &&
    '$exists' in filter['tags.0'] &&
    !record.tags.length
  ) {
    return false
  }

  if (filter.tags && typeof filter.tags === 'object' && '$elemMatch' in filter.tags) {
    const matcher = filter.tags as {
      $elemMatch?: { $regex?: string; $options?: string }
    }
    const regex = matcher.$elemMatch?.$regex
      ? new RegExp(matcher.$elemMatch.$regex, matcher.$elemMatch.$options || 'i')
      : null
    if (regex && !record.tags.some((tag) => regex.test(tag))) {
      return false
    }
  }

  if (Array.isArray(filter.$or)) {
    const clauses = filter.$or as Array<Record<string, unknown>>
    if (!clauses.some((clause) => matchesClause(record, clause))) {
      return false
    }
  }

  return true
}

function createFakeCollection(
  records: ReviewRecord[] = [],
  onUpdate?: (filter: Record<string, unknown>, update: Record<string, unknown>, options?: { upsert?: boolean }) => void
): ReviewCollectionLike {
  const storage = [...records]

  return {
    async createIndex() {
      return undefined
    },
    async findOne(filter) {
      return storage.find((record) => matchesFilter(record, filter)) || null
    },
    find(filter) {
      return {
        sort(sortSpec: Record<string, 1 | -1>) {
          return {
            async toArray() {
              const filtered = storage.filter((record) => matchesFilter(record, filter))
              const sorted = [...filtered]
              if (sortSpec.updatedAt === -1) {
                sorted.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt))
              }
              return sorted
            }
          }
        }
      }
    },
    async updateOne(filter, update, options) {
      onUpdate?.(filter, update, options)
      const mailboxKey = String((filter as Record<string, unknown>).mailboxKey || '')
      const messageId = String((filter as Record<string, unknown>).messageId || '')
      const existingIndex = storage.findIndex(
        (record) => record.mailboxKey === mailboxKey && record.messageId === messageId
      )
      const nextRecord = {
        ...(options?.upsert ? {} : storage[existingIndex] || {}),
        ...((update.$set as Record<string, unknown>) || {})
      } as ReviewRecord

      if (existingIndex >= 0) {
        storage[existingIndex] = nextRecord
      } else if (options?.upsert) {
        storage.push(nextRecord)
      }
      return {}
    },
    async deleteOne(filter) {
      const mailboxKey = String((filter as Record<string, unknown>).mailboxKey || '')
      const messageId = String((filter as Record<string, unknown>).messageId || '')
      const originalLength = storage.length
      for (let index = storage.length - 1; index >= 0; index--) {
        const record = storage[index]
        if (record.mailboxKey === mailboxKey && record.messageId === messageId) {
          storage.splice(index, 1)
        }
      }
      return { deletedCount: originalLength - storage.length }
    }
  }
}

describe('review store helpers', () => {
  it('normalizes review tags and builds search filters', () => {
    expect(normalizeReviewTags('Urgent, urgent\n follow up ')).toEqual(['Urgent', 'follow up'])

    const filter = buildReviewSearchFilter('/mailboxes/team.pst', {
      query: 'alice',
      flaggedOnly: true,
      taggedOnly: true,
      tag: 'Urgent',
      messageIds: ['message-1', 'message-2']
    })

    expect(filter.mailboxKey).toBe('/mailboxes/team.pst')
    expect(filter.messageId).toEqual({ $in: ['message-1', 'message-2'] })
    expect(filter.flagged).toBe(true)
    expect(filter['tags.0']).toEqual({ $exists: true })
    expect(filter.tags).toEqual({
      $elemMatch: {
        $regex: '^Urgent$',
        $options: 'i'
      }
    })
    expect(Array.isArray(filter.$or)).toBe(true)
  })

  it('upserts, clears, and queries reviews with the mongo-backed store API', async () => {
    const collection = createFakeCollection([
      makeRecord({
        messageId: 'message-2',
        subject: 'Second review',
        flagged: true,
        tags: ['Priority'],
        updatedAt: '2024-01-02T10:00:00.000Z'
      })
    ])
    const store = new MongoReviewStore(collection)
    const context = makeContext()

    const review = await store.upsertReview({
      ...context,
      flagged: true,
      tags: [' Urgent ', 'urgent', 'Follow up']
    })

    expect(review).toEqual({
      flagged: true,
      tags: ['Urgent', 'Follow up'],
      createdAt: expect.any(String),
      updatedAt: expect.any(String)
    })

    expect(await store.getReview(context.mailboxKey, context.messageId)).toEqual(review)

    const queue = await store.listReviews(context.mailboxKey, {
      flaggedOnly: true,
      tag: 'urgent'
    })
    expect(queue).toHaveLength(1)
    expect(queue[0].messageId).toBe(context.messageId)

    const searchResults = await store.listReviews(context.mailboxKey, {
      query: 'review'
    })
    expect(searchResults.map((record) => record.messageId)).toEqual([
      context.messageId,
      'message-2'
    ])

    await expect(store.deleteReview(context.mailboxKey, context.messageId)).resolves.toBe(true)
    expect(await store.getReview(context.mailboxKey, context.messageId)).toBeNull()

    await store.close()
  })

  it('upserts reviews without conflicting mongo update operators', async () => {
    let lastUpdate: Record<string, unknown> | null = null
    const collection = createFakeCollection([], (_filter, update) => {
      lastUpdate = update
    })
    const store = new MongoReviewStore(collection)
    const context = makeContext()

    await store.upsertReview({
      ...context,
      flagged: true,
      tags: ['Urgent']
    })

    expect(lastUpdate).toBeTruthy()
    expect(lastUpdate).toMatchObject({
      $set: expect.objectContaining({
        flagged: true,
        tags: ['Urgent']
      })
    })
    expect(lastUpdate).not.toHaveProperty('$setOnInsert')
  })
})
