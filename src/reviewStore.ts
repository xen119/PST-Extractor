import { MongoClient } from 'mongodb'
import type { MessageSummary } from './viewer'
import {
  ReviewContext,
  ReviewPatchInput,
  ReviewRecord,
  ReviewSearchOptions,
  ReviewState
} from './reviewTypes'

interface ReviewDocument extends ReviewRecord {}

interface ReviewCursorLike<T> {
  sort(sort: Record<string, 1 | -1>): { toArray(): Promise<T[]> }
}

export interface ReviewCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  findOne: (filter: Record<string, unknown>) => Promise<ReviewDocument | null>
  find: (filter: Record<string, unknown>) => ReviewCursorLike<ReviewDocument>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
  deleteOne: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
}

export interface ReviewStore {
  kind: 'memory' | 'mongo'
  isPersistent: boolean
  getReview(mailboxKey: string, messageId: string, reviewerUsername: string): Promise<ReviewState | null>
  getMany(
    mailboxKey: string,
    messageIds: string[],
    reviewerUsername: string
  ): Promise<Map<string, ReviewState>>
  upsertReview(input: ReviewPatchInput): Promise<ReviewState | null>
  deleteReview(mailboxKey: string, messageId: string, reviewerUsername: string): Promise<boolean>
  listReviews(mailboxKey: string, options?: ReviewSearchOptions): Promise<ReviewRecord[]>
  close(): Promise<void>
}

const DEFAULT_COLLECTION_NAME = 'pst_reviews'

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeReviewerUsername(value: unknown): string {
  return normalizeText(value) || 'anonymous'
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function buildMailboxKey(message: Pick<ReviewContext, 'mailboxKey'>): string {
  return normalizeText(message.mailboxKey)
}

function buildReviewKey(
  mailboxKey: string,
  messageId: string,
  reviewerUsername: string
): string {
  return `${normalizeText(mailboxKey)}::${normalizeText(messageId)}::${normalizeReviewerUsername(
    reviewerUsername
  )}`
}

export function normalizeReviewTags(input: unknown): string[] {
  const rawValues = Array.isArray(input)
    ? input
    : typeof input === 'string'
      ? input.split(/[\n,]/g)
      : []
  const seen = new Set<string>()
  const tags: string[] = []

  for (const rawValue of rawValues) {
    const tag = normalizeText(rawValue)
    if (!tag) {
      continue
    }
    const key = tag.toLowerCase()
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    tags.push(tag)
  }

  return tags
}

export function buildReviewContext(
  mailboxKey: string,
  fileName: string,
  summary: MessageSummary
): ReviewContext {
  return {
    mailboxKey: normalizeText(mailboxKey),
    reviewerUsername: 'anonymous',
    fileName: normalizeText(fileName),
    messageId: normalizeText(summary.id),
    descriptorId: normalizeText(summary.descriptorId),
    folderId: normalizeText(summary.folderId),
    folderPath: normalizeText(summary.folderPath),
    messageClass: normalizeText(summary.messageClass),
    kind: summary.kind,
    isMailLike: summary.isMailLike,
    subject: normalizeText(summary.subject),
    senderName: normalizeText(summary.senderName),
    senderEmailAddress: normalizeText(summary.senderEmailAddress),
    displayTo: normalizeText(summary.displayTo),
    displayCC: normalizeText(summary.displayCC),
    displayBCC: normalizeText(summary.displayBCC),
    resolvedDisplayTo: normalizeText(summary.resolvedDisplayTo),
    resolvedDisplayCC: normalizeText(summary.resolvedDisplayCC),
    resolvedDisplayBCC: normalizeText(summary.resolvedDisplayBCC)
  }
}

export function buildReviewSearchFilter(
  mailboxKey: string,
  options: ReviewSearchOptions = {}
): Record<string, unknown> {
  const filter: Record<string, unknown> = { mailboxKey: normalizeText(mailboxKey) }
  const reviewerUsername = normalizeText(options.reviewerUsername)
  if (reviewerUsername) {
    filter.reviewerUsername = reviewerUsername
  }
  const messageIds = Array.from(new Set((options.messageIds || []).map((value) => normalizeText(value)).filter(Boolean)))

  if (messageIds.length) {
    filter.messageId = { $in: messageIds }
  }

  if (options.flaggedOnly) {
    filter.flagged = true
  }

  if (options.taggedOnly) {
    filter['tags.0'] = { $exists: true }
  }

  if (options.tag) {
    filter.tags = {
      $elemMatch: {
        $regex: `^${escapeRegex(normalizeText(options.tag))}$`,
        $options: 'i'
      }
    }
  }

  if (options.query) {
    const search = normalizeText(options.query)
    if (search) {
      const regex = new RegExp(escapeRegex(search), 'i')
      filter.$or = [
        { subject: regex },
        { senderName: regex },
        { senderEmailAddress: regex },
        { folderPath: regex },
        { displayTo: regex },
        { displayCC: regex },
        { displayBCC: regex },
        { resolvedDisplayTo: regex },
        { resolvedDisplayCC: regex },
        { resolvedDisplayBCC: regex },
        { messageClass: regex },
        { tags: regex }
      ]
    }
  }

  return filter
}

function matchesReviewSearch(record: ReviewRecord, options: ReviewSearchOptions = {}): boolean {
  if (options.reviewerUsername && record.reviewerUsername !== normalizeReviewerUsername(options.reviewerUsername)) {
    return false
  }
  if (options.flaggedOnly && !record.flagged) {
    return false
  }
  if (options.taggedOnly && record.tags.length === 0) {
    return false
  }
  if (options.tag) {
    const needle = normalizeText(options.tag).toLowerCase()
    if (!record.tags.some((tag) => tag.toLowerCase() === needle)) {
      return false
    }
  }
  if (options.messageIds && options.messageIds.length) {
    const ids = new Set(options.messageIds.map((value) => normalizeText(value)))
    if (!ids.has(record.messageId)) {
      return false
    }
  }
  if (options.query) {
    const needle = normalizeText(options.query).toLowerCase()
    const haystack = [
      record.subject,
      record.senderName,
      record.senderEmailAddress,
      record.folderPath,
      record.displayTo,
      record.displayCC,
      record.displayBCC,
      record.resolvedDisplayTo,
      record.resolvedDisplayCC,
      record.resolvedDisplayBCC,
      record.messageClass,
      record.tags.join(' ')
    ]
      .join(' ')
      .toLowerCase()
    if (!haystack.includes(needle)) {
      return false
    }
  }
  return true
}

function sortReviews(records: ReviewRecord[]): ReviewRecord[] {
  return [...records].sort((left, right) => {
    const leftTime = Date.parse(left.updatedAt || left.createdAt || '')
    const rightTime = Date.parse(right.updatedAt || right.createdAt || '')
    if (rightTime !== leftTime) {
      return rightTime - leftTime
    }
    return left.subject.localeCompare(right.subject, undefined, { sensitivity: 'base' })
  })
}

function toReviewState(record: ReviewRecord | null): ReviewState | null {
  if (!record) {
    return null
  }
  return {
    flagged: record.flagged,
    tags: [...record.tags],
    createdAt: record.createdAt,
    updatedAt: record.updatedAt
  }
}

export class MemoryReviewStore implements ReviewStore {
  public kind: 'memory' = 'memory'
  public isPersistent = false
  private readonly records = new Map<string, ReviewRecord>()

  async getReview(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string
  ): Promise<ReviewState | null> {
    const record = this.records.get(buildReviewKey(mailboxKey, messageId, reviewerUsername)) || null
    return toReviewState(record)
  }

  async getMany(
    mailboxKey: string,
    messageIds: string[],
    reviewerUsername: string
  ): Promise<Map<string, ReviewState>> {
    const result = new Map<string, ReviewState>()
    for (const messageId of messageIds) {
      const record = this.records.get(buildReviewKey(mailboxKey, messageId, reviewerUsername))
      if (record) {
        result.set(record.messageId, toReviewState(record) as ReviewState)
      }
    }
    return result
  }

  async upsertReview(input: ReviewPatchInput): Promise<ReviewState | null> {
    const mailboxKey = buildMailboxKey(input)
    const messageId = normalizeText(input.messageId)
    const reviewerUsername = normalizeReviewerUsername(input.reviewerUsername)
    const key = buildReviewKey(mailboxKey, messageId, reviewerUsername)
    const existing = this.records.get(key) || null
    const tags =
      input.tags === undefined ? existing?.tags || [] : normalizeReviewTags(input.tags)
    const flagged = input.flagged === undefined ? existing?.flagged || false : Boolean(input.flagged)

    if (!flagged && tags.length === 0) {
      this.records.delete(key)
      return null
    }

    const now = new Date().toISOString()
    const record: ReviewRecord = {
      mailboxKey,
      reviewerUsername,
      fileName: normalizeText(input.fileName),
      messageId,
      descriptorId: normalizeText(input.descriptorId),
      folderId: normalizeText(input.folderId),
      folderPath: normalizeText(input.folderPath),
      messageClass: normalizeText(input.messageClass),
      kind: input.kind,
      isMailLike: Boolean(input.isMailLike),
      subject: normalizeText(input.subject),
      senderName: normalizeText(input.senderName),
      senderEmailAddress: normalizeText(input.senderEmailAddress),
      displayTo: normalizeText(input.displayTo),
      displayCC: normalizeText(input.displayCC),
      displayBCC: normalizeText(input.displayBCC),
      resolvedDisplayTo: normalizeText(input.resolvedDisplayTo),
      resolvedDisplayCC: normalizeText(input.resolvedDisplayCC),
      resolvedDisplayBCC: normalizeText(input.resolvedDisplayBCC),
      flagged,
      tags,
      createdAt: existing?.createdAt || now,
      updatedAt: now
    }
    this.records.set(key, record)
    return toReviewState(record)
  }

  async deleteReview(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string
  ): Promise<boolean> {
    return this.records.delete(buildReviewKey(mailboxKey, messageId, reviewerUsername))
  }

  async listReviews(mailboxKey: string, options: ReviewSearchOptions = {}): Promise<ReviewRecord[]> {
    const key = normalizeText(mailboxKey)
    return sortReviews(
      [...this.records.values()].filter(
        (record) => record.mailboxKey === key && matchesReviewSearch(record, options)
      )
    )
  }

  async close(): Promise<void> {
    this.records.clear()
  }
}

export class MongoReviewStore implements ReviewStore {
  public kind: 'mongo' = 'mongo'
  public isPersistent = true

  constructor(
    private readonly collection: ReviewCollectionLike,
    private readonly client?: MongoClient
  ) {}

  static async connect(uri: string, dbName = 'pst-extractor'): Promise<MongoReviewStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const collection = client.db(dbName).collection<ReviewDocument>(DEFAULT_COLLECTION_NAME)
    await collection.createIndex({ mailboxKey: 1, messageId: 1, reviewerUsername: 1 }, { unique: true })
    await collection.createIndex({ mailboxKey: 1, flagged: 1 })
    await collection.createIndex({ mailboxKey: 1, updatedAt: -1 })
    await collection.createIndex({ mailboxKey: 1, tags: 1 })
    await collection.createIndex({ mailboxKey: 1, reviewerUsername: 1 })
    return new MongoReviewStore(collection as unknown as ReviewCollectionLike, client)
  }

  async getReview(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string
  ): Promise<ReviewState | null> {
    const record = await this.collection.findOne({
      mailboxKey: normalizeText(mailboxKey),
      messageId: normalizeText(messageId),
      reviewerUsername: normalizeReviewerUsername(reviewerUsername)
    })
    return toReviewState(record)
  }

  async getMany(
    mailboxKey: string,
    messageIds: string[],
    reviewerUsername: string
  ): Promise<Map<string, ReviewState>> {
    const ids = Array.from(new Set(messageIds.map((value) => normalizeText(value)).filter(Boolean)))
    const result = new Map<string, ReviewState>()
    if (!ids.length) {
      return result
    }
    const records = await this.collection
      .find({
        mailboxKey: normalizeText(mailboxKey),
        reviewerUsername: normalizeReviewerUsername(reviewerUsername),
        messageId: { $in: ids }
      })
      .sort({ updatedAt: -1 })
      .toArray()
    for (const record of records) {
      result.set(record.messageId, toReviewState(record) as ReviewState)
    }
    return result
  }

  async upsertReview(input: ReviewPatchInput): Promise<ReviewState | null> {
    const mailboxKey = buildMailboxKey(input)
    const messageId = normalizeText(input.messageId)
    const reviewerUsername = normalizeReviewerUsername(input.reviewerUsername)
    const existing = await this.collection.findOne({ mailboxKey, messageId, reviewerUsername })
    const tags =
      input.tags === undefined ? existing?.tags || [] : normalizeReviewTags(input.tags)
    const flagged = input.flagged === undefined ? existing?.flagged || false : Boolean(input.flagged)

    if (!flagged && tags.length === 0) {
      await this.collection.deleteOne({ mailboxKey, messageId, reviewerUsername })
      return null
    }

    const now = new Date().toISOString()
    const record: ReviewRecord = {
      mailboxKey,
      reviewerUsername,
      fileName: normalizeText(input.fileName),
      messageId,
      descriptorId: normalizeText(input.descriptorId),
      folderId: normalizeText(input.folderId),
      folderPath: normalizeText(input.folderPath),
      messageClass: normalizeText(input.messageClass),
      kind: input.kind,
      isMailLike: Boolean(input.isMailLike),
      subject: normalizeText(input.subject),
      senderName: normalizeText(input.senderName),
      senderEmailAddress: normalizeText(input.senderEmailAddress),
      displayTo: normalizeText(input.displayTo),
      displayCC: normalizeText(input.displayCC),
      displayBCC: normalizeText(input.displayBCC),
      resolvedDisplayTo: normalizeText(input.resolvedDisplayTo),
      resolvedDisplayCC: normalizeText(input.resolvedDisplayCC),
      resolvedDisplayBCC: normalizeText(input.resolvedDisplayBCC),
      flagged,
      tags,
      createdAt: existing?.createdAt || now,
      updatedAt: now
    }

    await this.collection.updateOne(
      { mailboxKey, messageId, reviewerUsername },
      {
        $set: record
      },
      { upsert: true }
    )
    return toReviewState(record)
  }

  async deleteReview(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string
  ): Promise<boolean> {
    const result = await this.collection.deleteOne({
      mailboxKey: normalizeText(mailboxKey),
      messageId: normalizeText(messageId),
      reviewerUsername: normalizeReviewerUsername(reviewerUsername)
    })
    return Boolean(result.deletedCount && result.deletedCount > 0)
  }

  async listReviews(mailboxKey: string, options: ReviewSearchOptions = {}): Promise<ReviewRecord[]> {
    const records = await this.collection
      .find(buildReviewSearchFilter(mailboxKey, options))
      .sort({ updatedAt: -1 })
      .toArray()
    return sortReviews(records)
  }

  async close(): Promise<void> {
    if (this.client) {
      await this.client.close()
    }
  }
}

export async function createReviewStoreFromEnv(
  env: NodeJS.ProcessEnv = process.env
): Promise<ReviewStore> {
  const uri = normalizeText(env.MONGODB_URI)
  if (!uri) {
    return new MemoryReviewStore()
  }
  const dbName = normalizeText(env.MONGODB_DB) || 'pst-extractor'
  return MongoReviewStore.connect(uri, dbName)
}
