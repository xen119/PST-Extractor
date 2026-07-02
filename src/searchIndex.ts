import { randomBytes } from 'crypto'
import * as path from 'path'
import { MongoClient } from 'mongodb'
import type { ReviewStore } from './reviewStore'
import type { ReviewRecord, ReviewState } from './reviewTypes'
import type { MessageSummary, ViewerSessionIndex } from './viewer'
import type { ArchiveBundleItem, ArchiveBundleSourceType } from './archiveBundles'

export type SearchScope = 'all' | 'search' | 'pst'
export type SearchMode = 'and' | 'or'
export type HiddenRuleKind = 'address' | 'subject'
export type SearchSourceType = 'mailbox' | ArchiveBundleSourceType

export interface HiddenRuleRecord {
  filterId: string
  kind: HiddenRuleKind
  value: string
  label: string
  createdAt: string
  updatedAt: string
}

export interface SearchIndexDocument {
  id?: string
  sourceType: SearchSourceType
  mailboxKey: string
  scopePath: string
  scopeLabel: string
  fileName: string
  mailboxName: string
  messageId: string
  descriptorId: string
  folderId: string
  folderPath: string
  order: number
  messageClass: string
  kind: MessageSummary['kind']
  subject: string
  originalSubject: string
  senderName: string
  senderEmailAddress: string
  recipientText: string
  displayTo: string
  displayCC: string
  displayBCC: string
  resolvedDisplayTo: string
  resolvedDisplayCC: string
  resolvedDisplayBCC: string
  clientSubmitTime: string | null
  creationTime: string | null
  modificationTime: string | null
  messageDeliveryTime: string | null
  sortDate: string | null
  sortDateMs: number | null
  importance: number
  hasAttachments: boolean
  isRead: boolean
  isMailLike: boolean
  bodySearchText: string
  searchText: string
  searchTokens: string[]
  addressValues: string[]
  subjectValues: string[]
  archivePath?: string
  archiveEntryPath?: string
  archiveEntryChain?: string[]
  archiveEntryName?: string
  contentType?: string
  downloadFilename?: string
  previewKind?: 'text' | 'html' | 'binary'
  previewText?: string
  previewHtml?: string
  review: ReviewState
  reviewStates: Array<{
    reviewerUsername: string
    review: ReviewState
  }>
  reviewTagValues: string[]
  updatedAt: string
}

export interface SearchIndexSearchOptions {
  scope: SearchScope
  scopePath?: string
  mailboxKey?: string
  allowedMailboxKeys?: string[]
  reviewerUsername?: string
  sourceType?: SearchSourceType | 'all'
  query: string
  mode: SearchMode
  mailOnly: boolean
  sort: string
  page: number
  pageSize: number
  reviewFlaggedOnly: boolean
  reviewTaggedOnly: boolean
  reviewTag: string
}

export interface SearchIndexPage {
  items: SearchIndexDocument[]
  total: number
  page: number
  pageSize: number
  totalPages: number
  query: string
  mode: SearchMode
  mailOnly: boolean
  sort: string
  scope: SearchScope
  scopePath: string
  scopeLabel: string
  hiddenRules: HiddenRuleRecord[]
  sourceType: SearchSourceType | 'all'
  sourceCounts: Record<SearchSourceType, number>
  reviewFilters: {
    flaggedOnly: boolean
    taggedOnly: boolean
    tag: string
  }
}

export interface SearchIndexStore {
  kind: 'memory' | 'mongo'
  isPersistent: boolean
  replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void>
  deleteMailboxDocuments(mailboxKey: string): Promise<void>
  updateReviewState(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string,
    review: ReviewState | null
  ): Promise<void>
  clearAllDocuments(): Promise<void>
  listHiddenRules(): Promise<HiddenRuleRecord[]>
  upsertHiddenRule(input: {
    kind: HiddenRuleKind
    value: string
    label?: string
  }): Promise<HiddenRuleRecord>
  deleteHiddenRule(filterId: string): Promise<boolean>
  search(options: SearchIndexSearchOptions): Promise<SearchIndexPage>
  findDocumentById(id: string): Promise<SearchIndexDocument | null>
  promoteStagedDocuments?(stagingDocumentsCollectionName: string): Promise<void>
  close(): Promise<void>
}

interface SearchIndexCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  insertMany: (documents: SearchIndexDocument[]) => Promise<unknown>
  deleteMany: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
  findOne: (filter: Record<string, unknown>) => Promise<SearchIndexDocument | null>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
  find: (filter: Record<string, unknown>) => {
    sort: (sort: Record<string, 1 | -1>) => {
      skip: (count: number) => {
        limit: (count: number) => {
          toArray: () => Promise<SearchIndexDocument[]>
        }
      }
    }
  }
  countDocuments: (filter: Record<string, unknown>) => Promise<number>
}

interface HiddenRuleCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  find: (filter: Record<string, unknown>) => {
    sort: (sort: Record<string, 1 | -1>) => {
      toArray: () => Promise<HiddenRuleRecord[]>
    }
  }
  findOne: (filter: Record<string, unknown>) => Promise<HiddenRuleRecord | null>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
  deleteOne: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
  deleteMany: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
}

const DEFAULT_INDEX_COLLECTION = 'pst_search_documents'
const DEFAULT_RULE_COLLECTION = 'pst_search_hidden_rules'

export interface MongoSearchIndexStoreConnectOptions {
  documentsCollectionName?: string
  rulesCollectionName?: string
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeExactValue(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

function normalizeReviewerUsername(value: unknown): string {
  return normalizeText(value) || 'anonymous'
}

function normalizeSourceType(value: unknown): SearchSourceType {
  const normalized = normalizeText(value).toLowerCase()
  if (normalized === 'teams' || normalized === 'sharepoint') {
    return normalized
  }
  return 'mailbox'
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function uniqueStrings(values: string[]): string[] {
  return [...new Set(values.map((value) => normalizeExactValue(value)).filter(Boolean))]
}

function uniqueTextValues(values: string[]): string[] {
  return [...new Set(values.map((value) => normalizeText(value)).filter(Boolean))]
}

function dedupeSearchIndexDocuments(documents: SearchIndexDocument[]): SearchIndexDocument[] {
  const seen = new Set<string>()
  const deduped: SearchIndexDocument[] = []

  for (const document of documents) {
    const mailboxKey = normalizeText(document.mailboxKey)
    const messageId = normalizeText(document.messageId)
    if (!mailboxKey || !messageId) {
      continue
    }

    const dedupeKey = `${mailboxKey}\u0000${messageId}`
    if (seen.has(dedupeKey)) {
      continue
    }

    seen.add(dedupeKey)
    deduped.push({
      ...document,
      mailboxKey,
      sourceType: normalizeSourceType(document.sourceType)
    })
  }

  return deduped
}

function tokenizeSearchText(value: string): string[] {
  return uniqueStrings(
    normalizeExactValue(value)
      .match(/[a-z0-9@._%+-]+/g)
      ?.map((token) => token) || []
  )
}

function parseSearchTerms(
  query: string,
  fallbackMode: SearchMode = 'and'
): {
  mode: SearchMode
  terms: Array<{ text: string; phrase: boolean }>
} {
  const text = normalizeText(query)
  if (!text) {
    return {
      mode: fallbackMode,
      terms: []
    }
  }

  const terms: Array<{ text: string; phrase: boolean }> = []
  let mode: SearchMode = fallbackMode
  const pattern = /"([^"]+)"|(\S+)/g
  let match: RegExpExecArray | null
  while ((match = pattern.exec(text))) {
    const raw = match[1] || match[2] || ''
    const normalized = normalizeText(raw)
    if (normalized) {
      if (!match[1]) {
        const lowered = normalized.toLowerCase()
        if (lowered === '|' || lowered.startsWith('|')) {
          mode = 'or'
          const remainder = normalized.length > 1 ? normalizeText(normalized.slice(1)) : ''
          if (remainder) {
            terms.push({
              text: remainder,
              phrase: false
            })
          }
          continue
        }
        if (lowered === '+' || lowered.startsWith('+')) {
          mode = 'and'
          const remainder = normalized.length > 1 ? normalizeText(normalized.slice(1)) : ''
          if (remainder) {
            terms.push({
              text: remainder,
              phrase: false
            })
          }
          continue
        }
      }
      terms.push({
        text: normalized,
        phrase: Boolean(match[1])
      })
    }
  }
  return {
    mode,
    terms
  }
}

function extractEmailAddresses(...values: Array<string | undefined | null>): string[] {
  const emails = new Set<string>()
  for (const value of values) {
    const text = normalizeText(value)
    if (!text) {
      continue
    }
    const matches = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []
    for (const match of matches) {
      emails.add(normalizeExactValue(match))
    }
  }
  return [...emails]
}

function buildAddressMatchValues(...values: Array<string | undefined | null>): string[] {
  const matches = new Set<string>()
  for (const value of values) {
    const text = normalizeText(value)
    if (!text) {
      continue
    }
    matches.add(normalizeExactValue(text))
    for (const email of extractEmailAddresses(text)) {
      matches.add(email)
    }
  }
  return [...matches]
}

function buildSearchTextParts(
  document: Pick<
    SearchIndexDocument,
    | 'subject'
    | 'originalSubject'
    | 'senderName'
    | 'senderEmailAddress'
    | 'recipientText'
    | 'displayTo'
    | 'displayCC'
    | 'displayBCC'
    | 'resolvedDisplayTo'
    | 'resolvedDisplayCC'
    | 'resolvedDisplayBCC'
    | 'messageClass'
    | 'kind'
    | 'previewText'
    | 'previewHtml'
  >,
  bodySearchText: string
): string[] {
  return [
    document.subject,
    document.originalSubject,
    document.senderName,
    document.senderEmailAddress,
    document.recipientText,
    document.displayTo,
    document.displayCC,
    document.displayBCC,
    document.resolvedDisplayTo,
    document.resolvedDisplayCC,
    document.resolvedDisplayBCC,
    document.messageClass,
    document.kind,
    document.previewText,
    document.previewHtml,
    bodySearchText
  ].filter((value): value is string => Boolean(value))
}

function buildSearchText(
  document: Pick<
    SearchIndexDocument,
    | 'subject'
    | 'originalSubject'
    | 'senderName'
    | 'senderEmailAddress'
    | 'recipientText'
    | 'displayTo'
    | 'displayCC'
    | 'displayBCC'
    | 'resolvedDisplayTo'
    | 'resolvedDisplayCC'
    | 'resolvedDisplayBCC'
    | 'messageClass'
    | 'kind'
    | 'previewText'
    | 'previewHtml'
  >,
  bodySearchText: string
): string {
  return normalizeExactValue(buildSearchTextParts(document, bodySearchText).join(' '))
}

function normalizeReview(review: ReviewState | null | undefined): ReviewState {
  return (
    review || {
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    }
  )
}

function resolveReviewState(
  reviewStates: Array<{ reviewerUsername: string; review: ReviewState }> | undefined,
  reviewerUsername: string | undefined
): ReviewState | null {
  const normalizedReviewerUsername = normalizeReviewerUsername(reviewerUsername)
  const entry = (reviewStates || []).find(
    (state) => normalizeReviewerUsername(state.reviewerUsername) === normalizedReviewerUsername
  )
  return entry ? normalizeReview(entry.review) : null
}

function buildReviewStateEntries(records: ReviewRecord[]): Array<{
  reviewerUsername: string
  review: ReviewState
}> {
  return records.map((record) => ({
    reviewerUsername: normalizeReviewerUsername(record.reviewerUsername),
    review: normalizeReview(record)
  }))
}

function resolveReviewTagValues(
  reviewStates: Array<{ reviewerUsername: string; review: ReviewState }> | undefined,
  reviewerUsername: string | undefined
): string[] {
  const review = resolveReviewState(reviewStates, reviewerUsername)
  return review ? buildReviewTagValues(review) : []
}

function sortDocuments(left: SearchIndexDocument, right: SearchIndexDocument, sort: string): number {
  if (sort === 'order') {
    if (left.scopeLabel !== right.scopeLabel) {
      return left.scopeLabel.localeCompare(right.scopeLabel, undefined, { sensitivity: 'base' })
    }
    if (left.fileName !== right.fileName) {
      return left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
    }
    if (left.folderPath !== right.folderPath) {
      return left.folderPath.localeCompare(right.folderPath, undefined, { sensitivity: 'base' })
    }
    return left.order - right.order
  }

  const leftDate = left.sortDateMs ?? Number.MIN_SAFE_INTEGER
  const rightDate = right.sortDateMs ?? Number.MIN_SAFE_INTEGER
  if (leftDate !== rightDate) {
    return rightDate - leftDate
  }
  if (left.scopeLabel !== right.scopeLabel) {
    return left.scopeLabel.localeCompare(right.scopeLabel, undefined, { sensitivity: 'base' })
  }
  if (left.fileName !== right.fileName) {
    return left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
  }
  if (left.folderPath !== right.folderPath) {
    return left.folderPath.localeCompare(right.folderPath, undefined, { sensitivity: 'base' })
  }
  return left.order - right.order
}

function buildSearchClause(term: { text: string; phrase: boolean }): Record<string, unknown> {
  const normalized = normalizeExactValue(term.text)
  if (!normalized) {
    return {}
  }

  const regex = new RegExp(escapeRegex(normalized), 'i')
  if (term.phrase || normalized.includes(' ')) {
    return {
      searchText: regex
    }
  }

  return {
    $or: [
      { searchTokens: normalized },
      { addressValues: normalized },
      { searchText: regex }
    ]
  }
}

function buildSearchExpression(query: string, mode: SearchMode): Record<string, unknown> {
  const parsed = parseSearchTerms(query, mode)
  if (!parsed.terms.length) {
    return {}
  }

  const clauses = parsed.terms
    .map((term) => buildSearchClause(term))
    .filter((clause) => Object.keys(clause).length > 0)
  if (!clauses.length) {
    return {}
  }

  if (parsed.mode === 'or') {
    return {
      $or: clauses
    }
  }

  return {
    $and: clauses
  }
}

function buildFilterMatch(
  options: SearchIndexSearchOptions,
  hiddenRules: HiddenRuleRecord[]
): Record<string, unknown> {
  const filter: Record<string, unknown> = {}
  const mailboxKeysProvided = options.allowedMailboxKeys !== undefined
  const allowedMailboxKeys = uniqueTextValues(options.allowedMailboxKeys || [])

  if (options.scope === 'pst') {
    if (options.mailboxKey) {
      allowedMailboxKeys.push(normalizeText(options.mailboxKey))
    }
  } else if (options.scope === 'search') {
    const scopePath = normalizeText(options.scopePath)
    if (scopePath) {
      filter.scopePath = scopePath
    }
  }

  if (options.mailOnly) {
    filter.isMailLike = true
  }

  if (options.sourceType && options.sourceType !== 'all') {
    filter.sourceType = normalizeSourceType(options.sourceType)
  }

  const reviewerUsername = normalizeReviewerUsername(options.reviewerUsername)
  const reviewClause: Record<string, unknown> = { reviewerUsername }
  if (options.reviewFlaggedOnly) {
    reviewClause['review.flagged'] = true
  }

  if (options.reviewTaggedOnly) {
    reviewClause['review.tags.0'] = { $exists: true }
  }

  if (options.reviewTag) {
    reviewClause['review.tags'] = {
      $elemMatch: {
        $regex: `^${escapeRegex(normalizeText(options.reviewTag))}$`,
        $options: 'i'
      }
    }
  }

  if (Object.keys(reviewClause).length > 1) {
    filter.reviewStates = {
      $elemMatch: reviewClause
    }
  }

  const normalizedMailboxKeys = uniqueTextValues(allowedMailboxKeys)
  if (mailboxKeysProvided) {
    filter.mailboxKey = { $in: normalizedMailboxKeys }
  }

  const hiddenAddresses = uniqueStrings(
    hiddenRules.filter((rule) => rule.kind === 'address').map((rule) => rule.value)
  )
  const hiddenSubjects = uniqueStrings(
    hiddenRules.filter((rule) => rule.kind === 'subject').map((rule) => rule.value)
  )

  if (hiddenAddresses.length) {
    filter.addressValues = { $nin: hiddenAddresses }
  }

  if (hiddenSubjects.length) {
    filter.subjectValues = { $nin: hiddenSubjects }
  }

  const searchExpression = buildSearchExpression(options.query, options.mode)
  if (Object.keys(searchExpression).length) {
    Object.assign(filter, searchExpression)
  }

  return filter
}

function createSortSpec(sort: string): Record<string, 1 | -1> {
  if (sort === 'order') {
    return {
      scopeLabel: 1,
      fileName: 1,
      folderPath: 1,
      order: 1,
      messageId: 1
    }
  }

  return {
    sortDateMs: -1,
    scopeLabel: 1,
    fileName: 1,
    folderPath: 1,
    order: 1,
    messageId: 1
  }
}

function buildReviewTagValues(review: ReviewState): string[] {
  return uniqueStrings(review.tags)
}

function resolveSearchIndexDocument(
  record: SearchIndexDocument,
  reviewerUsername: string | undefined
): SearchIndexDocument {
  const review = resolveReviewState(record.reviewStates, reviewerUsername)
  const normalizedReview = normalizeReview(review)
  return {
    ...record,
    review: normalizedReview,
    reviewTagValues: resolveReviewTagValues(record.reviewStates, reviewerUsername)
  }
}

function toDocument(
  base: Omit<
    SearchIndexDocument,
    | 'searchText'
    | 'searchTokens'
    | 'addressValues'
    | 'subjectValues'
    | 'updatedAt'
    | 'reviewTagValues'
    | 'bodySearchText'
    | 'review'
    | 'reviewStates'
  >,
  bodySearchText: string,
  reviewStates: Array<{
    reviewerUsername: string
    review: ReviewState
  }>
): SearchIndexDocument {
  const searchText = buildSearchText(base, bodySearchText)
  const subjectValues = uniqueStrings([base.subject, base.originalSubject])
  return {
    ...base,
    bodySearchText: normalizeExactValue(bodySearchText),
    searchText,
    searchTokens: tokenizeSearchText(searchText),
    addressValues: extractEmailAddresses(
      base.senderEmailAddress,
      base.displayTo,
      base.displayCC,
      base.displayBCC,
      base.resolvedDisplayTo,
      base.resolvedDisplayCC,
      base.resolvedDisplayBCC
    ),
    subjectValues,
    review: normalizeReview(null),
    reviewStates: reviewStates.map((entry) => ({
      reviewerUsername: normalizeReviewerUsername(entry.reviewerUsername),
      review: normalizeReview(entry.review)
    })),
    reviewTagValues: [],
    updatedAt: new Date().toISOString()
  }
}

export function buildSearchIndexDocumentsFromSession(
  session: ViewerSessionIndex,
  context: {
    mailboxKey: string
    scopePath: string
    scopeLabel: string
    fileName: string
    mailboxName: string
  },
  reviewRecords: ReviewRecord[]
): SearchIndexDocument[] {
  const reviewStatesByMessageId = new Map<
    string,
    Array<{ reviewerUsername: string; review: ReviewState }>
  >()

  for (const record of reviewRecords) {
    const messageId = normalizeText(record.messageId)
    if (!messageId) {
      continue
    }
    const entries = reviewStatesByMessageId.get(messageId) || []
    entries.push({
      reviewerUsername: normalizeReviewerUsername(record.reviewerUsername),
      review: normalizeReview(record)
    })
    reviewStatesByMessageId.set(messageId, entries)
  }

  const documents: SearchIndexDocument[] = []

  for (const message of session.messages.values()) {
    const bodySearchText = normalizeExactValue(session.searchTextByMessageId.get(message.id) || '')
    documents.push(
      toDocument(
        {
          id: normalizeText(message.id),
          sourceType: 'mailbox',
          mailboxKey: normalizeText(context.mailboxKey),
          scopePath: normalizeText(context.scopePath),
          scopeLabel: normalizeText(context.scopeLabel),
          fileName: normalizeText(context.fileName),
          mailboxName: normalizeText(context.mailboxName),
          messageId: normalizeText(message.id),
          descriptorId: normalizeText(message.descriptorId),
          folderId: normalizeText(message.folderId),
          folderPath: normalizeText(message.folderPath),
          order: message.order,
          messageClass: normalizeText(message.messageClass),
          kind: message.kind,
          subject: normalizeText(message.subject),
          originalSubject: normalizeText(message.originalSubject),
          senderName: normalizeText(message.senderName),
          senderEmailAddress: normalizeText(message.senderEmailAddress),
          recipientText: normalizeText(message.recipientText),
          displayTo: normalizeText(message.displayTo),
          displayCC: normalizeText(message.displayCC),
          displayBCC: normalizeText(message.displayBCC),
          resolvedDisplayTo: normalizeText(message.resolvedDisplayTo),
          resolvedDisplayCC: normalizeText(message.resolvedDisplayCC),
          resolvedDisplayBCC: normalizeText(message.resolvedDisplayBCC),
          clientSubmitTime: message.clientSubmitTime,
          creationTime: message.creationTime,
          modificationTime: message.modificationTime,
          messageDeliveryTime: message.messageDeliveryTime,
          sortDate: message.sortDate,
          sortDateMs: message.sortDateMs,
          importance: message.importance,
          hasAttachments: message.hasAttachments,
          isRead: message.isRead,
          isMailLike: message.isMailLike
        },
        bodySearchText,
        reviewStatesByMessageId.get(message.id) || []
      )
    )
  }

  return dedupeSearchIndexDocuments(documents)
}

export function buildSearchIndexDocumentsFromArchiveItems(
  items: ArchiveBundleItem[]
): SearchIndexDocument[] {
  const documents = items.map((item) =>
    toDocument(
      {
        id: normalizeText(item.archiveItemId),
        sourceType: normalizeSourceType(item.sourceType),
        mailboxKey: normalizeText(item.bundlePath),
        scopePath: normalizeText(item.scopePath),
        scopeLabel: normalizeText(item.scopeLabel),
        fileName: normalizeText(item.bundleFileName),
        mailboxName: normalizeText(item.bundleFileName),
        messageId: normalizeText(item.archiveItemId),
        descriptorId: normalizeText(item.archiveItemId),
        folderId: normalizeText(item.entryPath),
        folderPath: normalizeText(path.dirname(item.entryPath)),
        order: 0,
        messageClass: item.sourceType === 'teams' ? 'IPM.Note' : 'IPM.Document',
        kind: 'other',
        subject: normalizeText(item.entryName),
        originalSubject: normalizeText(item.entryName),
        senderName: item.sourceType === 'teams' ? 'Teams' : 'SharePoint/OneDrive',
        senderEmailAddress: '',
        recipientText: item.entryPath,
        displayTo: '',
        displayCC: '',
        displayBCC: '',
        resolvedDisplayTo: '',
        resolvedDisplayCC: '',
        resolvedDisplayBCC: '',
        clientSubmitTime: item.modifiedAt,
        creationTime: item.modifiedAt,
        modificationTime: item.modifiedAt,
        messageDeliveryTime: item.modifiedAt,
        sortDate: item.modifiedAt,
        sortDateMs: item.modifiedAt ? Date.parse(item.modifiedAt) : null,
        importance: 0,
        hasAttachments: false,
        isRead: true,
        isMailLike: false,
        archivePath: normalizeText(item.bundlePath),
        archiveEntryPath: normalizeText(item.entryPath),
        archiveEntryChain: [...item.entryChain],
        archiveEntryName: normalizeText(item.entryName),
        contentType: normalizeText(item.contentType),
        downloadFilename: normalizeText(item.downloadFilename),
        previewKind: item.previewKind,
        previewText: item.previewText,
        previewHtml: item.previewHtml
      },
      item.searchText,
      []
    )
  )

  return dedupeSearchIndexDocuments(documents)
}

class MemoryHiddenRuleStore {
  private readonly rules = new Map<string, HiddenRuleRecord>()

  listHiddenRules(): HiddenRuleRecord[] {
    return [...this.rules.values()].sort((left, right) => {
      const leftTime = Date.parse(left.updatedAt || left.createdAt || '')
      const rightTime = Date.parse(right.updatedAt || right.createdAt || '')
      if (rightTime !== leftTime) {
        return rightTime - leftTime
      }
      return left.label.localeCompare(right.label, undefined, { sensitivity: 'base' })
    })
  }

  upsertHiddenRule(input: { kind: HiddenRuleKind; value: string; label?: string }): HiddenRuleRecord {
    const value = normalizeExactValue(input.value)
    if (!value) {
      throw new Error('Filter value is required')
    }
    const kind = input.kind
    const key = `${kind}::${value}`
    const now = new Date().toISOString()
    const existing = this.rules.get(key)
    const record: HiddenRuleRecord = {
      filterId: existing?.filterId || randomBytes(8).toString('hex'),
      kind,
      value,
      label: normalizeText(input.label || value),
      createdAt: existing?.createdAt || now,
      updatedAt: now
    }
    this.rules.set(key, record)
    return record
  }

  deleteHiddenRule(filterId: string): boolean {
    for (const [key, record] of this.rules.entries()) {
      if (record.filterId === filterId) {
        this.rules.delete(key)
        return true
      }
    }
    return false
  }

  clear(): void {
    this.rules.clear()
  }
}

export class MemorySearchIndexStore implements SearchIndexStore {
  public kind: 'memory' = 'memory'
  public isPersistent = false
  private readonly documents = new Map<string, Map<string, SearchIndexDocument>>()
  private readonly hiddenRules = new MemoryHiddenRuleStore()

  async replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void> {
    const key = normalizeText(mailboxKey)
    const records = new Map<string, SearchIndexDocument>()
    for (const document of dedupeSearchIndexDocuments(documents)) {
      records.set(document.messageId, {
        ...document,
        mailboxKey: key,
        sourceType: normalizeSourceType(document.sourceType)
      })
    }
    this.documents.set(key, records)
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    this.documents.delete(normalizeText(mailboxKey))
  }

  async updateReviewState(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string,
    review: ReviewState | null
  ): Promise<void> {
    const key = normalizeText(mailboxKey)
    const mailbox = this.documents.get(key)
    if (!mailbox) {
      return
    }
    const record = mailbox.get(normalizeText(messageId))
    if (!record) {
      return
    }
    const normalizedReviewerUsername = normalizeReviewerUsername(reviewerUsername)
    const normalizedReview = normalizeReview(review)
    const reviewStates = [...(record.reviewStates || [])].filter(
      (entry) => normalizeReviewerUsername(entry.reviewerUsername) !== normalizedReviewerUsername
    )
    if (normalizedReview.flagged || normalizedReview.tags.length > 0) {
      reviewStates.push({
        reviewerUsername: normalizedReviewerUsername,
        review: normalizedReview
      })
    }
    mailbox.set(record.messageId, {
      ...record,
      reviewStates,
      review: normalizedReview,
      reviewTagValues: buildReviewTagValues(normalizedReview),
      updatedAt: new Date().toISOString()
    })
  }

  async clearAllDocuments(): Promise<void> {
    this.documents.clear()
  }

  async listHiddenRules(): Promise<HiddenRuleRecord[]> {
    return this.hiddenRules.listHiddenRules()
  }

  async upsertHiddenRule(input: { kind: HiddenRuleKind; value: string; label?: string }): Promise<HiddenRuleRecord> {
    return this.hiddenRules.upsertHiddenRule(input)
  }

  async deleteHiddenRule(filterId: string): Promise<boolean> {
    return this.hiddenRules.deleteHiddenRule(filterId)
  }

  async findDocumentById(id: string): Promise<SearchIndexDocument | null> {
    const normalizedId = normalizeText(id)
    if (!normalizedId) {
      return null
    }
    for (const mailbox of this.documents.values()) {
      for (const record of mailbox.values()) {
        if (normalizeText(record.id || record.messageId) === normalizedId || normalizeText(record.messageId) === normalizedId) {
          return record
        }
      }
    }
    return null
  }

  async search(options: SearchIndexSearchOptions): Promise<SearchIndexPage> {
    const hiddenRules = this.hiddenRules.listHiddenRules()
    const mailboxKeysProvided = options.allowedMailboxKeys !== undefined
    const allowedMailboxKeys = uniqueTextValues(options.allowedMailboxKeys || [])
    const allRecords = [...this.documents.values()].flatMap((mailbox) => [...mailbox.values()])
    const countOptions = {
      ...options,
      sourceType: 'all' as const
    }
    const sourceCounts = allRecords
      .filter((record) => (mailboxKeysProvided ? allowedMailboxKeys.includes(record.mailboxKey) : true))
      .filter((record) => matchesDocument(record, countOptions, hiddenRules))
      .reduce<Record<SearchSourceType, number>>(
        (acc, record) => {
          const sourceType = normalizeSourceType(record.sourceType)
          acc[sourceType] = (acc[sourceType] || 0) + 1
          return acc
        },
        { mailbox: 0, teams: 0, sharepoint: 0 }
      )
    const records = mailboxKeysProvided
      ? allRecords.filter((record) => allowedMailboxKeys.includes(record.mailboxKey))
      : allRecords
    const matched = records
      .filter((record) => matchesDocument(record, options, hiddenRules))
      .sort((left, right) => sortDocuments(left, right, options.sort))
    const resolved = matched.map((record) =>
      resolveSearchIndexDocument(record, options.reviewerUsername)
    )
    return paginateSearchResults(resolved, options, hiddenRules, sourceCounts)
  }

  async close(): Promise<void> {
    this.documents.clear()
    this.hiddenRules.clear()
  }
}

function matchesDocument(
  record: SearchIndexDocument,
  options: SearchIndexSearchOptions,
  hiddenRules: HiddenRuleRecord[]
): boolean {
  if (options.sourceType && options.sourceType !== 'all') {
    if (normalizeSourceType(record.sourceType) !== normalizeSourceType(options.sourceType)) {
      return false
    }
  }

  if (options.scope === 'pst') {
    if (options.mailboxKey && record.mailboxKey !== normalizeText(options.mailboxKey)) {
      return false
    }
  } else if (options.scope === 'search') {
    if (options.scopePath && record.scopePath !== normalizeText(options.scopePath)) {
      return false
    }
  }

  if (options.mailOnly && !record.isMailLike) {
    return false
  }

  const review = resolveReviewState(record.reviewStates, options.reviewerUsername)
  if (options.reviewFlaggedOnly && !review?.flagged) {
    return false
  }
  if (options.reviewTaggedOnly && (!review || review.tags.length === 0)) {
    return false
  }
  if (options.reviewTag) {
    const needle = normalizeExactValue(options.reviewTag)
    if (!review || !review.tags.some((tag) => normalizeExactValue(tag) === needle)) {
      return false
    }
  }

  const hiddenAddresses = uniqueStrings(
    hiddenRules
      .filter((rule) => rule.kind === 'address')
      .flatMap((rule) => buildAddressMatchValues(rule.value))
  )
  const hiddenSubjects = uniqueStrings(
    hiddenRules.filter((rule) => rule.kind === 'subject').map((rule) => rule.value)
  )

  if (hiddenAddresses.length) {
    const addressValues = buildAddressMatchValues(
      record.senderEmailAddress,
      record.displayTo,
      record.displayCC,
      record.displayBCC,
      record.resolvedDisplayTo,
      record.resolvedDisplayCC,
      record.resolvedDisplayBCC,
      record.recipientText,
      ...record.addressValues
    )
    if (addressValues.some((value) => hiddenAddresses.includes(value))) {
      return false
    }
  }

  if (
    hiddenSubjects.length &&
    record.subjectValues.some((value) => hiddenSubjects.includes(value))
  ) {
    return false
  }

  if (!matchesSearchExpression(record, options.query, options.mode)) {
    return false
  }

  return true
}

function matchesSearchExpression(
  record: SearchIndexDocument,
  query: string,
  mode: SearchMode
): boolean {
  const parsed = parseSearchTerms(query, mode)
  if (!parsed.terms.length) {
    return true
  }

  const haystack = record.searchText
  const tokens = new Set(record.searchTokens)
  const addresses = new Set(record.addressValues)
  const matchesTerm = (term: { text: string; phrase: boolean }): boolean => {
    const normalized = normalizeExactValue(term.text)
    if (!normalized) {
      return false
    }
    const regexMatch = haystack.includes(normalized)
    if (term.phrase || normalized.includes(' ')) {
      return regexMatch
    }
    return tokens.has(normalized) || addresses.has(normalized) || regexMatch
  }

  return parsed.mode === 'or'
    ? parsed.terms.some((term) => matchesTerm(term))
    : parsed.terms.every((term) => matchesTerm(term))
}

function paginateSearchResults(
  records: SearchIndexDocument[],
  options: SearchIndexSearchOptions,
  hiddenRules: HiddenRuleRecord[],
  sourceCounts?: Record<SearchSourceType, number>
): SearchIndexPage {
  const total = records.length
  const totalPages = Math.max(1, Math.ceil(total / options.pageSize))
  const page = Math.min(Math.max(options.page, 1), totalPages)
  const start = (page - 1) * options.pageSize
  const parsed = parseSearchTerms(options.query, options.mode)
  const scopeLabel =
    options.scope === 'all'
      ? 'All cases/searches'
      : options.scope === 'search'
        ? normalizeText(options.scopePath ? options.scopePath.split('/').join(' / ') : '')
        : 'Selected PST'
  const scopePath =
    options.scope === 'search'
      ? normalizeText(options.scopePath)
      : options.scope === 'pst'
        ? ''
        : ''

  const resolvedSourceCounts =
    sourceCounts ||
    records.reduce<Record<SearchSourceType, number>>(
      (acc, record) => {
        const sourceType = normalizeSourceType(record.sourceType)
        acc[sourceType] = (acc[sourceType] || 0) + 1
        return acc
      },
      { mailbox: 0, teams: 0, sharepoint: 0 }
    )

  return {
    items: records.slice(start, start + options.pageSize),
    total,
    page,
    pageSize: options.pageSize,
    totalPages,
    query: normalizeText(options.query),
    mode: parsed.mode,
    mailOnly: options.mailOnly,
    sort: options.sort,
    scope: options.scope,
    scopePath,
    scopeLabel,
    hiddenRules,
    sourceType: options.sourceType || 'all',
    sourceCounts: resolvedSourceCounts,
    reviewFilters: {
      flaggedOnly: options.reviewFlaggedOnly,
      taggedOnly: options.reviewTaggedOnly,
      tag: normalizeText(options.reviewTag)
    }
  }
}

export class MongoSearchIndexStore implements SearchIndexStore {
  public kind: 'mongo' = 'mongo'
  public isPersistent = true

  constructor(
    private readonly documents: SearchIndexCollectionLike,
    private readonly rules: HiddenRuleCollectionLike,
    private readonly client?: MongoClient,
    private readonly dbName = 'pst-extractor',
    private readonly documentsCollectionName = DEFAULT_INDEX_COLLECTION,
    private readonly rulesCollectionName = DEFAULT_RULE_COLLECTION
  ) {}

  static async connect(
    uri: string,
    dbName = 'pst-extractor',
    options: MongoSearchIndexStoreConnectOptions = {}
  ): Promise<MongoSearchIndexStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const db = client.db(dbName)
    const documents = db.collection<SearchIndexDocument>(
      options.documentsCollectionName || DEFAULT_INDEX_COLLECTION
    )
    const rules = db.collection<HiddenRuleRecord>(options.rulesCollectionName || DEFAULT_RULE_COLLECTION)
    await documents.createIndex?.({ mailboxKey: 1, messageId: 1 }, { unique: true })
    await documents.createIndex?.({ messageId: 1 })
    await documents.createIndex?.({ sourceType: 1, scopePath: 1 })
    await documents.createIndex?.({ mailboxKey: 1, scopePath: 1 })
    await documents.createIndex?.({ scopePath: 1, searchTokens: 1 })
    await documents.createIndex?.({ scopePath: 1, addressValues: 1 })
    await documents.createIndex?.({ scopePath: 1, subjectValues: 1 })
    await documents.createIndex?.({ mailboxKey: 1, 'reviewStates.reviewerUsername': 1 })
    await documents.createIndex?.({ mailboxKey: 1, 'reviewStates.review.flagged': 1 })
    await documents.createIndex?.({ mailboxKey: 1, 'reviewStates.review.tags': 1 })
    await documents.createIndex?.({ mailboxKey: 1, sortDateMs: -1 })
    await rules.createIndex?.({ kind: 1, value: 1 }, { unique: true })
    await rules.createIndex?.({ updatedAt: -1 })
    return new MongoSearchIndexStore(
      documents as unknown as SearchIndexCollectionLike,
      rules as unknown as HiddenRuleCollectionLike,
      client,
      dbName,
      options.documentsCollectionName || DEFAULT_INDEX_COLLECTION,
      options.rulesCollectionName || DEFAULT_RULE_COLLECTION
    )
  }

  async replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void> {
    const key = normalizeText(mailboxKey)
    await this.documents.deleteMany({ mailboxKey: key })
    const uniqueDocuments = dedupeSearchIndexDocuments(documents)
    if (!uniqueDocuments.length) {
      return
    }
    await this.documents.insertMany(
      uniqueDocuments.map((document) => ({
        ...document,
        mailboxKey: key,
        sourceType: normalizeSourceType(document.sourceType)
      }))
    )
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    await this.documents.deleteMany({ mailboxKey: normalizeText(mailboxKey) })
  }

  async findDocumentById(id: string): Promise<SearchIndexDocument | null> {
    const normalizedId = normalizeText(id)
    if (!normalizedId) {
      return null
    }
    return this.documents.findOne({
      $or: [{ messageId: normalizedId }, { id: normalizedId }]
    })
  }

  async updateReviewState(
    mailboxKey: string,
    messageId: string,
    reviewerUsername: string,
    review: ReviewState | null
  ): Promise<void> {
    const normalizedMailboxKey = normalizeText(mailboxKey)
    const normalizedMessageId = normalizeText(messageId)
    const normalizedReviewerUsername = normalizeReviewerUsername(reviewerUsername)
    const normalizedReview = normalizeReview(review)
    const existing = await this.documents.findOne({
      mailboxKey: normalizedMailboxKey,
      messageId: normalizedMessageId
    })
    if (!existing) {
      return
    }
    const existingStates = Array.isArray((existing as SearchIndexDocument).reviewStates)
      ? [...(existing as SearchIndexDocument).reviewStates]
      : []
    const reviewStates = existingStates.filter(
      (entry) => normalizeReviewerUsername(entry.reviewerUsername) !== normalizedReviewerUsername
    )
    if (normalizedReview.flagged || normalizedReview.tags.length > 0) {
      reviewStates.push({
        reviewerUsername: normalizedReviewerUsername,
        review: normalizedReview
      })
    }
    await this.documents.updateOne(
      {
        mailboxKey: normalizedMailboxKey,
        messageId: normalizedMessageId
      },
      {
        $set: {
          reviewStates,
          review: normalizedReview,
          reviewTagValues: buildReviewTagValues(normalizedReview),
          updatedAt: new Date().toISOString()
        }
      }
    )
  }

  async clearAllDocuments(): Promise<void> {
    await this.documents.deleteMany({})
  }

  async listHiddenRules(): Promise<HiddenRuleRecord[]> {
    return this.rules.find({}).sort({ updatedAt: -1 }).toArray()
  }

  async upsertHiddenRule(input: { kind: HiddenRuleKind; value: string; label?: string }): Promise<HiddenRuleRecord> {
    const value = normalizeExactValue(input.value)
    if (!value) {
      throw new Error('Filter value is required')
    }
    const kind = input.kind
    const now = new Date().toISOString()
    const existing = await this.rules.findOne({ kind, value })
    const record: HiddenRuleRecord = {
      filterId: existing?.filterId || randomBytes(8).toString('hex'),
      kind,
      value,
      label: normalizeText(input.label || value),
      createdAt: existing?.createdAt || now,
      updatedAt: now
    }
    await this.rules.updateOne({ kind, value }, { $set: record }, { upsert: true })
    return record
  }

  async deleteHiddenRule(filterId: string): Promise<boolean> {
    const result = await this.rules.deleteOne({ filterId: normalizeText(filterId) })
    return Boolean(result.deletedCount && result.deletedCount > 0)
  }

  async promoteStagedDocuments(stagingDocumentsCollectionName: string): Promise<void> {
    if (!this.client) {
      throw new Error('Mongo client is not available')
    }

    const stagingName = normalizeText(stagingDocumentsCollectionName)
    if (!stagingName || stagingName === this.documentsCollectionName) {
      return
    }

    await this.client.db('admin').command({
      renameCollection: `${this.dbName}.${stagingName}`,
      to: `${this.dbName}.${this.documentsCollectionName}`,
      dropTarget: true
    })
  }

  async search(options: SearchIndexSearchOptions): Promise<SearchIndexPage> {
    const hiddenRules = await this.listHiddenRules()
    const filter = buildFilterMatch(options, hiddenRules)
    const countFilter = buildFilterMatch(
      {
        ...options,
        sourceType: 'all'
      },
      hiddenRules
    )
    const total = await this.documents.countDocuments(filter)
    const totalPages = Math.max(1, Math.ceil(total / options.pageSize))
    const page = Math.min(Math.max(options.page, 1), totalPages)
    const start = (page - 1) * options.pageSize
    const items = await this.documents
      .find(filter)
      .sort(createSortSpec(options.sort))
      .skip(start)
      .limit(options.pageSize)
      .toArray()
    const resolvedItems = items.map((record) => resolveSearchIndexDocument(record, options.reviewerUsername))
    const parsed = parseSearchTerms(options.query, options.mode)
    const sourceCounts = {
      mailbox: await this.documents.countDocuments({ ...countFilter, sourceType: 'mailbox' }),
      teams: await this.documents.countDocuments({ ...countFilter, sourceType: 'teams' }),
      sharepoint: await this.documents.countDocuments({ ...countFilter, sourceType: 'sharepoint' })
    }
    return {
      items: resolvedItems,
      total,
      page,
      pageSize: options.pageSize,
      totalPages,
      query: normalizeText(options.query),
      mode: parsed.mode,
      mailOnly: options.mailOnly,
      sort: options.sort,
      scope: options.scope,
      scopePath:
        options.scope === 'search'
          ? normalizeText(options.scopePath)
          : '',
      scopeLabel:
        options.scope === 'all'
          ? 'All cases/searches'
          : options.scope === 'search'
            ? normalizeText(options.scopePath ? options.scopePath.split('/').join(' / ') : '')
            : 'Selected PST',
      hiddenRules,
      sourceType: options.sourceType || 'all',
      sourceCounts,
      reviewFilters: {
        flaggedOnly: options.reviewFlaggedOnly,
        taggedOnly: options.reviewTaggedOnly,
        tag: normalizeText(options.reviewTag)
      }
    }
  }

  async close(): Promise<void> {
    if (this.client) {
      await this.client.close()
    }
  }
}

export async function createSearchIndexStoreFromEnv(
  env: NodeJS.ProcessEnv = process.env,
  options: MongoSearchIndexStoreConnectOptions = {}
): Promise<SearchIndexStore> {
  const uri = normalizeText(env.MONGODB_URI)
  if (!uri) {
    return new MemorySearchIndexStore()
  }
  const dbName = normalizeText(env.MONGODB_DB) || 'pst-extractor'
  return MongoSearchIndexStore.connect(uri, dbName, options)
}

export async function refreshSearchIndexFromCatalog(
  rootPath: string,
  reviewStore: ReviewStore,
  searchIndexStore: SearchIndexStore
): Promise<{
  mailboxCount: number
  messageCount: number
}> {
  const { listPstMailboxFiles, openPstMailbox } = await import('./pstCatalog')
  const { listArchiveBundleFiles, extractArchiveBundleItems } = await import('./archiveBundles')
  const catalog = listPstMailboxFiles(rootPath)
  const archiveCatalog = listArchiveBundleFiles(rootPath)
  await searchIndexStore.clearAllDocuments()

  let mailboxCount = 0
  let messageCount = 0
  const warnedMessages = new Set<string>()

  function warnOnce(scopeLabel: string, fileName: string, error: unknown): void {
    const message = error instanceof Error ? error.message : String(error)
    const key = `${scopeLabel}/${fileName}::${message}`
    if (warnedMessages.has(key)) {
      return
    }
    warnedMessages.add(key)
    console.warn(`Unable to refresh search index for ${scopeLabel}/${fileName}:`, message)
  }

  for (const scope of catalog.scopes) {
    for (const file of scope.files) {
      try {
        const session = openPstMailbox(rootPath, scope.scopePath, file.fileName)
        const mailboxKey = session.filePath
        const reviewRecords = await reviewStore.listReviews(mailboxKey)
        const documents = buildSearchIndexDocumentsFromSession(
          session,
          {
            mailboxKey,
            scopePath: scope.scopePath,
            scopeLabel: scope.scopeLabel,
            fileName: file.fileName,
            mailboxName: session.mailboxName
          },
          reviewRecords
        )
        await searchIndexStore.replaceMailboxDocuments(mailboxKey, documents)
        mailboxCount += 1
        messageCount += documents.length
      } catch (error) {
        warnOnce(scope.scopeLabel, file.fileName, error)
      }
    }
  }

  for (const scope of archiveCatalog.scopes) {
    for (const file of scope.files) {
      try {
        const bundlePath = path.resolve(rootPath, scope.scopePath ? scope.scopePath : '', file.fileName)
        const archiveItems = await extractArchiveBundleItems(bundlePath, scope.scopePath, file.fileName)
        const documents = buildSearchIndexDocumentsFromArchiveItems(archiveItems)
        await searchIndexStore.replaceMailboxDocuments(bundlePath, documents)
        mailboxCount += 1
        messageCount += documents.length
      } catch (error) {
        warnOnce(scope.scopeLabel, file.fileName, error)
      }
    }
  }

  return {
    mailboxCount,
    messageCount
  }
}
