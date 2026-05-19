import { randomBytes } from 'crypto'
import { MongoClient } from 'mongodb'
import type { ReviewStore } from './reviewStore'
import type { ReviewState } from './reviewTypes'
import type { MessageSummary, ViewerSessionIndex } from './viewer'

export type SearchScope = 'all' | 'search' | 'pst'
export type SearchMode = 'and' | 'or'
export type HiddenRuleKind = 'address' | 'subject'

export interface HiddenRuleRecord {
  filterId: string
  kind: HiddenRuleKind
  value: string
  label: string
  createdAt: string
  updatedAt: string
}

export interface SearchIndexDocument {
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
  review: ReviewState
  reviewTagValues: string[]
  updatedAt: string
}

export interface SearchIndexSearchOptions {
  scope: SearchScope
  scopePath?: string
  mailboxKey?: string
  allowedMailboxKeys?: string[]
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
  updateReviewState(mailboxKey: string, messageId: string, review: ReviewState | null): Promise<void>
  clearAllDocuments(): Promise<void>
  listHiddenRules(): Promise<HiddenRuleRecord[]>
  upsertHiddenRule(input: {
    kind: HiddenRuleKind
    value: string
    label?: string
  }): Promise<HiddenRuleRecord>
  deleteHiddenRule(filterId: string): Promise<boolean>
  search(options: SearchIndexSearchOptions): Promise<SearchIndexPage>
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

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeExactValue(value: unknown): string {
  return normalizeText(value).toLowerCase()
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
    bodySearchText
  ]
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

  if (options.reviewFlaggedOnly) {
    filter['review.flagged'] = true
  }

  if (options.reviewTaggedOnly) {
    filter['review.tags.0'] = { $exists: true }
  }

  if (options.reviewTag) {
    filter.reviewTagValues = normalizeExactValue(options.reviewTag)
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
  >,
  bodySearchText: string,
  review: ReviewState | null
): SearchIndexDocument {
  const normalizedReview = normalizeReview(review)
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
    review: normalizedReview,
    reviewTagValues: buildReviewTagValues(normalizedReview),
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
  reviewMap: Map<string, ReviewState>
): SearchIndexDocument[] {
  const documents: SearchIndexDocument[] = []

  for (const message of session.messages.values()) {
    const review = reviewMap.get(message.id) || null
    const bodySearchText = normalizeExactValue(session.searchTextByMessageId.get(message.id) || '')
    documents.push(
      toDocument(
        {
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
        review
      )
    )
  }

  return documents
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
    for (const document of documents) {
      records.set(document.messageId, { ...document, mailboxKey: key })
    }
    this.documents.set(key, records)
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    this.documents.delete(normalizeText(mailboxKey))
  }

  async updateReviewState(
    mailboxKey: string,
    messageId: string,
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
    const normalizedReview = normalizeReview(review)
    mailbox.set(record.messageId, {
      ...record,
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

  async search(options: SearchIndexSearchOptions): Promise<SearchIndexPage> {
    const hiddenRules = this.hiddenRules.listHiddenRules()
    const mailboxKeysProvided = options.allowedMailboxKeys !== undefined
    const allowedMailboxKeys = uniqueTextValues(options.allowedMailboxKeys || [])
    const allRecords = [...this.documents.values()].flatMap((mailbox) => [...mailbox.values()])
    const records = mailboxKeysProvided
      ? allRecords.filter((record) => allowedMailboxKeys.includes(record.mailboxKey))
      : allRecords
    const matched = records
      .filter((record) => matchesDocument(record, options, hiddenRules))
      .sort((left, right) => sortDocuments(left, right, options.sort))
    return paginateSearchResults(matched, options, hiddenRules)
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

  if (options.reviewFlaggedOnly && !record.review.flagged) {
    return false
  }
  if (options.reviewTaggedOnly && record.review.tags.length === 0) {
    return false
  }
  if (options.reviewTag) {
    const needle = normalizeExactValue(options.reviewTag)
    if (!record.reviewTagValues.includes(needle)) {
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
  hiddenRules: HiddenRuleRecord[]
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
    private readonly client?: MongoClient
  ) {}

  static async connect(uri: string, dbName = 'pst-extractor'): Promise<MongoSearchIndexStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const db = client.db(dbName)
    const documents = db.collection<SearchIndexDocument>(DEFAULT_INDEX_COLLECTION)
    const rules = db.collection<HiddenRuleRecord>(DEFAULT_RULE_COLLECTION)
    await documents.createIndex?.({ mailboxKey: 1, messageId: 1 }, { unique: true })
    await documents.createIndex?.({ mailboxKey: 1, scopePath: 1 })
    await documents.createIndex?.({ scopePath: 1, searchTokens: 1 })
    await documents.createIndex?.({ scopePath: 1, addressValues: 1 })
    await documents.createIndex?.({ scopePath: 1, subjectValues: 1 })
    await documents.createIndex?.({ mailboxKey: 1, 'review.flagged': 1 })
    await documents.createIndex?.({ mailboxKey: 1, reviewTagValues: 1 })
    await documents.createIndex?.({ mailboxKey: 1, sortDateMs: -1 })
    await rules.createIndex?.({ kind: 1, value: 1 }, { unique: true })
    await rules.createIndex?.({ updatedAt: -1 })
    return new MongoSearchIndexStore(
      documents as unknown as SearchIndexCollectionLike,
      rules as unknown as HiddenRuleCollectionLike,
      client
    )
  }

  async replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void> {
    const key = normalizeText(mailboxKey)
    await this.documents.deleteMany({ mailboxKey: key })
    if (!documents.length) {
      return
    }
    await this.documents.insertMany(documents.map((document) => ({ ...document, mailboxKey: key })))
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    await this.documents.deleteMany({ mailboxKey: normalizeText(mailboxKey) })
  }

  async updateReviewState(
    mailboxKey: string,
    messageId: string,
    review: ReviewState | null
  ): Promise<void> {
    const normalizedMailboxKey = normalizeText(mailboxKey)
    const normalizedMessageId = normalizeText(messageId)
    const normalizedReview = normalizeReview(review)
    await this.documents.updateOne(
      {
        mailboxKey: normalizedMailboxKey,
        messageId: normalizedMessageId
      },
      {
        $set: {
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

  async search(options: SearchIndexSearchOptions): Promise<SearchIndexPage> {
    const hiddenRules = await this.listHiddenRules()
    const filter = buildFilterMatch(options, hiddenRules)
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
    const parsed = parseSearchTerms(options.query, options.mode)
    return {
      items,
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
  env: NodeJS.ProcessEnv = process.env
): Promise<SearchIndexStore> {
  const uri = normalizeText(env.MONGODB_URI)
  if (!uri) {
    return new MemorySearchIndexStore()
  }
  const dbName = normalizeText(env.MONGODB_DB) || 'pst-extractor'
  return MongoSearchIndexStore.connect(uri, dbName)
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
  const catalog = listPstMailboxFiles(rootPath)
  await searchIndexStore.clearAllDocuments()

  let mailboxCount = 0
  let messageCount = 0

  for (const scope of catalog.scopes) {
    for (const file of scope.files) {
      try {
        const session = openPstMailbox(rootPath, scope.scopePath, file.fileName)
        const mailboxKey = session.filePath
        const messageIds = [...session.messages.keys()]
        const reviewMap = await reviewStore.getMany(mailboxKey, messageIds)
        const documents = buildSearchIndexDocumentsFromSession(
          session,
          {
            mailboxKey,
            scopePath: scope.scopePath,
            scopeLabel: scope.scopeLabel,
            fileName: file.fileName,
            mailboxName: session.mailboxName
          },
          reviewMap
        )
        await searchIndexStore.replaceMailboxDocuments(mailboxKey, documents)
        mailboxCount += 1
        messageCount += documents.length
      } catch (error) {
        console.warn(
          `Unable to refresh search index for ${scope.scopeLabel}/${file.fileName}:`,
          error instanceof Error ? error.message : error
        )
      }
    }
  }

  return {
    mailboxCount,
    messageCount
  }
}
