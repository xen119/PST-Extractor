import { createHash, randomBytes } from 'crypto'
import * as path from 'path'
import { MongoClient } from 'mongodb'
import type { ReviewStore } from './reviewStore'
import type { ReviewRecord, ReviewState } from './reviewTypes'
import {
  cloneMessageDetail,
  streamPstMailboxMessages,
  type MessageDetail,
  type MessageSummary,
  type ViewerIndexedMessage,
  type ViewerSessionIndex
} from './viewer'
import type { ArchiveBundleItem, ArchiveBundleSourceType } from './archiveBundles'

export type SearchScope = 'all' | 'search' | 'pst'
export type SearchMode = 'and' | 'or'
export type HiddenRuleKind = 'address' | 'subject'
export type SearchSourceType = 'mailbox' | ArchiveBundleSourceType
export type SearchIndexRefreshSource = 'mailboxes' | 'items'

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
  threadMetadata?: MailboxThreadMetadata
  threadCollapse?: SearchThreadCollapseReference[]
  threadCollapsePartitions?: string[]
  threadCollapseVersion?: number
  threadInfo?: SearchThreadInfo
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
  size: number
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
  mailboxDetail?: MessageDetail
  review: ReviewState
  reviewStates: Array<{
    reviewerUsername: string
    review: ReviewState
  }>
  reviewTagValues: string[]
  updatedAt: string
}

export interface SearchIndexFileFingerprint {
  source: SearchIndexRefreshSource
  mailboxKey: string
  fileName: string
  scopePath: string
  scopeLabel: string
  size: number
  modifiedAt: string | null
  indexVersion?: number
  updatedAt: string
}

export interface MailboxThreadMetadata {
  messageId: string
  inReplyToId: string
  referenceIds: string[]
  conversationId: string
  isForward: boolean
}

export interface SearchThreadInfo {
  threadId: string
  branchId: string
  branchIndex: number
  branchCount: number
  threadItemCount: number
  branchItemCount: number
  isRepresentative: boolean
}

export interface SearchThreadCollapseReference extends SearchThreadInfo {
  partitionKey: string
  representativeMailboxKey: string
  representativeMessageId: string
}

export type MailboxCollapseJobState = 'idle' | 'running' | 'succeeded' | 'failed' | 'reindex-required'

export interface MailboxCollapseJobStatus {
  jobId: string | null
  status: MailboxCollapseJobState
  version: number
  startedAt: string | null
  completedAt: string | null
  updatedAt: string
  processedPartitions: number
  totalPartitions: number
  completedPartitionKeys: string[]
  processedWorkUnits: number
  totalWorkUnits: number
  percentage: number
  provisional: boolean
  error: string | null
  reindexRequired: boolean
}

export interface SearchThreadBranch {
  branchId: string
  branchIndex: number
  branchCount: number
  representativeId: string
  items: SearchIndexDocument[]
}

export interface SearchThreadGroup {
  threadId: string
  selectedItemId: string
  branches: SearchThreadBranch[]
}

export interface SearchIndexRefreshPlan {
  source: SearchIndexRefreshSource
  mailboxCount: number
  messageCount: number
  changedCount: number
  skippedCount: number
  removedCount: number
  failedCount: number
  changedMailboxKeys: string[]
  removedMailboxKeys: string[]
  fingerprints: SearchIndexFileFingerprint[]
}

export interface SearchIndexSearchOptions {
  scope: SearchScope
  scopePath?: string
  casePath?: string
  mailboxKey?: string
  allowedMailboxKeys?: string[]
  reviewerUsername?: string
  sourceType?: SearchSourceType | 'all'
  requirePreviewPayload?: boolean
  collapseDuplicates?: boolean
  collapseProgress?: MailboxCollapseJobStatus
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
  flaggedSizeBytes: number
  collapseProgress?: MailboxCollapseJobStatus
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
  replaceMailboxDocumentsFromStream?(
    mailboxKey: string,
    documents: AsyncIterable<SearchIndexDocument>
  ): Promise<void>
  upsertMailboxDocument(mailboxKey: string, document: SearchIndexDocument): Promise<void>
  deleteMailboxDocuments(mailboxKey: string): Promise<void>
  rebuildMailboxCollapseMetadata?(): Promise<void>
  getMailboxCollapseDocuments?(): Promise<SearchIndexDocument[]>
  resetMailboxCollapseMetadata?(): Promise<void>
  writeMailboxCollapsePartition?(
    partitionKey: string,
    documents: SearchIndexDocument[],
    referencesByIdentity: Map<string, SearchThreadCollapseReference[]>
  ): Promise<void>
  finalizeMailboxCollapseMetadata?(version: number): Promise<void>
  getMailboxCollapseJob?(): Promise<MailboxCollapseJobStatus | null>
  saveMailboxCollapseJob?(status: MailboxCollapseJobStatus): Promise<void>
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
  findThreadById(
    id: string,
    options?: { allowedMailboxKeys?: string[]; reviewerUsername?: string }
  ): Promise<SearchThreadGroup | null>
  listFileFingerprints(source: SearchIndexRefreshSource): Promise<SearchIndexFileFingerprint[]>
  upsertFileFingerprint(source: SearchIndexRefreshSource, fingerprint: SearchIndexFileFingerprint): Promise<void>
  replaceFileFingerprints(source: SearchIndexRefreshSource, fingerprints: SearchIndexFileFingerprint[]): Promise<void>
  deleteFileFingerprints(source: SearchIndexRefreshSource, mailboxKeys: string[]): Promise<void>
  close(): Promise<void>
}

interface SearchIndexCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  aggregate?: (
    pipeline: Record<string, unknown>[]
  ) => {
    toArray: () => Promise<Array<{
      total?: Array<{ value?: number }>
      sourceCounts?: Array<{ _id?: unknown; count?: number }>
      flaggedSizeBytes?: Array<{ value?: number }>
    }>>
  }
  insertMany: (documents: SearchIndexDocument[]) => Promise<unknown>
  deleteMany: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
  findOne: (filter: Record<string, unknown>) => Promise<SearchIndexDocument | null>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
  updateMany?: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>
  ) => Promise<unknown>
  bulkWrite?: (
    operations: Array<{
      updateOne: {
        filter: Record<string, unknown>
        update: Record<string, unknown>
      }
    }>,
    options?: { ordered?: boolean }
  ) => Promise<unknown>
  find: (
    filter: Record<string, unknown>,
    options?: { projection?: Record<string, 0 | 1> }
  ) => {
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

interface FileFingerprintCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  find: (filter: Record<string, unknown>) => {
    sort: (sort: Record<string, 1 | -1>) => {
      toArray: () => Promise<SearchIndexFileFingerprint[]>
    }
  }
  findOne: (filter: Record<string, unknown>) => Promise<SearchIndexFileFingerprint | null>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
  deleteMany: (filter: Record<string, unknown>) => Promise<{ deletedCount?: number }>
}

const DEFAULT_INDEX_COLLECTION = 'pst_search_documents'
const DEFAULT_RULE_COLLECTION = 'pst_search_hidden_rules'
const DEFAULT_FINGERPRINT_COLLECTION = 'pst_search_file_fingerprints'
const DEFAULT_COLLAPSE_JOB_COLLECTION = 'pst_search_collapse_jobs'
export const CURRENT_MAILBOX_SEARCH_INDEX_VERSION = 5
export const CURRENT_MAILBOX_COLLAPSE_VERSION = 6
const CURRENT_ITEM_SEARCH_INDEX_VERSION = 0
const MAX_SEARCH_DOCUMENT_BYTES = 12 * 1024 * 1024
const MAX_INDEXED_BODY_TEXT_CHARS = 192 * 1024
const MAX_INDEXED_BODY_HTML_CHARS = 192 * 1024
const MAX_INDEXED_BODY_PREFIX_CHARS = 32 * 1024
const MAX_INDEXED_BODY_RTF_CHARS = 64 * 1024
const MAX_INDEXED_HEADERS_CHARS = 64 * 1024
const MAX_INDEXED_SEARCH_TEXT_CHARS = 256 * 1024
const MAX_INDEXED_TOKEN_COUNT = 50000
const MAX_INDEXED_ATTACHMENT_COUNT = 256
const MAX_INDEXED_ATTACHMENT_TEXT_CHARS = 4096
const MEMORY_STREAM_BATCH_SIZE = 25

interface MailboxCollapseJobCollectionLike {
  createIndex?: (index: Record<string, 1 | -1>, options?: { unique?: boolean; name?: string }) => Promise<unknown>
  findOne: (filter: Record<string, unknown>) => Promise<MailboxCollapseJobStatus | null>
  updateOne: (
    filter: Record<string, unknown>,
    update: Record<string, unknown>,
    options?: { upsert?: boolean }
  ) => Promise<unknown>
}

// Dedupe only needs relationship, sort, identity, and review metadata. In
// particular, do not pull mailboxDetail/body/searchText into the in-memory
// thread graph; some indexed messages are several megabytes each.
const COLLAPSE_DOCUMENT_PROJECTION: Record<string, 0 | 1> = {
  _id: 0,
  id: 1,
  sourceType: 1,
  threadMetadata: 1,
  threadCollapse: 1,
  threadCollapsePartitions: 1,
  threadCollapseVersion: 1,
  mailboxKey: 1,
  scopePath: 1,
  scopeLabel: 1,
  fileName: 1,
  mailboxName: 1,
  messageId: 1,
  descriptorId: 1,
  folderId: 1,
  folderPath: 1,
  order: 1,
  messageClass: 1,
  kind: 1,
  size: 1,
  subject: 1,
  originalSubject: 1,
  senderName: 1,
  senderEmailAddress: 1,
  recipientText: 1,
  displayTo: 1,
  displayCC: 1,
  displayBCC: 1,
  resolvedDisplayTo: 1,
  resolvedDisplayCC: 1,
  resolvedDisplayBCC: 1,
  sortDate: 1,
  sortDateMs: 1,
  isMailLike: 1,
  addressValues: 1,
  subjectValues: 1,
  reviewStates: 1,
  archivePath: 1,
  archiveEntryPath: 1,
  archiveEntryChain: 1,
  archiveEntryName: 1
}

export interface MongoSearchIndexStoreConnectOptions {
  documentsCollectionName?: string
  rulesCollectionName?: string
  fingerprintsCollectionName?: string
  collapseJobsCollectionName?: string
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

function normalizeRefreshSource(value: unknown): SearchIndexRefreshSource {
  return normalizeText(value).toLowerCase() === 'items' ? 'items' : 'mailboxes'
}

function getSearchIndexVersion(source: SearchIndexRefreshSource): number {
  return normalizeRefreshSource(source) === 'mailboxes'
    ? CURRENT_MAILBOX_SEARCH_INDEX_VERSION
    : CURRENT_ITEM_SEARCH_INDEX_VERSION
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

function normalizeThreadMessageId(value: unknown): string {
  const text = normalizeText(value).toLowerCase()
  if (!text) {
    return ''
  }
  const bracketed = text.match(/<[^<>]+>/)?.[0]
  const normalized = (bracketed || text.split(/\s+/)[0]).replace(/^</, '').replace(/>$/, '').trim()
  return normalized.includes('@') ? normalized : ''
}

function parseReferenceIds(value: unknown): string[] {
  const text = String(value || '')
  const bracketed = [...text.matchAll(/<([^<>]+)>/g)].map((match) => match[1])
  const values = bracketed.length ? bracketed : text.split(/\s+/)
  return [...new Set(values.map(normalizeThreadMessageId).filter(Boolean))]
}

function readRelationshipHeaders(headers: unknown): { inReplyTo: string; references: string } {
  const raw = String(headers || '')
  if (!raw || !/(?:^|\r?\n)(?:In-Reply-To|References)\s*:/i.test(raw)) {
    return { inReplyTo: '', references: '' }
  }
  const values = { inReplyTo: '', references: '' }
  const unfolded = raw.replace(/\r?\n[ \t]+/g, ' ')
  for (const line of unfolded.split(/\r?\n/)) {
    const match = line.match(/^(In-Reply-To|References)\s*:\s*(.*)$/i)
    if (!match) {
      continue
    }
    if (match[1].toLowerCase() === 'in-reply-to') {
      values.inReplyTo = match[2].trim()
    } else {
      values.references = match[2].trim()
    }
  }
  return values
}

function hasForwardSubject(subject: unknown, originalSubject: unknown): boolean {
  return [subject, originalSubject].some((value) => {
    const normalized = normalizeText(value)
    return /^(?:(?:re\s*:\s*)*)(?:fw|fwd)\s*:/i.test(normalized)
  })
}

function hasForwardBody(bodyText: unknown, bodyHtml: unknown): boolean {
  const htmlText = String(bodyHtml || '').replace(/<[^>]*>/g, ' ')
  const text = `${String(bodyText || '')}\n${htmlText}`.slice(0, 64 * 1024)
  return /begin\s+forwarded\s+message\s*:/i.test(text) ||
    /[-_]{2,}\s*original\s+message\s*[-_]{2,}/i.test(text)
}

export function buildMailboxThreadMetadata(
  detail:
    | (Pick<
        MessageDetail,
        'internetMessageId' | 'inReplyToId' | 'transportMessageHeaders' | 'conversationId'
      > &
        Partial<Pick<MessageDetail, 'subject' | 'originalSubject' | 'bodyText' | 'bodyHtml'>>)
    | undefined,
  kind: SearchIndexDocument['kind']
): MailboxThreadMetadata | undefined {
  const normalizedKind = normalizeText(kind).toLowerCase()
  if (normalizedKind !== 'mail' && normalizedKind !== 'appointment') {
    return undefined
  }

  const relationshipHeaders = readRelationshipHeaders(detail?.transportMessageHeaders)
  const messageId = normalizeThreadMessageId(detail?.internetMessageId)
  const inReplyToId = normalizeThreadMessageId(
    detail?.inReplyToId || relationshipHeaders.inReplyTo
  )
  const referenceIds = parseReferenceIds(relationshipHeaders.references)
  const conversationId = normalizedKind === 'appointment' ? normalizeExactValue(detail?.conversationId || '') : ''
  const isForward = normalizedKind === 'mail' &&
    (hasForwardSubject(detail?.subject, detail?.originalSubject) ||
      hasForwardBody(detail?.bodyText, detail?.bodyHtml))
  if (!messageId && !inReplyToId && !referenceIds.length && !conversationId && !isForward) {
    return undefined
  }
  return { messageId, inReplyToId, referenceIds, conversationId, isForward }
}

function getSearchDocumentIdentity(document: SearchIndexDocument): string {
  return `${normalizeText(document.sourceType)}\u0000${normalizeText(document.mailboxKey)}\u0000${normalizeText(document.messageId)}`
}

function getSearchDocumentPublicId(document: SearchIndexDocument): string {
  return document.sourceType === 'mailbox'
    ? buildMailboxSearchDocumentId(document.mailboxKey, document.messageId)
    : normalizeText(document.id || document.messageId)
}

function compareThreadRecency(left: SearchIndexDocument, right: SearchIndexDocument): number {
  const leftDate = typeof left.sortDateMs === 'number' && Number.isFinite(left.sortDateMs) ? left.sortDateMs : null
  const rightDate = typeof right.sortDateMs === 'number' && Number.isFinite(right.sortDateMs) ? right.sortDateMs : null
  const dateComparison = compareNumberValues(leftDate, rightDate, 'desc')
  if (dateComparison) {
    return dateComparison
  }

  const leftSortDate = normalizeText(left.sortDate || '')
  const rightSortDate = normalizeText(right.sortDate || '')
  if (leftSortDate !== rightSortDate) {
    return rightSortDate.localeCompare(leftSortDate, undefined, { sensitivity: 'base' })
  }

  return getSearchDocumentIdentity(right).localeCompare(getSearchDocumentIdentity(left), undefined, {
    sensitivity: 'base'
  })
}

interface SearchThreadIndex {
  groups: SearchThreadGroup[]
  groupByIdentity: Map<string, SearchThreadGroup>
}

function stableThreadHash(values: string[]): string {
  return createHash('sha256')
    .update([...values].sort().join('\n'))
    .digest('hex')
    .slice(0, 24)
}

function buildMailboxThreadIndex(scopeDocuments: SearchIndexDocument[]): SearchThreadIndex {
  const eligibleDocuments = scopeDocuments.filter((document) => {
    const kind = normalizeText(document.kind).toLowerCase()
    return normalizeSourceType(document.sourceType) === 'mailbox' && (kind === 'mail' || kind === 'appointment')
  })
  const indexesByIdentity = new Map<string, number>()
  eligibleDocuments.forEach((document, index) => {
    indexesByIdentity.set(getSearchDocumentIdentity(document), index)
  })

  const parents = eligibleDocuments.map((_, index) => index)
  const find = (index: number): number => {
    let root = index
    while (parents[root] !== root) {
      root = parents[root]
    }
    while (parents[index] !== index) {
      const next = parents[index]
      parents[index] = root
      index = next
    }
    return root
  }
  const union = (left: number, right: number): void => {
    const leftRoot = find(left)
    const rightRoot = find(right)
    if (leftRoot !== rightRoot) {
      parents[rightRoot] = leftRoot
    }
  }

  const messageIdIndexes = new Map<string, number[]>()
  const referenceIndexes = new Map<string, number[]>()
  const appointmentConversationIndexes = new Map<string, number[]>()
  const appendIndex = (map: Map<string, number[]>, key: string, index: number): void => {
    if (!key) {
      return
    }
    const indexes = map.get(key) || []
    indexes.push(index)
    map.set(key, indexes)
  }

  eligibleDocuments.forEach((document, index) => {
    const metadata = document.threadMetadata
    if (!metadata) {
      return
    }
    const messageId = normalizeThreadMessageId(metadata.messageId)
    appendIndex(messageIdIndexes, messageId, index)
    for (const referenceId of metadata.referenceIds || []) {
      appendIndex(referenceIndexes, normalizeThreadMessageId(referenceId), index)
    }
    const kind = normalizeText(document.kind).toLowerCase()
    if (kind === 'appointment') {
      appendIndex(appointmentConversationIndexes, normalizeExactValue(metadata.conversationId), index)
    }
  })

  const parentIndexesByIndex = new Map<number, number[]>()
  const childrenByIndex = new Map<number, Set<number>>()
  eligibleDocuments.forEach((document, index) => {
    const metadata = document.threadMetadata
    if (!metadata || metadata.isForward) {
      return
    }
    const parentId = normalizeThreadMessageId(metadata.inReplyToId)
    if (!parentId) {
      return
    }
    const parentIndexes = messageIdIndexes.get(parentId) || []
    if (!parentIndexes.length) {
      return
    }
    parentIndexesByIndex.set(index, parentIndexes)
    for (const parentIndex of parentIndexes) {
      const children = childrenByIndex.get(parentIndex) || new Set<number>()
      children.add(index)
      childrenByIndex.set(parentIndex, children)
    }
  })

  // A reply to a forward can carry the original conversation's References as
  // well. Track its forward ancestry so those copied references cannot bridge
  // the new branch back into the original conversation.
  const forwardBranchesByIndex = new Map<number, Set<string>>()
  eligibleDocuments.forEach((document, index) => {
    const metadata = document.threadMetadata
    const kind = normalizeText(document.kind).toLowerCase()
    const messageId = normalizeThreadMessageId(metadata?.messageId)
    if (metadata?.isForward && kind === 'mail' && messageId) {
      forwardBranchesByIndex.set(index, new Set([messageId]))
    }
  })
  const propagationQueue = [...forwardBranchesByIndex.keys()]
  for (let queueIndex = 0; queueIndex < propagationQueue.length; queueIndex += 1) {
    const parentIndex = propagationQueue[queueIndex]
    const parentBranches = forwardBranchesByIndex.get(parentIndex)
    if (!parentBranches?.size) {
      continue
    }
    for (const childIndex of childrenByIndex.get(parentIndex) || []) {
      const childMetadata = eligibleDocuments[childIndex].threadMetadata
      if (childMetadata?.isForward) {
        continue
      }
      const childBranches = forwardBranchesByIndex.get(childIndex) || new Set<string>()
      const previousSize = childBranches.size
      for (const branchId of parentBranches) {
        childBranches.add(branchId)
      }
      if (childBranches.size !== previousSize) {
        forwardBranchesByIndex.set(childIndex, childBranches)
        propagationQueue.push(childIndex)
      }
    }
  }

  eligibleDocuments.forEach((document, index) => {
    const metadata = document.threadMetadata
    if (!metadata) {
      return
    }
    const kind = normalizeText(document.kind).toLowerCase()
    if (kind === 'mail' && metadata.isForward) {
      // A forward's relationship headers often copy the original chain. It
      // starts a new branch and must not connect backward through those refs.
      return
    }
    const parentId = normalizeThreadMessageId(metadata.inReplyToId)
    const forwardBranches = forwardBranchesByIndex.get(index)
    if (forwardBranches?.size && parentId) {
      // Keep replies attached to their immediate forward branch parent. Do
      // not union their copied References into the original conversation.
      for (const parentIndex of parentIndexesByIndex.get(index) || []) {
        union(index, parentIndex)
      }
      return
    }
    const targets = [metadata.inReplyToId, ...(metadata.referenceIds || [])]
      .map(normalizeThreadMessageId)
      .filter(Boolean)
    for (const target of targets) {
      for (const targetIndex of [
        ...(messageIdIndexes.get(target) || []),
        ...(referenceIndexes.get(target) || [])
      ].filter((candidateIndex) => {
        const candidateMetadata = eligibleDocuments[candidateIndex].threadMetadata
        return !candidateMetadata?.isForward && !forwardBranchesByIndex.get(candidateIndex)?.size
      })) {
        union(index, targetIndex)
      }
    }
  })
  for (const indexes of referenceIndexes.values()) {
    const branchSafeIndexes = indexes.filter((index) => {
      const metadata = eligibleDocuments[index].threadMetadata
      return !metadata?.isForward && !forwardBranchesByIndex.get(index)?.size
    })
    if (branchSafeIndexes.length < 2) {
      continue
    }
    for (let index = 1; index < branchSafeIndexes.length; index++) {
      union(branchSafeIndexes[0], branchSafeIndexes[index])
    }
  }
  for (const indexes of appointmentConversationIndexes.values()) {
    for (let index = 1; index < indexes.length; index++) {
      union(indexes[0], indexes[index])
    }
  }

  const documentsByRoot = new Map<number, SearchIndexDocument[]>()
  eligibleDocuments.forEach((document, index) => {
    const root = find(index)
    const documents = documentsByRoot.get(root) || []
    documents.push(document)
    documentsByRoot.set(root, documents)
  })

  const groups: SearchThreadGroup[] = []
  const groupByIdentity = new Map<string, SearchThreadGroup>()
  const documentsByPublicId = new Map(
    eligibleDocuments.map((document) => [getSearchDocumentPublicId(document), document])
  )
  for (const component of documentsByRoot.values()) {
    const componentIndexes = component.map((document) => indexesByIdentity.get(getSearchDocumentIdentity(document)))
      .filter((index): index is number => index !== undefined)
    const appointmentComponent = component.every(
      (document) => normalizeText(document.kind).toLowerCase() === 'appointment'
    )
    const leafIndexes = componentIndexes.filter((index) => !childrenByIndex.has(index))
    const branchLeafIndexes = appointmentComponent
      ? [componentIndexes[0]]
      : leafIndexes.length
        ? leafIndexes
        : componentIndexes
    const threadId = `thread-${stableThreadHash(component.map(getSearchDocumentIdentity))}`
    const branches = branchLeafIndexes.map((leafIndex) => {
      const branchIndexes = appointmentComponent
        ? new Set(componentIndexes)
        : new Set<number>()
      if (!appointmentComponent) {
        const pending = [leafIndex]
        while (pending.length) {
          const currentIndex = pending.pop()
          if (currentIndex === undefined || branchIndexes.has(currentIndex)) {
            continue
          }
          branchIndexes.add(currentIndex)
          pending.push(...(parentIndexesByIndex.get(currentIndex) || []))
        }
      }
      const branchItems = [...branchIndexes]
        .map((index) => eligibleDocuments[index])
        .sort((left, right) => compareThreadRecency(left, right))
      const representative = branchItems.reduce((latest, document) => {
        return compareThreadRecency(document, latest) < 0 ? document : latest
      }, branchItems[0])
      return {
        branchId: `branch-${stableThreadHash(branchItems.map(getSearchDocumentIdentity))}`,
        branchIndex: 0,
        branchCount: 0,
        representativeId: getSearchDocumentPublicId(representative),
        items: branchItems
      }
    }).sort((left, right) => {
      const leftRepresentative = documentsByPublicId.get(left.representativeId)
      const rightRepresentative = documentsByPublicId.get(right.representativeId)
      if (!leftRepresentative || !rightRepresentative) {
        return left.branchId.localeCompare(right.branchId)
      }
      return compareThreadRecency(leftRepresentative, rightRepresentative) || left.branchId.localeCompare(right.branchId)
    })
    branches.forEach((branch, index) => {
      branch.branchIndex = index + 1
      branch.branchCount = branches.length
    })
    const hasRelationship = component.some((document) => {
      const metadata = document.threadMetadata
      return Boolean(metadata && (
        metadata.isForward ||
        metadata.inReplyToId ||
        metadata.referenceIds.length ||
        metadata.conversationId
      ))
    })
    if (!hasRelationship && component.length === 1) {
      continue
    }
    const group: SearchThreadGroup = {
      threadId,
      selectedItemId: '',
      branches
    }
    groups.push(group)
    for (const document of component) {
      groupByIdentity.set(getSearchDocumentIdentity(document), group)
    }
  }

  return { groups, groupByIdentity }
}

function addThreadInfo(
  document: SearchIndexDocument,
  group: SearchThreadGroup,
  branch: SearchThreadBranch
): SearchIndexDocument {
  return {
    ...document,
    threadInfo: {
      threadId: group.threadId,
      branchId: branch.branchId,
      branchIndex: branch.branchIndex,
      branchCount: branch.branchCount,
      threadItemCount: new Set(
        group.branches.flatMap((current) => current.items.map(getSearchDocumentIdentity))
      ).size,
      branchItemCount: branch.items.length,
      isRepresentative: true
    }
  }
}

function collapseLatestThreadDocuments(
  scopeDocuments: SearchIndexDocument[],
  matchingDocuments: SearchIndexDocument[]
): SearchIndexDocument[] {
  const threadIndex = buildMailboxThreadIndex(scopeDocuments)
  const matchingIdentities = new Set(matchingDocuments.map(getSearchDocumentIdentity))
  const representatives = new Map<string, SearchIndexDocument>()
  const standaloneMatches: SearchIndexDocument[] = []
  for (const matchingDocument of matchingDocuments) {
    const identity = getSearchDocumentIdentity(matchingDocument)
    const group = threadIndex.groupByIdentity.get(identity)
    if (!group) {
      standaloneMatches.push(matchingDocument)
      continue
    }
    group.selectedItemId = identity
    for (const branch of group.branches) {
      if (!branch.items.some((document) => matchingIdentities.has(getSearchDocumentIdentity(document)))) {
        continue
      }
      const representative = branch.items.find(
        (document) => getSearchDocumentPublicId(document) === branch.representativeId
      )
      if (representative) {
        representatives.set(getSearchDocumentIdentity(representative), addThreadInfo(representative, group, branch))
      }
    }
  }

  return [...standaloneMatches, ...representatives.values()]
}

export function getMailboxCollapsePartitionKeys(document: SearchIndexDocument): string[] {
  const keys = new Set<string>(['all', `pst:${normalizeText(document.mailboxKey)}`])
  const scopePath = normalizeText(document.scopePath)
  if (!scopePath) {
    return [...keys]
  }

  const parts = scopePath.split('/').filter(Boolean)
  for (let index = 1; index <= parts.length; index += 1) {
    keys.add(`case:${parts.slice(0, index).join('/')}`)
  }
  keys.add(`search:${scopePath}`)
  return [...keys]
}

function getCollapsePartitionCandidates(
  document: SearchIndexDocument,
  options: SearchIndexSearchOptions
): string[] {
  const mailboxPartition = `pst:${normalizeText(document.mailboxKey)}`
  const searchPartition = `search:${normalizeText(document.scopePath)}`
  const casePath = normalizeText(options.casePath)
  if (options.scope === 'pst') {
    return [`pst:${normalizeText(options.mailboxKey)}`]
  }
  if (options.scope === 'search') {
    return [`search:${normalizeText(options.scopePath)}`, mailboxPartition]
  }
  if (casePath) {
    return [`case:${casePath}`, searchPartition, mailboxPartition]
  }

  const casePartitions: string[] = []
  const parts = normalizeText(document.scopePath).split('/').filter(Boolean)
  for (let index = parts.length; index > 0; index -= 1) {
    casePartitions.push(`case:${parts.slice(0, index).join('/')}`)
  }
  return ['all', ...casePartitions, searchPartition, mailboxPartition]
}

function isCollapsePartitionCompleted(
  document: SearchIndexDocument,
  partitionKey: string,
  options: SearchIndexSearchOptions
): boolean {
  if (document.threadCollapsePartitions?.includes(partitionKey)) {
    if (!options.collapseProgress || options.collapseProgress.status === 'succeeded') {
      return true
    }
    return options.collapseProgress.completedPartitionKeys.includes(partitionKey)
  }
  return (
    document.threadCollapseVersion === CURRENT_MAILBOX_SEARCH_INDEX_VERSION &&
    !document.threadCollapsePartitions?.length &&
    (!options.collapseProgress || options.collapseProgress.status === 'succeeded') &&
    partitionKey === 'all'
  )
}

export function buildMailboxCollapsePartitionMetadata(
  partitionKey: string,
  partitionRecords: SearchIndexDocument[],
  hiddenRules: HiddenRuleRecord[]
): Map<string, SearchThreadCollapseReference[]> {
  const visibleOptions: SearchIndexSearchOptions = {
    scope: 'all',
    sourceType: 'all',
    query: '',
    mode: 'and',
    mailOnly: false,
    sort: 'date-desc',
    page: 1,
    pageSize: 1,
    reviewFlaggedOnly: false,
    reviewTaggedOnly: false,
    reviewTag: ''
  }
  const visibleRecords = partitionRecords.filter((record) =>
    matchesDocument(record, visibleOptions, hiddenRules)
  )
  const threadIndex = buildMailboxThreadIndex(visibleRecords)
  const referencesByIdentity = new Map<string, SearchThreadCollapseReference[]>()
  for (const group of threadIndex.groups) {
    const documentsByPublicId = new Map(
      group.branches.flatMap((branch) => branch.items).map((document) => [
        getSearchDocumentPublicId(document),
        document
      ])
    )
    for (const branch of group.branches) {
      const representative = documentsByPublicId.get(branch.representativeId)
      if (!representative) {
        continue
      }
      const reference: SearchThreadCollapseReference = {
        partitionKey,
        threadId: group.threadId,
        branchId: branch.branchId,
        branchIndex: branch.branchIndex,
        branchCount: branch.branchCount,
        threadItemCount: new Set(
          group.branches.flatMap((current) => current.items.map(getSearchDocumentIdentity))
        ).size,
        branchItemCount: branch.items.length,
        isRepresentative: true,
        representativeMailboxKey: normalizeText(representative.mailboxKey),
        representativeMessageId: normalizeText(representative.messageId)
      }
      for (const member of branch.items) {
        const identity = getSearchDocumentIdentity(member)
        const current = referencesByIdentity.get(identity) || []
        if (!current.some((entry) => entry.partitionKey === partitionKey && entry.branchId === reference.branchId)) {
          current.push(reference)
          referencesByIdentity.set(identity, current)
        }
      }
    }
  }
  return referencesByIdentity
}

export function buildMailboxCollapseMetadata(
  documents: SearchIndexDocument[],
  hiddenRules: HiddenRuleRecord[]
): Map<string, SearchThreadCollapseReference[]> {
  const partitionDocuments = new Map<string, SearchIndexDocument[]>()
  for (const document of documents) {
    if (normalizeSourceType(document.sourceType) !== 'mailbox') {
      continue
    }
    for (const partitionKey of getMailboxCollapsePartitionKeys(document)) {
      const current = partitionDocuments.get(partitionKey) || []
      current.push(document)
      partitionDocuments.set(partitionKey, current)
    }
  }

  const referencesByIdentity = new Map<string, SearchThreadCollapseReference[]>()
  for (const [partitionKey, partitionRecords] of partitionDocuments) {
    const partitionReferences = buildMailboxCollapsePartitionMetadata(
      partitionKey,
      partitionRecords,
      hiddenRules
    )
    for (const [identity, references] of partitionReferences) {
      referencesByIdentity.set(identity, [
        ...(referencesByIdentity.get(identity) || []),
        ...references
      ])
    }
  }

  return referencesByIdentity
}

function addThreadInfoFromCollapseReference(
  document: SearchIndexDocument,
  reference: SearchThreadCollapseReference
): SearchIndexDocument {
  return {
    ...document,
    threadInfo: {
      threadId: reference.threadId,
      branchId: reference.branchId,
      branchIndex: reference.branchIndex,
      branchCount: reference.branchCount,
      threadItemCount: reference.threadItemCount,
      branchItemCount: reference.branchItemCount,
      isRepresentative: reference.isRepresentative
    }
  }
}

function collapseLatestThreadDocumentsFromMetadata(
  scopeDocuments: SearchIndexDocument[],
  matchingDocuments: SearchIndexDocument[],
  options: SearchIndexSearchOptions
): SearchIndexDocument[] {
  const scopeByIdentity = new Map(
    scopeDocuments.map((document) => [getSearchDocumentIdentity(document), document])
  )
  const representatives = new Map<string, SearchIndexDocument>()
  const standaloneMatches: SearchIndexDocument[] = []

  for (const matchingDocument of matchingDocuments) {
    const partitionKey = getCollapsePartitionCandidates(matchingDocument, options).find((key) =>
      isCollapsePartitionCompleted(matchingDocument, key, options)
    )
    const reference = partitionKey
      ? (matchingDocument.threadCollapse || []).find((entry) => entry.partitionKey === partitionKey)
      : undefined
    if (!reference) {
      standaloneMatches.push(matchingDocument)
      continue
    }

    const representativeIdentity = getSearchDocumentIdentity({
      sourceType: 'mailbox',
      mailboxKey: reference.representativeMailboxKey,
      messageId: reference.representativeMessageId
    } as SearchIndexDocument)
    const representative = scopeByIdentity.get(representativeIdentity)
    if (!representative) {
      // A representative outside the active scope must never be leaked. The
      // stale/partial scope is treated as an independent visible match until
      // the next metadata rebuild completes.
      standaloneMatches.push(matchingDocument)
      continue
    }
    representatives.set(
      representativeIdentity,
      addThreadInfoFromCollapseReference(representative, reference)
    )
  }

  return [...standaloneMatches, ...representatives.values()]
}

function findMailboxThreadGroup(
  scopeDocuments: SearchIndexDocument[],
  itemId: string
): SearchThreadGroup | null {
  const targetIdentity = getSearchDocumentIdentityFromId(itemId, scopeDocuments)
  if (!targetIdentity) {
    return null
  }
  const threadIndex = buildMailboxThreadIndex(scopeDocuments)
  const group = threadIndex.groupByIdentity.get(targetIdentity)
  if (!group) {
    return null
  }
  const target = scopeDocuments.find((document) => getSearchDocumentIdentity(document) === targetIdentity)
  group.selectedItemId = target ? getSearchDocumentPublicId(target) : ''
  return group
}

function getSearchDocumentIdentityFromId(
  itemId: string,
  documents: SearchIndexDocument[]
): string | null {
  const normalizedId = normalizeText(itemId)
  const parsedId = parseMailboxSearchDocumentId(normalizedId)
  if (parsedId) {
    const exact = documents.find((document) =>
      normalizeSourceType(document.sourceType) === 'mailbox' &&
      normalizeText(document.mailboxKey) === normalizeText(parsedId.mailboxKey) &&
      normalizeText(document.messageId) === normalizeText(parsedId.messageId)
    )
    return exact ? getSearchDocumentIdentity(exact) : null
  }
  const exact = documents.find((document) =>
    normalizeText(document.id || document.messageId) === normalizedId ||
    normalizeText(document.messageId) === normalizedId
  )
  return exact ? getSearchDocumentIdentity(exact) : null
}

function resolveSearchThreadGroup(
  group: SearchThreadGroup,
  reviewerUsername?: string
): SearchThreadGroup {
  return {
    ...group,
    branches: group.branches.map((branch) => ({
      ...branch,
      items: branch.items.map((document) =>
        resolveSearchIndexDocument(document, reviewerUsername)
      )
    }))
  }
}

export function buildMailboxSearchDocumentId(mailboxKey: string, messageId: string): string {
  return `${encodeURIComponent(normalizeText(mailboxKey))}::${encodeURIComponent(normalizeText(messageId))}`
}

export function parseMailboxSearchDocumentId(
  itemId: string
): { mailboxKey: string; messageId: string } | null {
  const normalized = normalizeText(itemId)
  if (!normalized) {
    return null
  }

  const separatorIndex = normalized.indexOf('::')
  if (separatorIndex < 0) {
    return null
  }

  const mailboxKeyPart = normalized.slice(0, separatorIndex)
  const messageIdPart = normalized.slice(separatorIndex + 2)
  if (!mailboxKeyPart || !messageIdPart) {
    return null
  }

  try {
    return {
      mailboxKey: decodeURIComponent(mailboxKeyPart),
      messageId: decodeURIComponent(messageIdPart)
    }
  } catch {
    return null
  }
}

function compactMailboxDetail(detail: MessageDetail): MessageDetail {
  const cloned = cloneMessageDetail(detail)
  const limit = (value: unknown, max: number): string => {
    const text = String(value ?? '')
    return text.length > max ? text.slice(0, max) : text
  }
  const compactAttachment = (attachment: MessageDetail['attachments'][number]) => ({
    ...attachment,
    filename: limit(attachment.filename, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    longFilename: limit(attachment.longFilename, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    downloadFilename: limit(attachment.downloadFilename, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    mimeTag: limit(attachment.mimeTag, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    contentId: limit(attachment.contentId, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    pathname: limit(attachment.pathname, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    longPathname: limit(attachment.longPathname, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    downloadUrl: '',
    embeddedMessage: attachment.embeddedMessage ? compactMailboxDetail(attachment.embeddedMessage) : null
  })
  return {
    ...cloned,
    subject: limit(cloned.subject, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    originalSubject: limit(cloned.originalSubject, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    senderName: limit(cloned.senderName, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    senderEmailAddress: limit(cloned.senderEmailAddress, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    recipientText: limit(cloned.recipientText, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    displayTo: limit(cloned.displayTo, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    displayCC: limit(cloned.displayCC, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    displayBCC: limit(cloned.displayBCC, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    resolvedDisplayTo: limit(cloned.resolvedDisplayTo, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    resolvedDisplayCC: limit(cloned.resolvedDisplayCC, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    resolvedDisplayBCC: limit(cloned.resolvedDisplayBCC, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    bodyPrefix: limit(cloned.bodyPrefix, MAX_INDEXED_BODY_PREFIX_CHARS),
    bodyText: limit(cloned.bodyText, MAX_INDEXED_BODY_TEXT_CHARS),
    bodyHtml: limit(cloned.bodyHtml, MAX_INDEXED_BODY_HTML_CHARS),
    bodyRtf: limit(cloned.bodyRtf, MAX_INDEXED_BODY_RTF_CHARS),
    transportMessageHeaders: limit(cloned.transportMessageHeaders, MAX_INDEXED_HEADERS_CHARS),
    conversationTopic: limit(cloned.conversationTopic, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    internetMessageId: limit(cloned.internetMessageId, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    inReplyToId: limit(cloned.inReplyToId, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    returnPath: limit(cloned.returnPath, MAX_INDEXED_ATTACHMENT_TEXT_CHARS),
    attachments: (cloned.attachments || []).slice(0, MAX_INDEXED_ATTACHMENT_COUNT).map(compactAttachment)
  }
}

function estimateSearchDocumentBytes(document: SearchIndexDocument): number {
  return Buffer.byteLength(JSON.stringify(document), 'utf8')
}

function compactSearchIndexDocument(document: SearchIndexDocument): SearchIndexDocument {
  let compacted: SearchIndexDocument = {
    ...document,
    threadCollapse: document.threadCollapse?.slice(0, 16),
    bodySearchText: document.bodySearchText.slice(0, MAX_INDEXED_SEARCH_TEXT_CHARS),
    searchText: document.searchText.slice(0, MAX_INDEXED_SEARCH_TEXT_CHARS),
    searchTokens: document.searchTokens.slice(0, MAX_INDEXED_TOKEN_COUNT),
    mailboxDetail: document.mailboxDetail ? compactMailboxDetail(document.mailboxDetail) : undefined
  }

  if (estimateSearchDocumentBytes(compacted) <= MAX_SEARCH_DOCUMENT_BYTES) {
    return compacted
  }

  compacted = {
    ...compacted,
    bodySearchText: compacted.bodySearchText.slice(0, 64 * 1024),
    searchText: compacted.searchText.slice(0, 128 * 1024),
    searchTokens: compacted.searchTokens.slice(0, 20000),
    mailboxDetail: compacted.mailboxDetail
      ? {
          ...compacted.mailboxDetail,
          bodyPrefix: '',
          bodyHtml: '',
          bodyRtf: '',
          transportMessageHeaders: '',
          attachments: compacted.mailboxDetail.attachments.slice(0, 64)
        }
      : undefined
  }

  if (estimateSearchDocumentBytes(compacted) <= MAX_SEARCH_DOCUMENT_BYTES) {
    return compacted
  }

  // Keep the searchable metadata and a header-only preview rather than
  // allowing one pathological message to abort the whole mailbox refresh.
  return {
    ...compacted,
    bodySearchText: compacted.bodySearchText.slice(0, 16 * 1024),
    searchText: compacted.searchText.slice(0, 64 * 1024),
    searchTokens: compacted.searchTokens.slice(0, 10000),
    mailboxDetail: compacted.mailboxDetail
      ? {
          ...compacted.mailboxDetail,
          bodyText: '',
          bodyHtml: '',
          bodyPrefix: '',
          bodyRtf: '',
          transportMessageHeaders: '',
          attachments: []
        }
      : undefined
  }
}

function uniqueFingerprintValues(values: SearchIndexFileFingerprint[]): SearchIndexFileFingerprint[] {
  const seen = new Set<string>()
  const records: SearchIndexFileFingerprint[] = []
  for (const value of values) {
    const normalized = normalizeFingerprintRecord(value)
    const key = buildFingerprintKey(normalized.source, normalized.mailboxKey)
    if (seen.has(key)) {
      continue
    }
    seen.add(key)
    records.push(normalized)
  }
  return records
}

function buildFingerprintKey(source: SearchIndexRefreshSource, mailboxKey: string): string {
  return `${normalizeRefreshSource(source)}\u0000${normalizeText(mailboxKey)}`
}

function fingerprintMatches(
  left: SearchIndexFileFingerprint | null | undefined,
  right: SearchIndexFileFingerprint
): boolean {
  if (!left) {
    return false
  }
  return (
    normalizeRefreshSource(left.source) === normalizeRefreshSource(right.source) &&
    normalizeText(left.mailboxKey) === normalizeText(right.mailboxKey) &&
    Number(left.size || 0) === Number(right.size || 0) &&
    normalizeText(left.modifiedAt || '') === normalizeText(right.modifiedAt || '') &&
    Number(left.indexVersion || 0) === Number(right.indexVersion || 0)
  )
}

function normalizeFingerprintRecord(record: SearchIndexFileFingerprint): SearchIndexFileFingerprint {
  return {
    source: normalizeRefreshSource(record.source),
    mailboxKey: normalizeText(record.mailboxKey),
    fileName: normalizeText(record.fileName),
    scopePath: normalizeText(record.scopePath),
    scopeLabel: normalizeText(record.scopeLabel),
    size: Number.isFinite(record.size) ? Number(record.size) : 0,
    modifiedAt: record.modifiedAt ? normalizeText(record.modifiedAt) || null : null,
    indexVersion: Number.isFinite(record.indexVersion) ? Number(record.indexVersion) : 0,
    updatedAt: normalizeText(record.updatedAt) || new Date().toISOString()
  }
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

type ReviewStateEntry = {
  reviewerUsername: string
  review: ReviewState
}

function buildReviewStateEntries(records: ReviewRecord[]): ReviewStateEntry[] {
  return records.map((record) => ({
    reviewerUsername: normalizeReviewerUsername(record.reviewerUsername),
    review: normalizeReview(record)
  }))
}

function buildReviewStatesByMessageId(records: ReviewRecord[]): Map<string, ReviewStateEntry[]> {
  const states = new Map<string, ReviewStateEntry[]>()
  for (const record of records) {
    const messageId = normalizeText(record.messageId)
    if (!messageId) {
      continue
    }
    const entries = states.get(messageId) || []
    entries.push(...buildReviewStateEntries([record]))
    states.set(messageId, entries)
  }
  return states
}

function resolveReviewTagValues(
  reviewStates: Array<{ reviewerUsername: string; review: ReviewState }> | undefined,
  reviewerUsername: string | undefined
): string[] {
  const review = resolveReviewState(reviewStates, reviewerUsername)
  return review ? buildReviewTagValues(review) : []
}

function compareTextValues(left: string, right: string, direction: 'asc' | 'desc'): number {
  const normalizedLeft = normalizeText(left)
  const normalizedRight = normalizeText(right)
  if (normalizedLeft === normalizedRight) {
    return 0
  }
  if (!normalizedLeft) {
    return 1
  }
  if (!normalizedRight) {
    return -1
  }
  const comparison = normalizedLeft.localeCompare(normalizedRight, undefined, { sensitivity: 'base' })
  return direction === 'asc' ? comparison : -comparison
}

function compareNumberValues(left: number | null | undefined, right: number | null | undefined, direction: 'asc' | 'desc'): number {
  const leftValue = typeof left === 'number' && Number.isFinite(left) ? left : null
  const rightValue = typeof right === 'number' && Number.isFinite(right) ? right : null
  if (leftValue === rightValue) {
    return 0
  }
  if (leftValue === null) {
    return 1
  }
  if (rightValue === null) {
    return -1
  }
  return direction === 'asc' ? leftValue - rightValue : rightValue - leftValue
}

function getDocumentSubjectSortText(document: SearchIndexDocument): string {
  return normalizeText(document.subject || document.originalSubject || '')
}

function getDocumentSenderSortText(document: SearchIndexDocument): string {
  return normalizeText(document.senderName || document.senderEmailAddress || '')
}

function getDocumentLocationSortText(document: SearchIndexDocument): string {
  const locationParts =
    document.sourceType === 'mailbox'
      ? [document.scopeLabel, document.folderPath || document.mailboxName || document.fileName || document.scopePath]
      : [document.scopeLabel, document.archiveEntryPath || document.archivePath || document.fileName || document.scopePath]
  return normalizeText(locationParts.filter(Boolean).join(' · '))
}

function compareDefaultDocuments(left: SearchIndexDocument, right: SearchIndexDocument): number {
  const scopeLabelComparison = compareTextValues(left.scopeLabel, right.scopeLabel, 'asc')
  if (scopeLabelComparison) {
    return scopeLabelComparison
  }

  const fileNameComparison = compareTextValues(left.fileName, right.fileName, 'asc')
  if (fileNameComparison) {
    return fileNameComparison
  }

  const folderPathComparison = compareTextValues(left.folderPath, right.folderPath, 'asc')
  if (folderPathComparison) {
    return folderPathComparison
  }

  const archivePathComparison = compareTextValues(left.archivePath || '', right.archivePath || '', 'asc')
  if (archivePathComparison) {
    return archivePathComparison
  }

  const archiveEntryPathComparison = compareTextValues(left.archiveEntryPath || '', right.archiveEntryPath || '', 'asc')
  if (archiveEntryPathComparison) {
    return archiveEntryPathComparison
  }

  const orderComparison = compareNumberValues(left.order, right.order, 'asc')
  if (orderComparison) {
    return orderComparison
  }

  return compareTextValues(left.messageId, right.messageId, 'asc')
}

function sortDocuments(left: SearchIndexDocument, right: SearchIndexDocument, sort: string): number {
  switch (sort) {
    case 'order': {
      const scopeLabelComparison = compareTextValues(left.scopeLabel, right.scopeLabel, 'asc')
      if (scopeLabelComparison) {
        return scopeLabelComparison
      }

      const fileNameComparison = compareTextValues(left.fileName, right.fileName, 'asc')
      if (fileNameComparison) {
        return fileNameComparison
      }

      const folderPathComparison = compareTextValues(left.folderPath, right.folderPath, 'asc')
      if (folderPathComparison) {
        return folderPathComparison
      }

      const archivePathComparison = compareTextValues(left.archivePath || '', right.archivePath || '', 'asc')
      if (archivePathComparison) {
        return archivePathComparison
      }

      const archiveEntryPathComparison = compareTextValues(left.archiveEntryPath || '', right.archiveEntryPath || '', 'asc')
      if (archiveEntryPathComparison) {
        return archiveEntryPathComparison
      }

      const orderComparison = compareNumberValues(left.order, right.order, 'asc')
      if (orderComparison) {
        return orderComparison
      }

      return compareTextValues(left.messageId, right.messageId, 'asc')
    }
    case 'subject-asc':
    case 'subject-desc': {
      const comparison = compareTextValues(
        getDocumentSubjectSortText(left),
        getDocumentSubjectSortText(right),
        sort.endsWith('asc') ? 'asc' : 'desc'
      )
      return comparison || compareDefaultDocuments(left, right)
    }
    case 'sender-asc':
    case 'sender-desc': {
      const comparison = compareTextValues(
        getDocumentSenderSortText(left),
        getDocumentSenderSortText(right),
        sort.endsWith('asc') ? 'asc' : 'desc'
      )
      return comparison || compareDefaultDocuments(left, right)
    }
    case 'location-asc':
    case 'location-desc': {
      const comparison = compareTextValues(
        getDocumentLocationSortText(left),
        getDocumentLocationSortText(right),
        sort.endsWith('asc') ? 'asc' : 'desc'
      )
      return comparison || compareDefaultDocuments(left, right)
    }
    case 'date-asc':
    case 'date-desc':
    default: {
      const comparison = compareNumberValues(
        left.sortDateMs,
        right.sortDateMs,
        sort === 'date-asc' ? 'asc' : 'desc'
      )
      return comparison || compareDefaultDocuments(left, right)
    }
  }
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
  }

  const scopePathClauses: Record<string, unknown>[] = []
  if (options.scope === 'search') {
    const scopePath = normalizeText(options.scopePath)
    if (scopePath) {
      scopePathClauses.push({ scopePath })
    }
  }

  const casePath = normalizeText(options.casePath)
  if (casePath) {
    scopePathClauses.push({
      scopePath: {
        $regex: `^${escapeRegex(casePath)}(?:/|$)`
      }
    })
  }

  if (scopePathClauses.length === 1) {
    Object.assign(filter, scopePathClauses[0])
  } else if (scopePathClauses.length > 1) {
    filter.$and = scopePathClauses
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

  if (options.requirePreviewPayload) {
    const previewClause: Record<string, unknown> = {
      $or: [
        { sourceType: { $ne: 'mailbox' } },
        { mailboxDetail: { $exists: true } }
      ]
    }
    if (filter.$and) {
      ;(filter.$and as Record<string, unknown>[]).push(previewClause)
    } else {
      filter.$and = [previewClause]
    }
  }

  const searchExpression = buildSearchExpression(options.query, options.mode)
  if (Object.keys(searchExpression).length) {
    if (filter.$and) {
      ;(filter.$and as Record<string, unknown>[]).push(searchExpression)
    } else {
      Object.assign(filter, searchExpression)
    }
  }

  return filter
}

function createSortSpec(sort: string): Record<string, 1 | -1> {
  switch (sort) {
    case 'order':
      return {
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        archivePath: 1,
        archiveEntryPath: 1,
        order: 1,
        messageId: 1
      }
    case 'subject-asc':
      return {
        subject: 1,
        originalSubject: 1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
    case 'subject-desc':
      return {
        subject: -1,
        originalSubject: -1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
    case 'sender-asc':
      return {
        senderName: 1,
        senderEmailAddress: 1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
    case 'sender-desc':
      return {
        senderName: -1,
        senderEmailAddress: -1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
    case 'location-asc':
      return {
        scopeLabel: 1,
        folderPath: 1,
        archivePath: 1,
        archiveEntryPath: 1,
        fileName: 1,
        order: 1,
        messageId: 1
      }
    case 'location-desc':
      return {
        scopeLabel: -1,
        folderPath: -1,
        archivePath: -1,
        archiveEntryPath: -1,
        fileName: -1,
        order: -1,
        messageId: -1
      }
    case 'date-asc':
      return {
        sortDateMs: 1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
    case 'date-desc':
    default:
      return {
        sortDateMs: -1,
        scopeLabel: 1,
        fileName: 1,
        folderPath: 1,
        order: 1,
        messageId: 1
      }
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

function calculateFlaggedSizeBytes(records: SearchIndexDocument[]): number {
  return records.reduce((total, record) => {
    return record.review.flagged ? total + (Number(record.size) || 0) : total
  }, 0)
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
  const limit = (value: string, max = MAX_INDEXED_ATTACHMENT_TEXT_CHARS): string =>
    value.length > max ? value.slice(0, max) : value
  const boundedBodySearchText = limit(normalizeExactValue(bodySearchText), MAX_INDEXED_SEARCH_TEXT_CHARS)
  const boundedBase = {
    ...base,
    subject: limit(base.subject),
    originalSubject: limit(base.originalSubject),
    senderName: limit(base.senderName),
    senderEmailAddress: limit(base.senderEmailAddress),
    recipientText: limit(base.recipientText),
    displayTo: limit(base.displayTo),
    displayCC: limit(base.displayCC),
    displayBCC: limit(base.displayBCC),
    resolvedDisplayTo: limit(base.resolvedDisplayTo),
    resolvedDisplayCC: limit(base.resolvedDisplayCC),
    resolvedDisplayBCC: limit(base.resolvedDisplayBCC),
    messageClass: limit(base.messageClass),
    mailboxDetail: base.mailboxDetail ? compactMailboxDetail(base.mailboxDetail) : undefined,
    previewText: base.previewText ? limit(base.previewText, MAX_INDEXED_SEARCH_TEXT_CHARS) : base.previewText,
    previewHtml: base.previewHtml ? limit(base.previewHtml, MAX_INDEXED_SEARCH_TEXT_CHARS) : base.previewHtml
  }
  const searchText = buildSearchText(boundedBase, boundedBodySearchText).slice(0, MAX_INDEXED_SEARCH_TEXT_CHARS)
  const subjectValues = uniqueStrings([boundedBase.subject, boundedBase.originalSubject])
  return compactSearchIndexDocument({
    ...boundedBase,
    bodySearchText: boundedBodySearchText,
    searchText,
    searchTokens: tokenizeSearchText(searchText).slice(0, MAX_INDEXED_TOKEN_COUNT),
    addressValues: extractEmailAddresses(
      boundedBase.senderEmailAddress,
      boundedBase.displayTo,
      boundedBase.displayCC,
      boundedBase.displayBCC,
      boundedBase.resolvedDisplayTo,
      boundedBase.resolvedDisplayCC,
      boundedBase.resolvedDisplayBCC
    ),
    subjectValues,
    review: normalizeReview(null),
    reviewStates: reviewStates.map((entry) => ({
      reviewerUsername: normalizeReviewerUsername(entry.reviewerUsername),
      review: normalizeReview(entry.review)
    })),
    reviewTagValues: [],
    updatedAt: new Date().toISOString()
  })
}

export function buildSearchIndexDocumentFromMailboxMessage(
  record: ViewerIndexedMessage,
  context: {
    mailboxKey: string
    scopePath: string
    scopeLabel: string
    fileName: string
  },
  reviewStatesByMessageId: Map<string, ReviewStateEntry[]>
): SearchIndexDocument {
  const message = record.summary
  const mailboxDetail = record.detail ? compactMailboxDetail(record.detail) : undefined
  const threadMetadata = buildMailboxThreadMetadata(mailboxDetail, message.kind)
  return toDocument(
    {
      id: normalizeText(message.id),
      sourceType: 'mailbox',
      mailboxKey: normalizeText(context.mailboxKey),
      scopePath: normalizeText(context.scopePath),
      scopeLabel: normalizeText(context.scopeLabel),
      fileName: normalizeText(context.fileName),
      mailboxName: normalizeText(record.mailboxName),
      messageId: normalizeText(message.id),
      descriptorId: normalizeText(message.descriptorId),
      folderId: normalizeText(message.folderId),
      folderPath: normalizeText(message.folderPath),
      order: message.order,
      messageClass: normalizeText(message.messageClass),
      kind: message.kind,
      size: Number(message.size) || 0,
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
      isMailLike: message.isMailLike,
      mailboxDetail,
      threadMetadata
    },
    record.bodySearchText,
    reviewStatesByMessageId.get(message.id) || []
  )
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
  const reviewStatesByMessageId = buildReviewStatesByMessageId(reviewRecords)
  const documents: SearchIndexDocument[] = []

  for (const message of session.messages.values()) {
    documents.push(
      buildSearchIndexDocumentFromMailboxMessage(
        {
          mailboxName: context.mailboxName,
          summary: message,
          bodySearchText: session.searchTextByMessageId.get(message.id) || '',
          detail: session.messageDetailSnapshots.get(message.id)
        },
        context,
        reviewStatesByMessageId
      )
    )
  }

  return dedupeSearchIndexDocuments(documents)
}

export function buildSearchIndexDocumentsFromArchiveItems(
  items: ArchiveBundleItem[],
  reviewRecords: ReviewRecord[] = []
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
        size: Number(item.entrySize) || 0,
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
      reviewStatesByMessageId.get(item.archiveItemId) || []
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
  private readonly fingerprints = new Map<string, SearchIndexFileFingerprint>()
  private mailboxCollapseJob: MailboxCollapseJobStatus | null = null

  private invalidateMailboxCollapseMetadata(): void {
    for (const mailbox of this.documents.values()) {
      for (const [messageId, document] of mailbox) {
        if (normalizeSourceType(document.sourceType) !== 'mailbox') {
          continue
        }
        mailbox.set(messageId, {
          ...document,
          threadCollapse: [],
          threadCollapsePartitions: [],
          threadCollapseVersion: 0
        })
      }
    }
  }

  async replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void> {
    this.invalidateMailboxCollapseMetadata()
    const key = normalizeText(mailboxKey)
    const records = new Map<string, SearchIndexDocument>()
    for (const document of dedupeSearchIndexDocuments(documents.map(compactSearchIndexDocument))) {
      records.set(document.messageId, {
        ...document,
        mailboxKey: key,
        sourceType: normalizeSourceType(document.sourceType),
        threadCollapse: [],
        threadCollapsePartitions: [],
        threadCollapseVersion: 0
      })
    }
    this.documents.set(key, records)
  }

  async replaceMailboxDocumentsFromStream(
    mailboxKey: string,
    documents: AsyncIterable<SearchIndexDocument>
  ): Promise<void> {
    const key = normalizeText(mailboxKey)
    const seen = new Set<string>()
    const batch: SearchIndexDocument[] = []
    let firstBatch = true

    const flush = async (): Promise<void> => {
      if (!batch.length) {
        return
      }

      const currentBatch = batch.splice(0, batch.length)
      if (firstBatch) {
        firstBatch = false
        // Dispatch through the public method so test doubles and callers that
        // observe mailbox replacement retain their existing behavior.
        await this.replaceMailboxDocuments(key, currentBatch)
        return
      }

      const mailbox = this.documents.get(key) || new Map<string, SearchIndexDocument>()
      for (const document of currentBatch) {
        mailbox.set(document.messageId, {
          ...document,
          mailboxKey: key,
          sourceType: normalizeSourceType(document.sourceType),
          threadCollapse: [],
          threadCollapsePartitions: [],
          threadCollapseVersion: 0
        })
      }
      this.documents.set(key, mailbox)
    }

    for await (const document of documents) {
      const normalizedDocument = compactSearchIndexDocument(document)
      const messageId = normalizeText(normalizedDocument.messageId)
      if (!messageId) {
        continue
      }
      const dedupeKey = `${key}\u0000${messageId}`
      if (seen.has(dedupeKey)) {
        continue
      }
      seen.add(dedupeKey)
      batch.push({
        ...normalizedDocument,
        mailboxKey: key,
        sourceType: normalizeSourceType(normalizedDocument.sourceType)
      })
      if (firstBatch || batch.length >= MEMORY_STREAM_BATCH_SIZE) {
        await flush()
      }
    }

    if (batch.length) {
      await flush()
    } else if (firstBatch) {
      await this.replaceMailboxDocuments(key, [])
    }
  }

  async upsertMailboxDocument(mailboxKey: string, document: SearchIndexDocument): Promise<void> {
    this.invalidateMailboxCollapseMetadata()
    const key = normalizeText(mailboxKey)
    const mailbox = this.documents.get(key) || new Map<string, SearchIndexDocument>()
    const [normalizedDocument] = dedupeSearchIndexDocuments([compactSearchIndexDocument(document)])
    if (!normalizedDocument) {
      return
    }
    mailbox.set(normalizedDocument.messageId, {
      ...normalizedDocument,
      mailboxKey: key,
      sourceType: normalizeSourceType(normalizedDocument.sourceType),
      threadCollapse: [],
      threadCollapsePartitions: [],
      threadCollapseVersion: 0
    })
    this.documents.set(key, mailbox)
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    this.invalidateMailboxCollapseMetadata()
    const key = normalizeText(mailboxKey)
    this.documents.delete(key)
  }

  async getMailboxCollapseDocuments(): Promise<SearchIndexDocument[]> {
    return [...this.documents.values()]
      .flatMap((mailbox) => [...mailbox.values()])
      .filter((document) => normalizeSourceType(document.sourceType) === 'mailbox')
      .map((document) => compactSearchIndexDocument({ ...document }))
  }

  async resetMailboxCollapseMetadata(): Promise<void> {
    this.invalidateMailboxCollapseMetadata()
  }

  async writeMailboxCollapsePartition(
    partitionKey: string,
    documents: SearchIndexDocument[],
    referencesByIdentity: Map<string, SearchThreadCollapseReference[]>
  ): Promise<void> {
    for (const document of documents) {
      const mailbox = this.documents.get(normalizeText(document.mailboxKey))
      const current = mailbox?.get(normalizeText(document.messageId))
      if (!mailbox || !current) {
        continue
      }
      const references = (current.threadCollapse || []).filter((reference) => reference.partitionKey !== partitionKey)
      references.push(...(referencesByIdentity.get(getSearchDocumentIdentity(document)) || []))
      mailbox.set(current.messageId, {
        ...current,
        threadCollapse: references,
        threadCollapsePartitions: [...new Set([...(current.threadCollapsePartitions || []), partitionKey])],
        threadCollapseVersion: 0
      })
    }
  }

  async finalizeMailboxCollapseMetadata(version: number): Promise<void> {
    for (const mailbox of this.documents.values()) {
      for (const [messageId, document] of mailbox) {
        if (normalizeSourceType(document.sourceType) !== 'mailbox') {
          continue
        }
        mailbox.set(messageId, { ...document, threadCollapseVersion: version })
      }
    }
  }

  async getMailboxCollapseJob(): Promise<MailboxCollapseJobStatus | null> {
    return this.mailboxCollapseJob ? { ...this.mailboxCollapseJob } : null
  }

  async saveMailboxCollapseJob(status: MailboxCollapseJobStatus): Promise<void> {
    this.mailboxCollapseJob = { ...status }
  }

  async rebuildMailboxCollapseMetadata(): Promise<void> {
    const documents = [...this.documents.values()]
      .flatMap((mailbox) => [...mailbox.values()])
      .filter((document) => normalizeSourceType(document.sourceType) === 'mailbox')
    const referencesByIdentity = buildMailboxCollapseMetadata(
      documents,
      this.hiddenRules.listHiddenRules()
    )
    for (const mailbox of this.documents.values()) {
      for (const [messageId, document] of mailbox) {
        if (normalizeSourceType(document.sourceType) !== 'mailbox') {
          continue
        }
        mailbox.set(messageId, {
          ...document,
          threadCollapse: referencesByIdentity.get(getSearchDocumentIdentity(document)) || [],
          threadCollapsePartitions: getMailboxCollapsePartitionKeys(document),
          threadCollapseVersion: CURRENT_MAILBOX_COLLAPSE_VERSION
        })
      }
    }
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

  async listFileFingerprints(source: SearchIndexRefreshSource): Promise<SearchIndexFileFingerprint[]> {
    const normalizedSource = normalizeRefreshSource(source)
    return [...this.fingerprints.values()]
      .filter((record) => normalizeRefreshSource(record.source) === normalizedSource)
      .map((record) => ({ ...record }))
      .sort((left, right) => {
        if (left.scopeLabel !== right.scopeLabel) {
          return left.scopeLabel.localeCompare(right.scopeLabel, undefined, { sensitivity: 'base' })
        }
        return left.mailboxKey.localeCompare(right.mailboxKey, undefined, { sensitivity: 'base' })
      })
  }

  async upsertFileFingerprint(
    source: SearchIndexRefreshSource,
    fingerprint: SearchIndexFileFingerprint
  ): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    const normalized = normalizeFingerprintRecord({
      ...fingerprint,
      source: normalizedSource
    })
    this.fingerprints.set(buildFingerprintKey(normalizedSource, normalized.mailboxKey), normalized)
  }

  async replaceFileFingerprints(
    source: SearchIndexRefreshSource,
    fingerprints: SearchIndexFileFingerprint[]
  ): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    for (const key of [...this.fingerprints.keys()]) {
      if (normalizeRefreshSource(this.fingerprints.get(key)?.source) === normalizedSource) {
        this.fingerprints.delete(key)
      }
    }
    for (const record of fingerprints) {
      await this.upsertFileFingerprint(normalizedSource, record)
    }
  }

  async deleteFileFingerprints(source: SearchIndexRefreshSource, mailboxKeys: string[]): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    for (const mailboxKey of mailboxKeys) {
      this.fingerprints.delete(buildFingerprintKey(normalizedSource, mailboxKey))
    }
  }

  async listHiddenRules(): Promise<HiddenRuleRecord[]> {
    return this.hiddenRules.listHiddenRules()
  }

  async upsertHiddenRule(input: { kind: HiddenRuleKind; value: string; label?: string }): Promise<HiddenRuleRecord> {
    const record = this.hiddenRules.upsertHiddenRule(input)
    await this.rebuildMailboxCollapseMetadata()
    return record
  }

  async deleteHiddenRule(filterId: string): Promise<boolean> {
    const deleted = this.hiddenRules.deleteHiddenRule(filterId)
    if (deleted) {
      await this.rebuildMailboxCollapseMetadata()
    }
    return deleted
  }

  async findDocumentById(id: string): Promise<SearchIndexDocument | null> {
    const normalizedId = normalizeText(id)
    if (!normalizedId) {
      return null
    }
    const parsedId = parseMailboxSearchDocumentId(normalizedId)
    if (parsedId) {
      const mailbox = this.documents.get(parsedId.mailboxKey)
      const exactRecord = mailbox?.get(parsedId.messageId) || null
      if (exactRecord) {
        return exactRecord
      }
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

  async findThreadById(
    id: string,
    options: { allowedMailboxKeys?: string[]; reviewerUsername?: string } = {}
  ): Promise<SearchThreadGroup | null> {
    const target = await this.findDocumentById(id)
    if (!target || normalizeSourceType(target.sourceType) !== 'mailbox') {
      return null
    }
    const allowedMailboxKeys = options.allowedMailboxKeys
      ? new Set(uniqueTextValues(options.allowedMailboxKeys))
      : null
    const documents = [...this.documents.values()]
      .flatMap((mailbox) => [...mailbox.values()])
      .filter((document) => normalizeSourceType(document.sourceType) === 'mailbox')
      .filter((document) => !allowedMailboxKeys || allowedMailboxKeys.has(normalizeText(document.mailboxKey)))
    const group = findMailboxThreadGroup(documents, id)
    return group ? resolveSearchThreadGroup(group, options.reviewerUsername) : null
  }

  async search(options: SearchIndexSearchOptions): Promise<SearchIndexPage> {
    const hiddenRules = this.hiddenRules.listHiddenRules()
    const mailboxKeysProvided = options.allowedMailboxKeys !== undefined
    const allowedMailboxKeys = uniqueTextValues(options.allowedMailboxKeys || [])
    let allRecords = [...this.documents.values()].flatMap((mailbox) => [...mailbox.values()])
    let records = mailboxKeysProvided
      ? allRecords.filter((record) => allowedMailboxKeys.includes(record.mailboxKey))
      : allRecords
    if (
      options.collapseDuplicates &&
      (!options.sourceType || options.sourceType === 'all' || options.sourceType === 'mailbox')
    ) {
      // The memory store is also used directly by unit tests and local
      // development without a coordinator. Preserve that convenience while
      // production Mongo requests use the resumable coordinator.
      if (
        this.mailboxCollapseJob === null &&
        records.some(
          (record) => normalizeSourceType(record.sourceType) === 'mailbox' &&
            record.threadCollapseVersion !== CURRENT_MAILBOX_COLLAPSE_VERSION
        )
      ) {
        await this.rebuildMailboxCollapseMetadata()
        allRecords = [...this.documents.values()].flatMap((mailbox) => [...mailbox.values()])
        records = mailboxKeysProvided
          ? allRecords.filter((record) => allowedMailboxKeys.includes(record.mailboxKey))
          : allRecords
      }
      const scopeOptions: SearchIndexSearchOptions = {
        ...options,
        sourceType: 'all',
        query: '',
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      }
      const matchingOptions: SearchIndexSearchOptions = {
        ...scopeOptions,
        query: options.query,
        mode: options.mode
      }
      const reviewOptions: SearchIndexSearchOptions = {
        ...scopeOptions,
        reviewFlaggedOnly: options.reviewFlaggedOnly,
        reviewTaggedOnly: options.reviewTaggedOnly,
        reviewTag: options.reviewTag
      }
      const scopeRecords = records.filter((record) => matchesDocument(record, scopeOptions, hiddenRules))
      const matchingRecords = scopeRecords.filter((record) => matchesDocument(record, matchingOptions, hiddenRules))
      const representatives = collapseLatestThreadDocumentsFromMetadata(scopeRecords, matchingRecords, scopeOptions)
        .filter((record) => matchesDocument(record, reviewOptions, hiddenRules))
        .sort((left, right) => sortDocuments(left, right, options.sort))
      const resolvedRepresentatives = representatives.map((record) =>
        resolveSearchIndexDocument(record, options.reviewerUsername)
      )
      const resolvedSourceCounts = resolvedRepresentatives.reduce<Record<SearchSourceType, number>>(
        (acc, record) => {
          const sourceType = normalizeSourceType(record.sourceType)
          acc[sourceType] = (acc[sourceType] || 0) + 1
          return acc
        },
        { mailbox: 0, teams: 0, sharepoint: 0 }
      )
      const selectedRepresentatives = options.sourceType && options.sourceType !== 'all'
        ? resolvedRepresentatives.filter(
            (record) => normalizeSourceType(record.sourceType) === normalizeSourceType(options.sourceType)
          )
        : resolvedRepresentatives
      return paginateSearchResults(selectedRepresentatives, options, hiddenRules, resolvedSourceCounts)
    }

    const countOptions = {
      ...options,
      sourceType: 'all' as const
    }
    const sourceCounts = records
      .filter((record) => matchesDocument(record, countOptions, hiddenRules))
      .reduce<Record<SearchSourceType, number>>(
        (acc, record) => {
          const sourceType = normalizeSourceType(record.sourceType)
          acc[sourceType] = (acc[sourceType] || 0) + 1
          return acc
        },
        { mailbox: 0, teams: 0, sharepoint: 0 }
      )
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
    this.fingerprints.clear()
  }
}

function matchesDocument(
  record: SearchIndexDocument,
  options: SearchIndexSearchOptions,
  hiddenRules: HiddenRuleRecord[]
): boolean {
  if (options.requirePreviewPayload && record.sourceType === 'mailbox' && !record.mailboxDetail) {
    return false
  }

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

  const casePath = normalizeText(options.casePath)
  if (casePath) {
    const normalizedScopePath = normalizeText(record.scopePath)
    if (
      !normalizedScopePath ||
      (normalizedScopePath !== casePath && !normalizedScopePath.startsWith(`${casePath}/`))
    ) {
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
    flaggedSizeBytes: calculateFlaggedSizeBytes(records),
    collapseProgress: options.collapseProgress,
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
    private readonly fingerprints: FileFingerprintCollectionLike,
    private readonly client?: MongoClient,
    private readonly dbName = 'pst-extractor',
    private readonly documentsCollectionName = DEFAULT_INDEX_COLLECTION,
    private readonly rulesCollectionName = DEFAULT_RULE_COLLECTION,
    private readonly fingerprintsCollectionName = DEFAULT_FINGERPRINT_COLLECTION,
    private readonly collapseJobs?: MailboxCollapseJobCollectionLike,
    private readonly collapseJobsCollectionName = DEFAULT_COLLAPSE_JOB_COLLECTION
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
    const fingerprints = db.collection<SearchIndexFileFingerprint>(
      options.fingerprintsCollectionName || DEFAULT_FINGERPRINT_COLLECTION
    )
    const collapseJobs = db.collection<MailboxCollapseJobStatus>(
      options.collapseJobsCollectionName || DEFAULT_COLLAPSE_JOB_COLLECTION
    )
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
    await documents.createIndex?.({ sourceType: 1, threadCollapseVersion: 1 })
    await rules.createIndex?.({ kind: 1, value: 1 }, { unique: true })
    await rules.createIndex?.({ updatedAt: -1 })
    await fingerprints.createIndex?.({ source: 1, mailboxKey: 1 }, { unique: true })
    await fingerprints.createIndex?.({ source: 1, updatedAt: -1 })
    await collapseJobs.createIndex?.({ updatedAt: -1 })
    return new MongoSearchIndexStore(
      documents as unknown as SearchIndexCollectionLike,
      rules as unknown as HiddenRuleCollectionLike,
      fingerprints as unknown as FileFingerprintCollectionLike,
      client,
      dbName,
      options.documentsCollectionName || DEFAULT_INDEX_COLLECTION,
      options.rulesCollectionName || DEFAULT_RULE_COLLECTION,
      options.fingerprintsCollectionName || DEFAULT_FINGERPRINT_COLLECTION,
      collapseJobs as unknown as MailboxCollapseJobCollectionLike,
      options.collapseJobsCollectionName || DEFAULT_COLLAPSE_JOB_COLLECTION
    )
  }

  async replaceMailboxDocuments(mailboxKey: string, documents: SearchIndexDocument[]): Promise<void> {
    const key = normalizeText(mailboxKey)
    await this.documents.deleteMany({ mailboxKey: key })
    const uniqueDocuments = dedupeSearchIndexDocuments(documents.map(compactSearchIndexDocument))
    if (!uniqueDocuments.length) {
      await this.documents.updateMany?.(
        { sourceType: 'mailbox' },
        { $set: { threadCollapse: [], threadCollapsePartitions: [], threadCollapseVersion: 0 } }
      )
      return
    }
    const normalizedDocuments = uniqueDocuments.map((document) => ({
      ...document,
      mailboxKey: key,
      sourceType: normalizeSourceType(document.sourceType),
      threadCollapse: [],
      threadCollapsePartitions: [],
      threadCollapseVersion: 0
    }))
    const batchSize = 100
    for (let start = 0; start < normalizedDocuments.length; start += batchSize) {
      await this.documents.insertMany(normalizedDocuments.slice(start, start + batchSize))
    }
  }

  async replaceMailboxDocumentsFromStream(
    mailboxKey: string,
    documents: AsyncIterable<SearchIndexDocument>
  ): Promise<void> {
    const key = normalizeText(mailboxKey)
    await this.documents.deleteMany({ mailboxKey: key })

    const seen = new Set<string>()
    const batch: SearchIndexDocument[] = []
    const flush = async (): Promise<void> => {
      if (!batch.length) {
        return
      }
      await this.documents.insertMany(batch.splice(0, batch.length))
    }

    for await (const document of documents) {
      const normalizedDocument = compactSearchIndexDocument(document)
      const messageId = normalizeText(normalizedDocument.messageId)
      if (!messageId) {
        continue
      }
      const dedupeKey = `${key}\u0000${messageId}`
      if (seen.has(dedupeKey)) {
        continue
      }
      seen.add(dedupeKey)
      batch.push({
        ...normalizedDocument,
        mailboxKey: key,
        sourceType: normalizeSourceType(normalizedDocument.sourceType),
        threadCollapse: [],
        threadCollapsePartitions: [],
        threadCollapseVersion: 0
      })
      if (batch.length >= 100) {
        await flush()
      }
    }

    await flush()
  }

  async upsertMailboxDocument(mailboxKey: string, document: SearchIndexDocument): Promise<void> {
    const key = normalizeText(mailboxKey)
    const [normalizedDocument] = dedupeSearchIndexDocuments([compactSearchIndexDocument(document)])
    if (!normalizedDocument) {
      return
    }
    await this.documents.updateOne(
      {
        mailboxKey: key,
        messageId: normalizedDocument.messageId
      },
      {
        $set: {
          ...normalizedDocument,
          mailboxKey: key,
          sourceType: normalizeSourceType(normalizedDocument.sourceType),
          threadCollapse: [],
          threadCollapsePartitions: [],
          threadCollapseVersion: 0
        }
      },
      { upsert: true }
    )
  }

  async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
    const key = normalizeText(mailboxKey)
    await this.documents.deleteMany({ mailboxKey: key })
    await this.documents.updateMany?.(
      { sourceType: 'mailbox' },
      { $set: { threadCollapse: [], threadCollapsePartitions: [], threadCollapseVersion: 0 } }
    )
  }

  async getMailboxCollapseDocuments(): Promise<SearchIndexDocument[]> {
    return this.documents
      .find({ sourceType: 'mailbox' }, { projection: COLLAPSE_DOCUMENT_PROJECTION })
      .sort({})
      .skip(0)
      .limit(0)
      .toArray()
  }

  async resetMailboxCollapseMetadata(): Promise<void> {
    await this.documents.updateMany?.(
      { sourceType: 'mailbox' },
      { $set: { threadCollapse: [], threadCollapsePartitions: [], threadCollapseVersion: 0 } }
    )
  }

  async writeMailboxCollapsePartition(
    partitionKey: string,
    documents: SearchIndexDocument[],
    referencesByIdentity: Map<string, SearchThreadCollapseReference[]>
  ): Promise<void> {
    const operations = documents.map((document) => {
      const references = referencesByIdentity.get(getSearchDocumentIdentity(document)) || []
      const addToSet: Record<string, unknown> = {
        threadCollapsePartitions: partitionKey
      }
      if (references.length) {
        addToSet.threadCollapse = { $each: references }
      }
      return {
        updateOne: {
          filter: {
            mailboxKey: normalizeText(document.mailboxKey),
            messageId: normalizeText(document.messageId)
          },
          update: {
            $set: { threadCollapseVersion: 0 },
            $addToSet: addToSet
          }
        }
      }
    })
    const batchSize = 500
    for (let start = 0; start < operations.length; start += batchSize) {
      const batch = operations.slice(start, start + batchSize)
      if (this.documents.bulkWrite) {
        await this.documents.bulkWrite(batch, { ordered: false })
      } else {
        for (const operation of batch) {
          await this.documents.updateOne(operation.updateOne.filter, operation.updateOne.update)
        }
      }
    }
  }

  async finalizeMailboxCollapseMetadata(version: number): Promise<void> {
    await this.documents.updateMany?.(
      { sourceType: 'mailbox' },
      { $set: { threadCollapseVersion: version } }
    )
  }

  async getMailboxCollapseJob(): Promise<MailboxCollapseJobStatus | null> {
    const status = this.collapseJobs ? await this.collapseJobs.findOne({ _id: 'mailbox' }) : null
    if (!status) {
      return null
    }
    return {
      jobId: status.jobId || null,
      status: status.status,
      version: status.version,
      startedAt: status.startedAt || null,
      completedAt: status.completedAt || null,
      updatedAt: status.updatedAt,
      processedPartitions: status.processedPartitions,
      totalPartitions: status.totalPartitions,
      completedPartitionKeys: [...(status.completedPartitionKeys || [])],
      processedWorkUnits: status.processedWorkUnits,
      totalWorkUnits: status.totalWorkUnits,
      percentage: status.percentage,
      provisional: status.provisional,
      error: status.error || null,
      reindexRequired: status.reindexRequired
    }
  }

  async saveMailboxCollapseJob(status: MailboxCollapseJobStatus): Promise<void> {
    if (!this.collapseJobs) {
      return
    }
    await this.collapseJobs.updateOne(
      { _id: 'mailbox' },
      { $set: status },
      { upsert: true }
    )
  }

  async rebuildMailboxCollapseMetadata(): Promise<void> {
    await this.resetMailboxCollapseMetadata()
    const documents = await this.getMailboxCollapseDocuments()
    const referencesByIdentity = buildMailboxCollapseMetadata(
      documents,
      await this.listHiddenRules()
    )
    for (const partitionKey of new Set(documents.flatMap(getMailboxCollapsePartitionKeys))) {
      const partitionDocuments = documents.filter((document) =>
        getMailboxCollapsePartitionKeys(document).includes(partitionKey)
      )
      const partitionReferences = new Map<string, SearchThreadCollapseReference[]>()
      for (const document of partitionDocuments) {
        const references = (referencesByIdentity.get(getSearchDocumentIdentity(document)) || [])
          .filter((reference) => reference.partitionKey === partitionKey)
        if (references.length) {
          partitionReferences.set(getSearchDocumentIdentity(document), references)
        }
      }
      await this.writeMailboxCollapsePartition(partitionKey, partitionDocuments, partitionReferences)
    }
    await this.finalizeMailboxCollapseMetadata(CURRENT_MAILBOX_COLLAPSE_VERSION)
  }

  async listFileFingerprints(source: SearchIndexRefreshSource): Promise<SearchIndexFileFingerprint[]> {
    const normalizedSource = normalizeRefreshSource(source)
    return this.fingerprints
      .find({ source: normalizedSource })
      .sort({ scopeLabel: 1, mailboxKey: 1 })
      .toArray()
  }

  async upsertFileFingerprint(
    source: SearchIndexRefreshSource,
    fingerprint: SearchIndexFileFingerprint
  ): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    const normalized = normalizeFingerprintRecord({
      ...fingerprint,
      source: normalizedSource
    })
    await this.fingerprints.updateOne(
      { source: normalizedSource, mailboxKey: normalized.mailboxKey },
      {
        $set: {
          ...normalized,
          source: normalizedSource
        }
      },
      { upsert: true }
    )
  }

  async replaceFileFingerprints(
    source: SearchIndexRefreshSource,
    fingerprints: SearchIndexFileFingerprint[]
  ): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    await this.fingerprints.deleteMany({ source: normalizedSource })
    const uniqueFingerprints = uniqueFingerprintValues(
      fingerprints.map((record) => ({
        ...record,
        source: normalizedSource
      }))
    )
    for (const record of uniqueFingerprints) {
      await this.upsertFileFingerprint(normalizedSource, record)
    }
  }

  async deleteFileFingerprints(source: SearchIndexRefreshSource, mailboxKeys: string[]): Promise<void> {
    const normalizedSource = normalizeRefreshSource(source)
    const keys = uniqueTextValues(mailboxKeys)
    if (!keys.length) {
      return
    }
    await this.fingerprints.deleteMany({
      source: normalizedSource,
      mailboxKey: { $in: keys }
    })
  }

  async findDocumentById(id: string): Promise<SearchIndexDocument | null> {
    const normalizedId = normalizeText(id)
    if (!normalizedId) {
      return null
    }
    const parsedId = parseMailboxSearchDocumentId(normalizedId)
    if (parsedId) {
      const exactRecord = await this.documents.findOne({
        mailboxKey: parsedId.mailboxKey,
        messageId: parsedId.messageId
      })
      if (exactRecord) {
        return exactRecord
      }
    }
    return this.documents.findOne({
      $or: [{ messageId: normalizedId }, { id: normalizedId }]
    })
  }

  async findThreadById(
    id: string,
    options: { allowedMailboxKeys?: string[]; reviewerUsername?: string } = {}
  ): Promise<SearchThreadGroup | null> {
    const target = await this.findDocumentById(id)
    if (!target || normalizeSourceType(target.sourceType) !== 'mailbox') {
      return null
    }
    const filter: Record<string, unknown> = { sourceType: 'mailbox' }
    if (options.allowedMailboxKeys) {
      filter.mailboxKey = { $in: uniqueTextValues(options.allowedMailboxKeys) }
    }
    const documents = await this.documents
      .find(filter)
      .sort({ sortDateMs: -1 })
      .skip(0)
      .limit(0)
      .toArray()
    const group = findMailboxThreadGroup(documents, id)
    return group ? resolveSearchThreadGroup(group, options.reviewerUsername) : null
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
    await this.rebuildMailboxCollapseMetadata()
    return record
  }

  async deleteHiddenRule(filterId: string): Promise<boolean> {
    const result = await this.rules.deleteOne({ filterId: normalizeText(filterId) })
    if (result.deletedCount && result.deletedCount > 0) {
      await this.rebuildMailboxCollapseMetadata()
    }
    return Boolean(result.deletedCount && result.deletedCount > 0)
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
    const sourceTypeFilter =
      options.sourceType && options.sourceType !== 'all' ? normalizeSourceType(options.sourceType) : null
    const normalizedReviewerUsername = normalizeReviewerUsername(options.reviewerUsername)
    const collapseDuplicates = Boolean(
      options.collapseDuplicates &&
        (!sourceTypeFilter || sourceTypeFilter === 'mailbox')
    )

    if (collapseDuplicates) {
      const scopeOptions: SearchIndexSearchOptions = {
        ...options,
        sourceType: 'all',
        query: '',
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      }
      const matchingOptions: SearchIndexSearchOptions = {
        ...scopeOptions,
        query: options.query,
        mode: options.mode
      }
      const reviewOptions: SearchIndexSearchOptions = {
        ...scopeOptions,
        reviewFlaggedOnly: options.reviewFlaggedOnly,
        reviewTaggedOnly: options.reviewTaggedOnly,
        reviewTag: options.reviewTag,
        // The Mongo filter has already enforced this before projection. The
        // projected graph intentionally does not carry mailboxDetail.
        requirePreviewPayload: false
      }
      const matchingItems = await this.documents
        .find(buildFilterMatch(matchingOptions, hiddenRules), {
          projection: COLLAPSE_DOCUMENT_PROJECTION
        })
        .sort(createSortSpec(options.sort))
        .skip(0)
        .limit(0)
        .toArray()
      const representativeClauses = new Map<string, { mailboxKey: string; messageId: string }>()
      for (const matchingItem of matchingItems) {
        const partitionKey = getCollapsePartitionCandidates(matchingItem, scopeOptions).find((key) =>
          isCollapsePartitionCompleted(matchingItem, key, scopeOptions)
        )
        const reference = partitionKey
          ? (matchingItem.threadCollapse || []).find((entry) => entry.partitionKey === partitionKey)
          : undefined
        if (reference) {
          const identity = `${reference.representativeMailboxKey}\u0000${reference.representativeMessageId}`
          representativeClauses.set(identity, {
            mailboxKey: normalizeText(reference.representativeMailboxKey),
            messageId: normalizeText(reference.representativeMessageId)
          })
        }
      }
      const representativeItems = representativeClauses.size
        ? await this.documents
            .find({
              $and: [
                buildFilterMatch(scopeOptions, hiddenRules),
                { $or: [...representativeClauses.values()] }
              ]
            }, {
              projection: COLLAPSE_DOCUMENT_PROJECTION
            })
            .sort(createSortSpec(options.sort))
            .skip(0)
            .limit(0)
            .toArray()
        : []
      const scopeItems = [
        ...new Map(
          [...matchingItems, ...representativeItems].map((record) => [getSearchDocumentIdentity(record), record])
        ).values()
      ]
      const representatives = collapseLatestThreadDocumentsFromMetadata(
        scopeItems,
        matchingItems,
        scopeOptions
      )
        .filter((record) => matchesDocument(record, reviewOptions, hiddenRules))
        .sort((left, right) => sortDocuments(left, right, options.sort))
      const resolvedAllItems = representatives.map((record) =>
        resolveSearchIndexDocument(record, options.reviewerUsername)
      )
      const sourceCounts = resolvedAllItems.reduce<Record<SearchSourceType, number>>(
        (acc, record) => {
          const sourceType = normalizeSourceType(record.sourceType)
          acc[sourceType] = (acc[sourceType] || 0) + 1
          return acc
        },
        { mailbox: 0, teams: 0, sharepoint: 0 }
      )
      const resolvedItems = sourceTypeFilter
        ? resolvedAllItems.filter((record) => normalizeSourceType(record.sourceType) === sourceTypeFilter)
        : resolvedAllItems
      const total = resolvedItems.length
      const totalPages = Math.max(1, Math.ceil(total / options.pageSize))
      const page = Math.min(Math.max(options.page, 1), totalPages)
      const start = (page - 1) * options.pageSize
      const pageRepresentatives = resolvedItems.slice(start, start + options.pageSize)
      const pageIdentityClauses = pageRepresentatives.map((record) => ({
        mailboxKey: normalizeText(record.mailboxKey),
        messageId: normalizeText(record.messageId)
      }))
      const pageDocuments = pageIdentityClauses.length
        ? await this.documents.find({ $or: pageIdentityClauses }).sort({}).skip(0).limit(0).toArray()
        : []
      const pageDocumentsByIdentity = new Map(
        pageDocuments.map((record) => [getSearchDocumentIdentity(record), record])
      )
      const resolvedPageItems = pageRepresentatives.map((representative) => {
        const fullRecord = pageDocumentsByIdentity.get(getSearchDocumentIdentity(representative))
        return resolveSearchIndexDocument(
          fullRecord
            ? { ...fullRecord, threadInfo: representative.threadInfo }
            : representative,
          options.reviewerUsername
        )
      })
      const parsed = parseSearchTerms(options.query, options.mode)
      return {
        items: resolvedPageItems,
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
        flaggedSizeBytes: calculateFlaggedSizeBytes(resolvedItems),
        collapseProgress: options.collapseProgress,
        reviewFilters: {
          flaggedOnly: options.reviewFlaggedOnly,
          taggedOnly: options.reviewTaggedOnly,
          tag: normalizeText(options.reviewTag)
        }
      }
    }

    const aggregateCursor = this.documents.aggregate?.([
      { $match: countFilter },
      {
        $facet: {
          total: sourceTypeFilter
            ? [{ $match: { sourceType: sourceTypeFilter } }, { $count: 'value' }]
            : [{ $count: 'value' }],
          sourceCounts: [
            {
              $group: {
                _id: '$sourceType',
                count: { $sum: 1 }
              }
            }
          ],
          flaggedSizeBytes: [
            ...(sourceTypeFilter ? [{ $match: { sourceType: sourceTypeFilter } }] : []),
            {
              $match: {
                reviewStates: {
                  $elemMatch: {
                    reviewerUsername: normalizedReviewerUsername,
                    'review.flagged': true
                  }
                }
              }
            },
            {
              $group: {
                _id: null,
                value: { $sum: '$size' }
              }
            }
          ]
        }
      }
    ])

    let total = 0
    let sourceCounts: Record<SearchSourceType, number> = {
      mailbox: 0,
      teams: 0,
      sharepoint: 0
    }
    let flaggedSizeBytes = 0

    if (aggregateCursor) {
      const [facetResult] = await aggregateCursor.toArray()
      total = facetResult?.total?.[0]?.value || 0
      sourceCounts = (facetResult?.sourceCounts || []).reduce<Record<SearchSourceType, number>>(
        (acc, entry) => {
          const sourceType = normalizeSourceType(entry._id)
          acc[sourceType] = typeof entry.count === 'number' && Number.isFinite(entry.count) ? entry.count : 0
          return acc
        },
        {
          mailbox: 0,
          teams: 0,
          sharepoint: 0
        }
      )
      flaggedSizeBytes = facetResult?.flaggedSizeBytes?.[0]?.value || 0
    } else {
      total = await this.documents.countDocuments(filter)
      sourceCounts = {
        mailbox: await this.documents.countDocuments({ ...countFilter, sourceType: 'mailbox' }),
        teams: await this.documents.countDocuments({ ...countFilter, sourceType: 'teams' }),
        sharepoint: await this.documents.countDocuments({ ...countFilter, sourceType: 'sharepoint' })
      }
    }

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
      flaggedSizeBytes,
      collapseProgress: options.collapseProgress,
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

async function discoverRefreshSourceFiles(
  rootPath: string,
  source: SearchIndexRefreshSource
): Promise<Array<{
  mailboxKey: string
  fileName: string
  scopePath: string
  scopeLabel: string
  size: number
  modifiedAt: string | null
}>> {
  const { listPstMailboxFiles } = await import('./pstCatalog')
  const { listArchiveBundleFiles } = await import('./archiveBundles')

  if (normalizeRefreshSource(source) === 'items') {
    const catalog = listArchiveBundleFiles(rootPath)
    return catalog.scopes.flatMap((scope) =>
      scope.files.map((file) => ({
        mailboxKey: path.resolve(rootPath, scope.scopePath ? scope.scopePath : '', file.fileName),
        fileName: file.fileName,
        scopePath: scope.scopePath,
        scopeLabel: scope.scopeLabel,
        size: file.size,
        modifiedAt: file.modifiedAt || null
      }))
    )
  }

  const catalog = listPstMailboxFiles(rootPath)
  return catalog.scopes.flatMap((scope) =>
    scope.files.map((file) => ({
      mailboxKey: path.resolve(rootPath, scope.scopePath ? scope.scopePath : '', file.fileName),
      fileName: file.fileName,
      scopePath: scope.scopePath,
      scopeLabel: scope.scopeLabel,
      size: file.size,
      modifiedAt: file.modifiedAt || null
    }))
  )
}

export async function refreshSearchIndexSourceFromCatalog(
  rootPath: string,
  source: SearchIndexRefreshSource,
  reviewStore: ReviewStore,
  searchIndexStore: SearchIndexStore,
  options: {
    pruneRemovedFiles?: boolean
  } = {}
): Promise<SearchIndexRefreshPlan> {
  const { extractArchiveBundleItems } = await import('./archiveBundles')
  const normalizedSource = normalizeRefreshSource(source)
  const discoveredFiles = await discoverRefreshSourceFiles(rootPath, normalizedSource)
  const existingFingerprints = new Map(
    (await searchIndexStore.listFileFingerprints(normalizedSource)).map((record) => [
      buildFingerprintKey(record.source, record.mailboxKey),
      record
    ])
  )
  const nextFingerprints: SearchIndexFileFingerprint[] = []
  const discoveredMailboxKeys = new Set<string>()
  const removedMailboxKeys: string[] = []
  const changedMailboxKeys: string[] = []
  const warnedMessages = new Set<string>()
  let mailboxCount = 0
  let messageCount = 0
  let skippedCount = 0
  let failedCount = 0

  function warnOnce(scopeLabel: string, fileName: string, error: unknown): void {
    const message = error instanceof Error ? error.message : String(error)
    const key = `${scopeLabel}/${fileName}::${message}`
    if (warnedMessages.has(key)) {
      return
    }
    warnedMessages.add(key)
    console.warn(`Unable to refresh search index for ${scopeLabel}/${fileName}:`, message)
  }

  for (const file of discoveredFiles) {
    const fingerprint: SearchIndexFileFingerprint = normalizeFingerprintRecord({
      source: normalizedSource,
      mailboxKey: file.mailboxKey,
      fileName: file.fileName,
      scopePath: file.scopePath,
      scopeLabel: file.scopeLabel,
      size: file.size,
      modifiedAt: file.modifiedAt,
      indexVersion: getSearchIndexVersion(normalizedSource),
      updatedAt: new Date().toISOString()
    })
    discoveredMailboxKeys.add(fingerprint.mailboxKey)
    const existingFingerprint = existingFingerprints.get(buildFingerprintKey(normalizedSource, fingerprint.mailboxKey))
    if (fingerprintMatches(existingFingerprint, fingerprint)) {
      skippedCount += 1
      nextFingerprints.push(existingFingerprint ? normalizeFingerprintRecord(existingFingerprint) : fingerprint)
      continue
    }

    try {
      let documents: SearchIndexDocument[] = []
      let replacedFromStream = false
      let indexedMessageCount = 0
      if (normalizedSource === 'mailboxes') {
        const reviewStatesByMessageId = buildReviewStatesByMessageId(
          await reviewStore.listReviews(fingerprint.mailboxKey)
        )
        const documentStream = (async function* (): AsyncGenerator<SearchIndexDocument> {
          for await (const record of streamPstMailboxMessages(
            fingerprint.mailboxKey,
            file.fileName,
            { compactForIndex: true }
          )) {
            indexedMessageCount += 1
            yield buildSearchIndexDocumentFromMailboxMessage(
              record,
              {
                mailboxKey: fingerprint.mailboxKey,
                scopePath: file.scopePath,
                scopeLabel: file.scopeLabel,
                fileName: file.fileName
              },
              reviewStatesByMessageId
            )
          }
        })()

        if (searchIndexStore.replaceMailboxDocumentsFromStream) {
          await searchIndexStore.replaceMailboxDocumentsFromStream(
            fingerprint.mailboxKey,
            documentStream
          )
          replacedFromStream = true
        } else {
          for await (const document of documentStream) {
            documents.push(document)
          }
        }
      } else {
        const reviewRecords = await reviewStore.listReviews(fingerprint.mailboxKey)
        const archiveItems = await extractArchiveBundleItems(fingerprint.mailboxKey, file.scopePath, file.fileName)
        documents = buildSearchIndexDocumentsFromArchiveItems(archiveItems, reviewRecords)
      }

      if (!replacedFromStream) {
        await searchIndexStore.replaceMailboxDocuments(fingerprint.mailboxKey, documents)
      }
      await searchIndexStore.upsertFileFingerprint(normalizedSource, fingerprint)
      changedMailboxKeys.push(fingerprint.mailboxKey)
      nextFingerprints.push(fingerprint)
      mailboxCount += 1
      messageCount += normalizedSource === 'mailboxes' ? indexedMessageCount : documents.length
    } catch (error) {
      failedCount += 1
      warnOnce(file.scopeLabel, file.fileName, error)
    }
  }

  for (const record of existingFingerprints.values()) {
    if (!discoveredMailboxKeys.has(record.mailboxKey)) {
      removedMailboxKeys.push(record.mailboxKey)
    }
  }

  if (options.pruneRemovedFiles !== false && removedMailboxKeys.length) {
    for (const mailboxKey of removedMailboxKeys) {
      await searchIndexStore.deleteMailboxDocuments(mailboxKey)
    }
    await searchIndexStore.deleteFileFingerprints(normalizedSource, removedMailboxKeys)
  }

  return {
    source: normalizedSource,
    mailboxCount: nextFingerprints.length,
    messageCount,
    changedCount: changedMailboxKeys.length,
    skippedCount,
    removedCount: removedMailboxKeys.length,
    failedCount,
    changedMailboxKeys,
    removedMailboxKeys,
    fingerprints: nextFingerprints
  }
}

export async function refreshSearchIndexFromCatalog(
  rootPath: string,
  reviewStore: ReviewStore,
  searchIndexStore: SearchIndexStore
): Promise<{
  mailboxCount: number
  messageCount: number
}> {
  const mailboxSummary = await refreshSearchIndexSourceFromCatalog(
    rootPath,
    'mailboxes',
    reviewStore,
    searchIndexStore
  )
  const itemSummary = await refreshSearchIndexSourceFromCatalog(rootPath, 'items', reviewStore, searchIndexStore)
  return {
    mailboxCount: mailboxSummary.mailboxCount + itemSummary.mailboxCount,
    messageCount: mailboxSummary.messageCount + itemSummary.messageCount
  }
}
