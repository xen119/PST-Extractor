import type { AttachmentDetail, FolderMessagePage, MessageDetail, MessageSummary } from './viewer'
import type { ReviewState } from './reviewTypes'

export const EXTRACTION_FIELD_GROUPS = [
  'summary',
  'participants',
  'routing',
  'dates',
  'content',
  'attachments',
  'headers',
  'review'
] as const

export type ExtractionFieldGroup = (typeof EXTRACTION_FIELD_GROUPS)[number]

export interface MessageExtractionRecord {
  [group: string]: unknown
}

export interface FolderExtractionPage {
  folder: {
    id: string
    descriptorId: string
    displayName: string
    path: string
  }
  fields: ExtractionFieldGroup[]
  paging: {
    page: number
    pageSize: number
    total: number
    totalPages: number
  }
  items: Array<{
    messageId: string
    record: MessageExtractionRecord
  }>
}

function normalizeFieldValue(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function dedupePreserveOrder<T>(values: T[]): T[] {
  return Array.from(new Set(values))
}

function flattenFieldValues(value: unknown): string[] {
  if (Array.isArray(value)) {
    const result: string[] = []
    for (const item of value) {
      result.push(...flattenFieldValues(item))
    }
    return result
  }
  if (typeof value === 'string') {
    return value.split(',')
  }
  return []
}

export function normalizeExtractionFields(value: unknown): ExtractionFieldGroup[] {
  const rawValues = flattenFieldValues(value)
  const normalized = rawValues
    .map((item) => normalizeFieldValue(item).toLowerCase())
    .filter(Boolean)

  if (normalized.includes('all')) {
    return [...EXTRACTION_FIELD_GROUPS]
  }

  const filtered = normalized.filter((item): item is ExtractionFieldGroup =>
    (EXTRACTION_FIELD_GROUPS as readonly string[]).includes(item)
  )
  const fields = dedupePreserveOrder(filtered)
  return fields.length > 0 ? fields : ['summary']
}

export function isSummaryOnlyExtraction(fields: ExtractionFieldGroup[]): boolean {
  return fields.every((field) => field === 'summary' || field === 'review')
}

function buildReviewProjection(review: ReviewState | null): ReviewState {
  return review || {
    flagged: false,
    tags: [],
    createdAt: '',
    updatedAt: ''
  }
}

function buildSummaryProjection(summary: MessageSummary): Record<string, unknown> {
  return {
    id: summary.id,
    descriptorId: summary.descriptorId,
    folderId: summary.folderId,
    folderPath: summary.folderPath,
    order: summary.order,
    messageClass: summary.messageClass,
    kind: summary.kind,
    subject: summary.subject,
    senderName: summary.senderName,
    senderEmailAddress: summary.senderEmailAddress,
    recipientText: summary.recipientText,
    displayTo: summary.displayTo,
    displayCC: summary.displayCC,
    displayBCC: summary.displayBCC,
    resolvedDisplayTo: summary.resolvedDisplayTo,
    resolvedDisplayCC: summary.resolvedDisplayCC,
    resolvedDisplayBCC: summary.resolvedDisplayBCC,
    clientSubmitTime: summary.clientSubmitTime,
    creationTime: summary.creationTime,
    modificationTime: summary.modificationTime,
    messageDeliveryTime: summary.messageDeliveryTime,
    sortDate: summary.sortDate,
    sortDateMs: summary.sortDateMs,
    importance: summary.importance,
    hasAttachments: summary.hasAttachments,
    isRead: summary.isRead,
    isMailLike: summary.isMailLike
  }
}

function buildParticipantsProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    senderName: detail.senderName,
    senderEmailAddress: detail.senderEmailAddress,
    sentRepresentingName: detail.sentRepresentingName,
    sentRepresentingAddressType: detail.sentRepresentingAddressType,
    sentRepresentingEmailAddress: detail.sentRepresentingEmailAddress,
    receivedByName: detail.receivedByName,
    receivedByAddressType: detail.receivedByAddressType,
    receivedByAddress: detail.receivedByAddress,
    replyRecipientNames: detail.replyRecipientNames,
    displayTo: detail.displayTo,
    displayCC: detail.displayCC,
    displayBCC: detail.displayBCC,
    resolvedDisplayTo: detail.resolvedDisplayTo,
    resolvedDisplayCC: detail.resolvedDisplayCC,
    resolvedDisplayBCC: detail.resolvedDisplayBCC,
    originalDisplayTo: detail.originalDisplayTo,
    originalDisplayCC: detail.originalDisplayCC,
    originalDisplayBCC: detail.originalDisplayBCC
  }
}

function buildRoutingProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    folderId: detail.folderId,
    folderPath: detail.folderPath,
    messageClass: detail.messageClass,
    conversationTopic: detail.conversationTopic,
    originalSubject: detail.originalSubject,
    internetMessageId: detail.internetMessageId,
    inReplyToId: detail.inReplyToId,
    returnPath: detail.returnPath
  }
}

function buildDatesProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    clientSubmitTime: detail.clientSubmitTime,
    creationTime: detail.creationTime,
    modificationTime: detail.modificationTime,
    messageDeliveryTime: detail.messageDeliveryTime,
    sortDate: detail.sortDate,
    sortDateMs: detail.sortDateMs
  }
}

function buildContentProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    bodyPrefix: detail.bodyPrefix,
    bodyText: detail.bodyText,
    bodyHtml: detail.bodyHtml,
    bodyRtf: detail.bodyRtf
  }
}

function buildHeadersProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    transportMessageHeaders: detail.transportMessageHeaders,
    originalSubject: detail.originalSubject,
    originalDisplayTo: detail.originalDisplayTo,
    originalDisplayCC: detail.originalDisplayCC,
    originalDisplayBCC: detail.originalDisplayBCC,
    internetMessageId: detail.internetMessageId,
    inReplyToId: detail.inReplyToId,
    returnPath: detail.returnPath
  }
}

function buildAttachmentsProjection(detail: MessageDetail): Record<string, unknown> {
  return {
    hasAttachments: detail.attachments.length > 0,
    attachments: detail.attachments.map((attachment) => projectAttachment(attachment))
  }
}

function projectAttachment(attachment: AttachmentDetail): Record<string, unknown> {
  return {
    attachmentId: attachment.attachmentId,
    index: attachment.index,
    filename: attachment.filename,
    longFilename: attachment.longFilename,
    downloadFilename: attachment.downloadFilename,
    mimeTag: attachment.mimeTag,
    size: attachment.size,
    attachMethod: attachment.attachMethod,
    contentId: attachment.contentId,
    pathname: attachment.pathname,
    longPathname: attachment.longPathname,
    isEmbeddedMessage: attachment.isEmbeddedMessage,
    isDownloadable: attachment.isDownloadable,
    downloadUrl: attachment.downloadUrl,
    parseError: attachment.parseError || undefined
  }
}

function buildReviewProjectionFromState(review: ReviewState | null): Record<string, unknown> {
  const state = buildReviewProjection(review)
  return {
    flagged: state.flagged,
    tags: [...state.tags],
    createdAt: state.createdAt,
    updatedAt: state.updatedAt
  }
}

export function buildMessageExtractionRecord(
  detail: MessageDetail,
  review: ReviewState | null,
  fields: ExtractionFieldGroup[]
): MessageExtractionRecord {
  const selectedFields = normalizeExtractionFields(fields)
  const record: MessageExtractionRecord = {}

  if (selectedFields.includes('summary')) {
    record.summary = buildSummaryProjection(detail)
  }
  if (selectedFields.includes('participants')) {
    record.participants = buildParticipantsProjection(detail)
  }
  if (selectedFields.includes('routing')) {
    record.routing = buildRoutingProjection(detail)
  }
  if (selectedFields.includes('dates')) {
    record.dates = buildDatesProjection(detail)
  }
  if (selectedFields.includes('content')) {
    record.content = buildContentProjection(detail)
  }
  if (selectedFields.includes('attachments')) {
    record.attachments = buildAttachmentsProjection(detail)
  }
  if (selectedFields.includes('headers')) {
    record.headers = buildHeadersProjection(detail)
  }
  if (selectedFields.includes('review')) {
    record.review = buildReviewProjectionFromState(review)
  }

  return record
}

export function buildSummaryExtractionRecord(
  summary: MessageSummary,
  review: ReviewState | null,
  fields: ExtractionFieldGroup[]
): MessageExtractionRecord {
  const selectedFields = normalizeExtractionFields(fields)
  const record: MessageExtractionRecord = {}

  if (selectedFields.includes('summary')) {
    record.summary = buildSummaryProjection(summary)
  }
  if (selectedFields.includes('review')) {
    record.review = buildReviewProjectionFromState(review)
  }

  return record
}

export function buildFolderExtractionPage(
  page: FolderMessagePage,
  reviews: Map<string, ReviewState>,
  fields: ExtractionFieldGroup[],
  buildRecord: (summary: MessageSummary, review: ReviewState | null, fields: ExtractionFieldGroup[]) => MessageExtractionRecord
): FolderExtractionPage {
  return {
    folder: {
      id: page.folder.id,
      descriptorId: page.folder.descriptorId,
      displayName: page.folder.displayName,
      path: page.folder.path
    },
    fields: normalizeExtractionFields(fields),
    paging: {
      page: page.page,
      pageSize: page.pageSize,
      total: page.total,
      totalPages: page.totalPages
    },
    items: page.items.map((item) => ({
      messageId: item.id,
      record: buildRecord(item, reviews.get(item.id) || null, fields)
    }))
  }
}
