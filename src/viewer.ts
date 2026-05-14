import Long from 'long'
import { PSTAttachment } from './PSTAttachment.class'
import { PSTFile } from './PSTFile.class'
import { PSTFolder } from './PSTFolder.class'
import { PSTMessage } from './PSTMessage.class'
import { PSTNodeInputStream } from './PSTNodeInputStream.class'
import { PSTUtil } from './PSTUtil.class'

export type MessageKind = 'mail' | 'appointment' | 'contact' | 'task' | 'activity' | 'other'

export interface ViewerStats {
  folderCount: number
  messageCount: number
  mailCount: number
  warningCount: number
}

export interface ViewerSessionIndex {
  fileName: string
  filePath: string
  mailboxName: string
  createdAt: string
  warnings: string[]
  stats: ViewerStats
  rootFolderId: string
  folders: Map<string, FolderSummary>
  messages: Map<string, MessageSummary>
}

export interface FolderSummary {
  id: string
  descriptorId: string
  parentId: string | null
  displayName: string
  path: string
  childFolderIds: string[]
  messageIds: string[]
  folderType: number
  contentCount: number
  unreadCount: number
  hasSubfolders: boolean
  containerClass: string
  containerFlags: number
  emailCount: number
  subFolderCount: number
  indexedMessageCount: number
  mailMessageCount: number
}

export interface MessageSummary {
  id: string
  descriptorId: string
  folderId: string
  folderPath: string
  order: number
  messageClass: string
  kind: MessageKind
  subject: string
  senderName: string
  senderEmailAddress: string
  recipientText: string
  displayTo: string
  displayCC: string
  displayBCC: string
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
  parseError?: string
}

export interface AttachmentDetail {
  attachmentId: string
  index: number
  filename: string
  longFilename: string
  downloadFilename: string
  mimeTag: string
  size: number
  attachMethod: number
  contentId: string
  pathname: string
  longPathname: string
  isEmbeddedMessage: boolean
  embeddedMessage: MessageDetail | null
  downloadUrl: string
  parseError?: string
}

export interface MessageDetail extends MessageSummary {
  sentRepresentingName: string
  sentRepresentingAddressType: string
  sentRepresentingEmailAddress: string
  receivedByName: string
  receivedByAddressType: string
  receivedByAddress: string
  replyRecipientNames: string
  originalDisplayTo: string
  originalDisplayCC: string
  originalDisplayBCC: string
  bodyPrefix: string
  bodyText: string
  bodyHtml: string
  bodyRtf: string
  transportMessageHeaders: string
  conversationTopic: string
  originalSubject: string
  internetMessageId: string
  inReplyToId: string
  returnPath: string
  attachments: AttachmentDetail[]
}

export interface FolderMessagePage {
  folder: FolderSummary
  items: MessageSummary[]
  total: number
  page: number
  pageSize: number
  totalPages: number
  query: string
  mailOnly: boolean
  sort: string
}

export interface FolderTreeNode {
  id: string
  descriptorId: string
  displayName: string
  path: string
  folderType: number
  contentCount: number
  unreadCount: number
  hasSubfolders: boolean
  containerClass: string
  containerFlags: number
  emailCount: number
  subFolderCount: number
  indexedMessageCount: number
  mailMessageCount: number
  childFolderIds: string[]
  messageIds: string[]
  children: FolderTreeNode[]
}

interface DetailBuildOptions {
  messageId?: string
  attachmentBaseUrl?: string
  embeddedDepth?: number
  folderIdOverride?: string
  folderPathOverride?: string
  virtualId?: string
}

interface MessageObjectWithMaybeAttachments extends PSTMessage {
  getAttachment(index: number): PSTAttachment
  numberOfAttachments: number
}

const MESSAGE_ID_PREFIX = 'message:'
const FOLDER_ID_PREFIX = 'folder:'

const MAIL_CLASS_PREFIXES = [
  'IPM.NOTE',
  'REPORT.IPM.NOTE',
  'IPM.POST.RSS',
  'IPM.DISTLIST',
  'IPM.STICKYNOTE',
  'IPM.NOTE.SMIME.MULTIPARTSIGNED',
  'IPM.NOTE.RULES.OOFTEMPLATE.MICROSOFT'
]

function isObjectLike(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeClassName(messageClass: string): string {
  return messageClass.trim().toUpperCase()
}

function safeString(value: unknown, fallback = ''): string {
  if (typeof value === 'string') {
    return value
  }
  if (value == null) {
    return fallback
  }
  return String(value)
}

function safeRead<T>(getter: () => T, fallback: T): T {
  try {
    const value = getter()
    return value === undefined || value === null ? fallback : value
  } catch {
    return fallback
  }
}

function toSafeNumber(value: unknown, fallback: number): number {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  if (typeof value === 'bigint') {
    return Number(value)
  }
  if (typeof value === 'string') {
    const parsed = Number(value)
    return Number.isFinite(parsed) ? parsed : fallback
  }
  if (
    isObjectLike(value) &&
    typeof (value as { toNumber?: unknown }).toNumber === 'function'
  ) {
    try {
      const parsed = (value as { toNumber: () => number }).toNumber()
      return Number.isFinite(parsed) ? parsed : fallback
    } catch {
      return fallback
    }
  }
  return fallback
}

function readFolderString(
  folder: PSTFolder,
  getter: (folder: PSTFolder) => unknown,
  fallback = ''
): string {
  return safeString(safeRead(() => getter(folder), fallback), fallback)
}

function readFolderNumber(
  folder: PSTFolder,
  getter: (folder: PSTFolder) => unknown,
  fallback = 0
): number {
  return toSafeNumber(safeRead(() => getter(folder), fallback), fallback)
}

function readFolderBoolean(
  folder: PSTFolder,
  getter: (folder: PSTFolder) => unknown,
  fallback = false
): boolean {
  return Boolean(safeRead(() => getter(folder), fallback))
}

function collapseWhitespace(value: string): string {
  return value.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim()
}

function dateToIso(value: Date | null | undefined): string | null {
  if (!value) {
    return null
  }
  return value.toISOString()
}

function dateToMs(value: string | null): number | null {
  if (!value) {
    return null
  }
  const ms = Date.parse(value)
  return Number.isNaN(ms) ? null : ms
}

function formatAddress(name: string, email: string): string {
  const cleanName = collapseWhitespace(name)
  const cleanEmail = collapseWhitespace(email)
  if (cleanName && cleanEmail && cleanName !== cleanEmail) {
    return `${cleanName} <${cleanEmail}>`
  }
  return cleanName || cleanEmail
}

function parseDescriptorId(value: string): Long {
  return Long.fromString(value)
}

function buildFolderId(descriptorId: string): string {
  return `${FOLDER_ID_PREFIX}${descriptorId}`
}

function buildMessageId(descriptorId: string): string {
  return `${MESSAGE_ID_PREFIX}${descriptorId}`
}

function buildVirtualId(messageId: string, attachmentIndex: number): string {
  return `${messageId}#embedded:${attachmentIndex}`
}

function safeFolderName(name: string, fallback: string): string {
  const cleaned = collapseWhitespace(name)
  return cleaned || fallback
}

function classifyMessageClass(messageClass: string): MessageKind {
  const normalized = normalizeClassName(messageClass)
  if (!normalized) {
    return 'other'
  }
  if (
    normalized.startsWith('IPM.CONTACT')
  ) {
    return 'contact'
  }
  if (normalized.startsWith('IPM.TASK')) {
    return 'task'
  }
  if (normalized.startsWith('IPM.ACTIVITY')) {
    return 'activity'
  }
  if (
    normalized.startsWith('IPM.APPOINTMENT') ||
    normalized.startsWith('IPM.SCHEDULE.MEETING')
  ) {
    return 'appointment'
  }
  if (
    normalized.startsWith('IPM.NOTE') ||
    normalized.startsWith('REPORT.IPM.NOTE') ||
    normalized.startsWith('IPM.POST.RSS') ||
    normalized.startsWith('IPM.DISTLIST') ||
    normalized.startsWith('IPM.STICKYNOTE')
  ) {
    return 'mail'
  }
  return 'other'
}

export function isMailLikeMessageClass(messageClass: string): boolean {
  const kind = classifyMessageClass(messageClass)
  return kind === 'mail'
}

function firstNonEmpty(...values: Array<string | null | undefined>): string {
  for (const value of values) {
    if (value && value.trim()) {
      return value
    }
  }
  return ''
}

function safeDate(value: Date | null | undefined): string | null {
  return dateToIso(value ?? null)
}

function buildSortDate(message: PSTMessage, summary: Partial<MessageSummary>): string | null {
  return (
    safeDate(message.clientSubmitTime) ||
    safeDate(message.messageDeliveryTime) ||
    safeDate(message.creationTime) ||
    safeDate(message.modificationTime) ||
    summary.sortDate ||
    null
  )
}

function normalizePreviewText(value: string): string {
  return collapseWhitespace(value).slice(0, 180)
}

function isPlainObjectWithId(value: unknown): value is { id: string } {
  return isObjectLike(value) && typeof value.id === 'string'
}

function getDescriptorIdString(value: Long | { toString(): string } | string): string {
  return typeof value === 'string' ? value : value.toString()
}

function buildSummaryFromMessage(
  message: PSTMessage,
  folderId: string,
  folderPath: string,
  order: number,
  messageId?: string,
  overrideFolderId?: string,
  overrideFolderPath?: string
): MessageSummary {
  const descriptorId = message.descriptorNodeId.toString()
  const id = messageId || buildMessageId(descriptorId)
  const messageClass = safeString(message.messageClass)
  const kind = classifyMessageClass(messageClass)
  const senderName = firstNonEmpty(
    safeString(message.senderName),
    safeString(message.sentRepresentingName),
    safeString(message.receivedByName)
  )
  const senderEmailAddress = firstNonEmpty(
    safeString(message.senderEmailAddress),
    safeString(message.sentRepresentingEmailAddress),
    safeString(message.emailAddress),
    safeString(message.receivedByAddress)
  )
  const displayTo = safeString(message.displayTo)
  const displayCC = safeString(message.displayCC)
  const displayBCC = safeString(message.displayBCC)
  const recipientText = [displayTo, displayCC, displayBCC]
    .map((value) => collapseWhitespace(value))
    .filter(Boolean)
    .join(' | ')
  const subject = collapseWhitespace(
    firstNonEmpty(
      safeString(message.subject),
      safeString(message.originalSubject),
      safeString(message.conversationTopic),
      '(no subject)'
    )
  )
  const clientSubmitTime = safeDate(message.clientSubmitTime)
  const creationTime = safeDate(message.creationTime)
  const modificationTime = safeDate(message.modificationTime)
  const messageDeliveryTime = safeDate(message.messageDeliveryTime)
  const sortDate = buildSortDate(message, {
    sortDate: null
  })
  const sortDateMs = dateToMs(sortDate)
  const hasAttachments = message.hasAttachments
  const isMailLike = kind === 'mail'

  return {
    id,
    descriptorId,
    folderId: overrideFolderId || folderId,
    folderPath: overrideFolderPath || folderPath,
    order,
    messageClass,
    kind,
    subject,
    senderName,
    senderEmailAddress,
    recipientText,
    displayTo,
    displayCC,
    displayBCC,
    clientSubmitTime,
    creationTime,
    modificationTime,
    messageDeliveryTime,
    sortDate,
    sortDateMs,
    importance: message.importance,
    hasAttachments,
    isRead: message.isRead,
    isMailLike
  }
}

function buildEmbeddedDetailId(parentId: string, attachmentIndex: number): string {
  return buildVirtualId(parentId, attachmentIndex)
}

function readNodeInputStreamToBuffer(stream: PSTNodeInputStream): Buffer {
  const chunks: Buffer[] = []
  const totalLength = stream.length.toNumber()
  if (totalLength <= 0) {
    return Buffer.alloc(0)
  }
  const chunkSize = 8192
  let remaining = totalLength
  while (remaining > 0) {
    const currentChunk = Buffer.alloc(Math.min(chunkSize, remaining))
    const bytesRead = stream.read(currentChunk)
    if (bytesRead <= 0) {
      break
    }
    chunks.push(currentChunk.subarray(0, bytesRead))
    remaining -= bytesRead
  }
  return Buffer.concat(chunks)
}

function htmlToText(html: string): string {
  if (!html) {
    return ''
  }
  const withoutScripts = html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
  const text = withoutScripts
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\r/g, '')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n\s*\n+/g, '\n\n')
    .trim()
}

function wrapBase64(base64: string, lineLength = 76): string {
  const lines: string[] = []
  for (let index = 0; index < base64.length; index += lineLength) {
    lines.push(base64.slice(index, index + lineLength))
  }
  return lines.join('\r\n')
}

function encodeHeaderValue(value: string): string {
  const sanitized = value.replace(/[\r\n]+/g, ' ').trim()
  if (!sanitized) {
    return ''
  }
  if (/^[\x00-\x7F]*$/.test(sanitized)) {
    return sanitized
  }
  const encoded = Buffer.from(sanitized, 'utf8').toString('base64')
  return `=?UTF-8?B?${encoded}?=`
}

function quoteHeaderValue(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')
}

function sanitizeFilename(name: string, fallback: string): string {
  const trimmed = collapseWhitespace(name)
    .replace(/[<>:"/\\|?*\u0000-\u001f]+/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
  const candidate = trimmed || fallback
  return candidate.slice(0, 180)
}

function guessMimeType(filename: string, fallback = 'application/octet-stream'): string {
  const extension = filename.toLowerCase().split('.').pop() || ''
  const table: Record<string, string> = {
    txt: 'text/plain',
    text: 'text/plain',
    csv: 'text/csv',
    html: 'text/html',
    htm: 'text/html',
    eml: 'message/rfc822',
    json: 'application/json',
    xml: 'application/xml',
    pdf: 'application/pdf',
    doc: 'application/msword',
    docx:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt: 'application/vnd.ms-powerpoint',
    pptx:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    png: 'image/png',
    gif: 'image/gif',
    svg: 'image/svg+xml',
    ics: 'text/calendar'
  }
  return table[extension] || fallback
}

function buildBodySection(detail: MessageDetail): string {
  const bodyText = detail.bodyText || htmlToText(detail.bodyHtml)
  const bodyHtml = detail.bodyHtml || ''
  const hasHtml = Boolean(bodyHtml.trim())
  const hasText = Boolean(bodyText.trim())

  if (hasHtml && hasText && bodyText !== htmlToText(bodyHtml)) {
    const boundary = `alt_${detail.id.replace(/[^a-zA-Z0-9]/g, '_')}`
    return [
      `Content-Type: multipart/alternative; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyText),
      '',
      `--${boundary}`,
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyHtml),
      '',
      `--${boundary}--`,
      ''
    ].join('\r\n')
  }

  if (hasHtml) {
    return [
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyHtml),
      ''
    ].join('\r\n')
  }

  return [
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: 8bit',
    '',
    normalizeLineEndings(bodyText || detail.parseError || ''),
    ''
  ].join('\r\n')
}

function normalizeLineEndings(value: string): string {
  return value.replace(/\r?\n/g, '\r\n')
}

function buildAttachmentPart(attachment: AttachmentDetail): string {
  const fileName = sanitizeFilename(
    attachment.downloadFilename || attachment.longFilename || attachment.filename,
    `attachment-${attachment.index}`
  )
  if (attachment.parseError) {
    const errorText = [
      `Attachment failed to load: ${attachment.parseError}`,
      `Original name: ${attachment.longFilename || attachment.filename || ''}`
    ]
      .filter(Boolean)
      .join('\n')
    return [
      `Content-Type: text/plain; charset=utf-8; name="${quoteHeaderValue(fileName)}.txt"`,
      'Content-Transfer-Encoding: 8bit',
      `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}.txt"`,
      '',
      normalizeLineEndings(errorText),
      ''
    ].join('\r\n')
  }

  if (attachment.embeddedMessage) {
    const nestedEml = buildMessageEmlFromDetail(attachment.embeddedMessage)
    return [
      `Content-Type: message/rfc822; name="${quoteHeaderValue(fileName)}"`,
      'Content-Transfer-Encoding: 8bit',
      `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}"`,
      attachment.contentId ? `Content-ID: <${attachment.contentId}>` : null,
      '',
      nestedEml
    ]
      .filter((line): line is string => line !== null)
      .join('\r\n')
  }

  const contentType =
    attachment.mimeTag && attachment.mimeTag.trim()
      ? attachment.mimeTag.trim()
      : guessMimeType(fileName)
  const body = attachment.downloadFilename
    ? `ATTACHMENT:${attachment.downloadFilename}`
    : ''

  if (!body) {
    return [
      `Content-Type: ${contentType}; name="${quoteHeaderValue(fileName)}"`,
      'Content-Transfer-Encoding: 8bit',
      `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}"`,
      attachment.contentId ? `Content-ID: <${attachment.contentId}>` : null,
      '',
      'Attachment content unavailable in offline export.',
      ''
    ]
      .filter((line): line is string => line !== null)
      .join('\r\n')
  }

  // The raw binary payload is only emitted by the server download route.
  return [
    `Content-Type: ${contentType}; name="${quoteHeaderValue(fileName)}"`,
    'Content-Transfer-Encoding: base64',
    `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}"`,
    attachment.contentId ? `Content-ID: <${attachment.contentId}>` : null,
    '',
    wrapBase64(Buffer.from(body, 'utf8').toString('base64')),
    ''
  ]
    .filter((line): line is string => line !== null)
    .join('\r\n')
}

export function buildMessageEmlFromDetail(detail: MessageDetail): string {
  const headers: string[] = []
  const fromAddress = formatAddress(detail.senderName, detail.senderEmailAddress)
  const toAddress = detail.displayTo.trim()
  const ccAddress = detail.displayCC.trim()
  const bccAddress = detail.displayBCC.trim()
  const subject = encodeHeaderValue(detail.subject || '(no subject)')
  const dateHeader = detail.clientSubmitTime || detail.messageDeliveryTime || detail.creationTime
  const messageId = detail.internetMessageId || ''

  headers.push(`From: ${encodeHeaderValue(fromAddress)}`)
  if (toAddress) {
    headers.push(`To: ${encodeHeaderValue(toAddress)}`)
  }
  if (ccAddress) {
    headers.push(`Cc: ${encodeHeaderValue(ccAddress)}`)
  }
  if (bccAddress) {
    headers.push(`Bcc: ${encodeHeaderValue(bccAddress)}`)
  }
  headers.push(`Subject: ${subject}`)
  if (dateHeader) {
    headers.push(`Date: ${dateHeader}`)
  }
  if (messageId) {
    headers.push(`Message-ID: ${messageId}`)
  }
  headers.push('MIME-Version: 1.0')
  headers.push(`X-PST-Message-Class: ${encodeHeaderValue(detail.messageClass)}`)
  headers.push(`X-PST-Folder: ${encodeHeaderValue(detail.folderPath)}`)

  const attachmentParts = detail.attachments
  const hasAttachments = attachmentParts.length > 0
  const bodySection = buildBodySection(detail)

  if (!hasAttachments) {
    return [...headers, '', bodySection].join('\r\n')
  }

  const boundary = `mixed_${detail.id.replace(/[^a-zA-Z0-9]/g, '_')}`
  const renderedParts = attachmentParts.map((attachment) => buildAttachmentPart(attachment))
  const bodyLines = [
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    bodySection,
    ...renderedParts.flatMap((part) => [`--${boundary}`, part]),
    `--${boundary}--`,
    ''
  ]
  return [...headers, ...bodyLines].join('\r\n')
}

function buildAttachmentEmlFromMessage(
  attachment: PSTAttachment,
  index: number,
  messageId: string,
  embeddedDepth: number,
  summary: MessageSummary
): string {
  const fileName = sanitizeFilename(
    safeString(attachment.longFilename) ||
      safeString(attachment.filename) ||
      `attachment-${index}`,
    `attachment-${index}`
  )
  const embedded = attachment.embeddedPSTMessage
  if (embedded && embeddedDepth > 0) {
    const embeddedSummary = buildSummaryFromMessage(
      embedded,
      summary.folderId,
      summary.folderPath,
      -1,
      buildEmbeddedDetailId(messageId, index),
      summary.folderId,
      summary.folderPath
    )
    const nestedEml = buildMessageEmlFromMessage(embedded, embeddedSummary, {
      embeddedDepth: embeddedDepth - 1
    })
    return [
      `Content-Type: message/rfc822; name="${quoteHeaderValue(fileName)}"`,
      'Content-Transfer-Encoding: 8bit',
      `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}"`,
      safeString(attachment.contentId) ? `Content-ID: <${safeString(attachment.contentId)}>` : null,
      '',
      nestedEml,
      ''
    ]
      .filter((line): line is string => line !== null)
      .join('\r\n')
  }

  const contentType =
    safeString(attachment.mimeTag).trim() || guessMimeType(fileName)
  const stream = attachment.fileInputStream
  const rawBuffer = stream ? readNodeInputStreamToBuffer(stream) : Buffer.alloc(0)
  const headers = [
    `Content-Type: ${contentType}; name="${quoteHeaderValue(fileName)}"`,
    'Content-Transfer-Encoding: base64',
    `Content-Disposition: attachment; filename="${quoteHeaderValue(fileName)}"`,
    safeString(attachment.contentId) ? `Content-ID: <${safeString(attachment.contentId)}>` : null,
    ''
  ].filter((line): line is string => line !== null)

  if (rawBuffer.length === 0) {
    headers.push(
      wrapBase64(Buffer.from('Attachment content unavailable', 'utf8').toString('base64')),
      ''
    )
    return headers.join('\r\n')
  }

  headers.push(wrapBase64(rawBuffer.toString('base64')), '')
  return headers.join('\r\n')
}

export function buildMessageEmlFromMessage(
  message: PSTMessage,
  summary: MessageSummary,
  options: { embeddedDepth?: number } = {}
): string {
  const embeddedDepth = options.embeddedDepth ?? 1
  const headers: string[] = []
  const fromAddress = formatAddress(
    safeString(message.senderName) || safeString(message.sentRepresentingName),
    firstNonEmpty(
      safeString(message.senderEmailAddress),
      safeString(message.sentRepresentingEmailAddress),
      safeString(message.emailAddress)
    )
  )
  const toAddress = safeString(message.displayTo).trim()
  const ccAddress = safeString(message.displayCC).trim()
  const bccAddress = safeString(message.displayBCC).trim()
  const subject = encodeHeaderValue(summary.subject || '(no subject)')
  const dateHeader =
    summary.clientSubmitTime ||
    summary.messageDeliveryTime ||
    summary.creationTime ||
    summary.modificationTime
  const messageId =
    safeString(message.internetMessageId) ||
    `<${summary.descriptorId}.${summary.id}@pst-explorer>`

  if (fromAddress) {
    headers.push(`From: ${encodeHeaderValue(fromAddress)}`)
  }
  if (toAddress) {
    headers.push(`To: ${encodeHeaderValue(toAddress)}`)
  }
  if (ccAddress) {
    headers.push(`Cc: ${encodeHeaderValue(ccAddress)}`)
  }
  if (bccAddress) {
    headers.push(`Bcc: ${encodeHeaderValue(bccAddress)}`)
  }
  headers.push(`Subject: ${subject}`)
  if (dateHeader) {
    headers.push(`Date: ${dateHeader}`)
  }
  headers.push(`Message-ID: ${messageId}`)
  headers.push('MIME-Version: 1.0')
  headers.push(`X-PST-Message-Class: ${encodeHeaderValue(summary.messageClass)}`)
  headers.push(`X-PST-Folder: ${encodeHeaderValue(summary.folderPath)}`)

  let bodyText = safeString(message.body)
  let bodyHtml = safeString(message.bodyHTML)
  const hasHtml = Boolean(bodyHtml.trim())
  const hasText = Boolean(bodyText.trim())
  const attachments: string[] = []

  const attachmentCount = (() => {
    try {
      return (message as MessageObjectWithMaybeAttachments).numberOfAttachments || 0
    } catch (err) {
      return 0
    }
  })()

  for (let index = 0; index < attachmentCount; index++) {
    try {
      const attachment = (message as MessageObjectWithMaybeAttachments).getAttachment(index)
      attachments.push(
        buildAttachmentEmlFromMessage(attachment, index, summary.id, embeddedDepth, summary)
      )
    } catch (err) {
      const errorText = err instanceof Error ? err.message : String(err)
      attachments.push(
        [
          'Content-Type: text/plain; charset=utf-8',
          'Content-Transfer-Encoding: 8bit',
          `Content-Disposition: attachment; filename="attachment-${index}-error.txt"`,
          '',
          `Unable to export attachment ${index}: ${errorText}`,
          ''
        ].join('\r\n')
      )
    }
  }

  const hasAttachments = attachments.length > 0

  if (hasAttachments) {
    const boundary = `mixed_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}`
    const mixedLines = [
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      ...(hasHtml && hasText && bodyText !== htmlToText(bodyHtml)
        ? [
            `Content-Type: multipart/alternative; boundary="alt_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}"`,
            '',
            `--alt_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}`,
            'Content-Type: text/plain; charset=utf-8',
            'Content-Transfer-Encoding: 8bit',
            '',
            normalizeLineEndings(bodyText),
            '',
            `--alt_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}`,
            'Content-Type: text/html; charset=utf-8',
            'Content-Transfer-Encoding: 8bit',
            '',
            normalizeLineEndings(bodyHtml),
            '',
            `--alt_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}--`,
            ''
          ]
        : hasHtml
          ? [
              'Content-Type: text/html; charset=utf-8',
              'Content-Transfer-Encoding: 8bit',
              '',
              normalizeLineEndings(bodyHtml),
              ''
            ]
          : [
              'Content-Type: text/plain; charset=utf-8',
              'Content-Transfer-Encoding: 8bit',
              '',
              normalizeLineEndings(bodyText || summary.parseError || ''),
              ''
            ]),
      ...attachments.flatMap((part) => [`--${boundary}`, part]),
      `--${boundary}--`,
      ''
    ]
    return [...headers, ...mixedLines].join('\r\n')
  }

  if (hasHtml && hasText && bodyText !== htmlToText(bodyHtml)) {
    const boundary = `alt_${summary.id.replace(/[^a-zA-Z0-9]/g, '_')}`
    return [
      ...headers,
      `Content-Type: multipart/alternative; boundary="${boundary}"`,
      '',
      `--${boundary}`,
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyText),
      '',
      `--${boundary}`,
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyHtml),
      '',
      `--${boundary}--`,
      ''
    ].join('\r\n')
  }

  if (hasHtml) {
    return [
      ...headers,
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      normalizeLineEndings(bodyHtml),
      ''
    ].join('\r\n')
  }

  return [
    ...headers,
    'Content-Type: text/plain; charset=utf-8',
    'Content-Transfer-Encoding: 8bit',
    '',
    normalizeLineEndings(bodyText || summary.parseError || ''),
    ''
  ].join('\r\n')
}

function buildAttachmentSummary(
  messageId: string,
  attachment: PSTAttachment,
  index: number,
  embeddedDepth: number,
  detailFolderId: string,
  detailFolderPath: string
): AttachmentDetail {
  const attachmentId = `${messageId}:attachment:${index}`
  const filename = safeString(attachment.filename)
  const longFilename = safeString(attachment.longFilename)
  const longPathname = safeString(attachment.longPathname)
  const pathname = safeString(attachment.pathname)
  const downloadFilename = sanitizeFilename(
    longFilename || filename || longPathname || pathname || `attachment-${index}`,
    `attachment-${index}`
  )

  const fileSize = toSafeNumber(safeRead(() => attachment.filesize, 0), 0)
  const size = fileSize > 0 ? fileSize : toSafeNumber(safeRead(() => attachment.size, 0), 0)

  let embeddedMessage: MessageDetail | null = null
  let parseError: string | undefined
  let isEmbeddedMessage = false

  try {
    const embedded = attachment.embeddedPSTMessage
    if (embedded) {
      isEmbeddedMessage = true
      if (embeddedDepth > 0) {
        const embeddedSummary = buildSummaryFromMessage(
          embedded,
          detailFolderId,
          detailFolderPath,
          -1,
          buildEmbeddedDetailId(messageId, index),
          detailFolderId,
          detailFolderPath
        )
        embeddedMessage = buildMessageDetail(embedded, embeddedSummary, {
          embeddedDepth: embeddedDepth - 1,
          folderIdOverride: detailFolderId,
          folderPathOverride: detailFolderPath,
          virtualId: buildEmbeddedDetailId(messageId, index)
        })
      }
    }
  } catch (err) {
    parseError = err instanceof Error ? err.message : String(err)
  }

  return {
    attachmentId,
    index,
    filename,
    longFilename,
    downloadFilename,
    mimeTag: safeString(attachment.mimeTag),
    size,
    attachMethod: attachment.attachMethod,
    contentId: safeString(attachment.contentId),
    pathname,
    longPathname,
    isEmbeddedMessage,
    embeddedMessage,
    downloadUrl: '',
    parseError
  }
}

function buildMessageDetailFromMessage(
  message: PSTMessage,
  summary: MessageSummary,
  options: DetailBuildOptions = {}
): MessageDetail {
  const messageId = options.messageId || summary.id
  const embeddedDepth = options.embeddedDepth ?? 1
  const folderId = options.folderIdOverride || summary.folderId
  const folderPath = options.folderPathOverride || summary.folderPath
  const attachmentBaseUrl = options.attachmentBaseUrl || ''
  const detail: MessageDetail = {
    ...summary,
    id: messageId,
    folderId,
    folderPath,
    sentRepresentingName: safeString(message.sentRepresentingName),
    sentRepresentingAddressType: safeString(message.sentRepresentingAddressType),
    sentRepresentingEmailAddress: safeString(message.sentRepresentingEmailAddress),
    receivedByName: safeString(message.receivedByName),
    receivedByAddressType: safeString(message.receivedByAddressType),
    receivedByAddress: safeString(message.receivedByAddress),
    replyRecipientNames: safeString(message.replyRecipientNames),
    originalDisplayTo: safeString(message.originalDisplayTo),
    originalDisplayCC: safeString(message.originalDisplayCc),
    originalDisplayBCC: safeString(message.originalDisplayBcc),
    bodyPrefix: safeString(message.bodyPrefix),
    bodyText: safeString(message.body),
    bodyHtml: safeString(message.bodyHTML),
    bodyRtf: safeString(message.bodyRTF),
    transportMessageHeaders: safeString(message.transportMessageHeaders),
    conversationTopic: safeString(message.conversationTopic),
    originalSubject: safeString(message.originalSubject),
    internetMessageId: safeString(message.internetMessageId),
    inReplyToId: safeString(message.inReplyToId),
    returnPath: safeString(message.returnPath),
    attachments: []
  }

  const attachmentCount = (() => {
    try {
      return (message as MessageObjectWithMaybeAttachments).numberOfAttachments || 0
    } catch (err) {
      return 0
    }
  })()

  for (let index = 0; index < attachmentCount; index++) {
    let attachmentDetail: AttachmentDetail
    try {
      const attachment = (message as MessageObjectWithMaybeAttachments).getAttachment(index)
      attachmentDetail = buildAttachmentSummary(
        messageId,
        attachment,
        index,
        embeddedDepth,
        folderId,
        folderPath
      )
      if (attachmentBaseUrl) {
        attachmentDetail.downloadUrl = `${attachmentBaseUrl}${index}`
      }
    } catch (err) {
      attachmentDetail = {
        attachmentId: `${messageId}:attachment:${index}`,
        index,
        filename: '',
        longFilename: '',
        downloadFilename: `attachment-${index}`,
        mimeTag: '',
        size: 0,
        attachMethod: 0,
        contentId: '',
        pathname: '',
        longPathname: '',
        isEmbeddedMessage: false,
        embeddedMessage: null,
        downloadUrl: attachmentBaseUrl ? `${attachmentBaseUrl}${index}` : '',
        parseError: err instanceof Error ? err.message : String(err)
      }
    }
    detail.attachments.push(attachmentDetail)
  }

  return detail
}

export function buildEmptyMessageDetail(summary: MessageSummary): MessageDetail {
  return {
    ...summary,
    sentRepresentingName: '',
    sentRepresentingAddressType: '',
    sentRepresentingEmailAddress: '',
    receivedByName: '',
    receivedByAddressType: '',
    receivedByAddress: '',
    replyRecipientNames: '',
    originalDisplayTo: '',
    originalDisplayCC: '',
    originalDisplayBCC: '',
    bodyPrefix: '',
    bodyText: '',
    bodyHtml: '',
    bodyRtf: '',
    transportMessageHeaders: '',
    conversationTopic: '',
    originalSubject: '',
    internetMessageId: '',
    inReplyToId: '',
    returnPath: '',
    attachments: []
  }
}

export function isMailLikeSummary(summary: MessageSummary): boolean {
  return summary.isMailLike
}

export function messageMatchesQuery(summary: MessageSummary, query: string): boolean {
  if (!query) {
    return true
  }
  const haystack = [
    summary.subject,
    summary.senderName,
    summary.senderEmailAddress,
    summary.displayTo,
    summary.displayCC,
    summary.displayBCC,
    summary.recipientText,
    summary.messageClass,
    summary.kind
  ]
    .map((value) => value.toLowerCase())
    .join(' ')
  return haystack.includes(query.toLowerCase())
}

export function sortMessageSummaries(messages: MessageSummary[], sort: string): MessageSummary[] {
  const output = [...messages]
  if (sort === 'order') {
    return output.sort((a, b) => a.order - b.order)
  }
  return output.sort((a, b) => {
    const aDate = a.sortDateMs ?? Number.MIN_SAFE_INTEGER
    const bDate = b.sortDateMs ?? Number.MIN_SAFE_INTEGER
    if (aDate !== bDate) {
      return bDate - aDate
    }
    return a.order - b.order
  })
}

function makeFolderSummaryCopy(folder: FolderSummary): FolderSummary {
  return {
    ...folder,
    childFolderIds: [...folder.childFolderIds],
    messageIds: [...folder.messageIds]
  }
}

function buildFolderTreeNode(session: ViewerSessionIndex, folderId: string): FolderTreeNode {
  const folder = session.folders.get(folderId)
  if (!folder) {
    throw new Error(`Unknown folder: ${folderId}`)
  }
  const children = folder.childFolderIds.map((childId) => buildFolderTreeNode(session, childId))
  return {
    id: folder.id,
    descriptorId: folder.descriptorId,
    displayName: folder.displayName,
    path: folder.path,
    folderType: folder.folderType,
    contentCount: folder.contentCount,
    unreadCount: folder.unreadCount,
    hasSubfolders: folder.hasSubfolders,
    containerClass: folder.containerClass,
    containerFlags: folder.containerFlags,
    emailCount: folder.emailCount,
    subFolderCount: folder.subFolderCount,
    indexedMessageCount: folder.indexedMessageCount,
    mailMessageCount: folder.mailMessageCount,
    childFolderIds: [...folder.childFolderIds],
    messageIds: [...folder.messageIds],
    children
  }
}

function safeLoadSubFolders(folder: PSTFolder, warnings: string[], folderPath: string): PSTFolder[] {
  try {
    return folder.getSubFolders()
  } catch (err) {
    warnings.push(
      `Unable to load subfolders for "${folderPath}": ${err instanceof Error ? err.message : String(err)}`
    )
    return []
  }
}

function loadNextChild(folder: PSTFolder): PSTMessage | null {
  const child = folder.getNextChild()
  if (!child) {
    return null
  }
  if (child instanceof PSTMessage) {
    return child
  }
  if (isPlainObjectWithId(child) && 'messageClass' in child) {
    return child as unknown as PSTMessage
  }
  return null
}

function indexFolder(
  session: ViewerSessionIndex,
  folder: PSTFolder,
  parentFolderId: string | null,
  parentPath: string,
  visited: Set<string>
): void {
  const descriptorId = safeRead(() => folder.descriptorNodeId.toString(), '')
  const folderId = buildFolderId(descriptorId || `${session.folders.size}`)
  if (visited.has(folderId)) {
    session.warnings.push(`Skipping repeated folder ${folderId} to avoid a loop.`)
    return
  }
  visited.add(folderId)

  const displayName = safeFolderName(
    readFolderString(folder, (value) => value.displayName, parentFolderId ? 'Untitled folder' : session.mailboxName || 'Root'),
    parentFolderId ? 'Untitled folder' : session.mailboxName || 'Root'
  )
  const pathName = parentPath ? `${parentPath}/${displayName}` : displayName

  const folderSummary: FolderSummary = {
    id: folderId,
    descriptorId,
    parentId: parentFolderId,
    displayName,
    path: pathName,
    childFolderIds: [],
    messageIds: [],
    folderType: readFolderNumber(folder, (value) => value.folderType, 0),
    contentCount: readFolderNumber(folder, (value) => value.contentCount, 0),
    unreadCount: readFolderNumber(folder, (value) => value.unreadCount, 0),
    hasSubfolders: readFolderBoolean(folder, (value) => value.hasSubfolders, false),
    containerClass: readFolderString(folder, (value) => value.containerClass, ''),
    containerFlags: readFolderNumber(folder, (value) => value.containerFlags, 0),
    emailCount: readFolderNumber(folder, (value) => value.emailCount, 0),
    subFolderCount: readFolderNumber(folder, (value) => value.subFolderCount, 0),
    indexedMessageCount: 0,
    mailMessageCount: 0
  }

  session.folders.set(folderId, folderSummary)
  if (parentFolderId) {
    const parent = session.folders.get(parentFolderId)
    if (parent) {
      parent.childFolderIds.push(folderId)
    }
  } else {
    session.rootFolderId = folderId
  }

  const childFolders = safeLoadSubFolders(folder, session.warnings, pathName)
  for (const childFolder of childFolders) {
    indexFolder(session, childFolder, folderId, pathName, visited)
  }

  try {
    folder.moveChildCursorTo(0)
  } catch (err) {
    session.warnings.push(
      `Unable to reset cursor for folder "${pathName}": ${err instanceof Error ? err.message : String(err)}`
    )
  }

  let order = 0
  while (true) {
    let child: PSTMessage | null = null
    try {
      child = loadNextChild(folder)
    } catch (err) {
      session.warnings.push(
        `Unable to read child message in "${pathName}": ${err instanceof Error ? err.message : String(err)}`
      )
      break
    }

    if (!child) {
      break
    }

    const childDescriptorId = safeRead(() => child.descriptorNodeId.toString(), '')
    const messageId = buildMessageId(childDescriptorId || `${folderId}:${order}`)
    try {
      const summary = buildSummaryFromMessage(child, folderId, pathName, order, messageId)
      session.messages.set(messageId, summary)
      folderSummary.messageIds.push(messageId)
      folderSummary.indexedMessageCount += 1
      if (summary.isMailLike) {
        folderSummary.mailMessageCount += 1
      }
    } catch (err) {
      const parseError = err instanceof Error ? err.message : String(err)
      session.warnings.push(
        `Unable to index message ${messageId} in "${pathName}": ${parseError}`
      )
      const fallbackSummary: MessageSummary = {
        id: messageId,
        descriptorId: childDescriptorId,
        folderId,
        folderPath: pathName,
        order,
        messageClass: safeString(safeRead(() => child.messageClass, ''), ''),
        kind: classifyMessageClass(safeString(safeRead(() => child.messageClass, ''), '')),
        subject: '(unavailable)',
        senderName: '',
        senderEmailAddress: '',
        recipientText: '',
        displayTo: '',
        displayCC: '',
        displayBCC: '',
        clientSubmitTime: null,
        creationTime: null,
        modificationTime: null,
        messageDeliveryTime: null,
        sortDate: null,
        sortDateMs: null,
        importance: 0,
        hasAttachments: false,
        isRead: false,
        isMailLike: false,
        parseError
      }
      session.messages.set(messageId, fallbackSummary)
      folderSummary.messageIds.push(messageId)
      folderSummary.indexedMessageCount += 1
    }
    order += 1
  }
}

function computeStats(session: ViewerSessionIndex): ViewerStats {
  const folders = [...session.folders.values()]
  const messages = [...session.messages.values()]
  return {
    folderCount: folders.length,
    messageCount: messages.length,
    mailCount: messages.filter((item) => item.isMailLike).length,
    warningCount: session.warnings.length
  }
}

function withPstFile<T>(filePath: string, callback: (pstFile: PSTFile) => T): T {
  const pstFile = new PSTFile(filePath)
  try {
    return callback(pstFile)
  } finally {
    pstFile.close()
  }
}

function loadMessageById(
  pstFile: PSTFile,
  summary: MessageSummary
): PSTMessage {
  return PSTUtil.detectAndLoadPSTObject(
    pstFile,
    parseDescriptorId(summary.descriptorId)
  ) as PSTMessage
}

export function createViewerSession(filePath: string, fileName: string): ViewerSessionIndex {
  return withPstFile(filePath, (pstFile) => {
    const messageStore = pstFile.getMessageStore()
    const mailboxName = safeFolderName(
      messageStore.displayName,
      safeString(fileName, 'Mailbox')
    )
    const session: ViewerSessionIndex = {
      fileName,
      filePath,
      mailboxName,
      createdAt: new Date().toISOString(),
      warnings: [],
      stats: {
        folderCount: 0,
        messageCount: 0,
        mailCount: 0,
        warningCount: 0
      },
      rootFolderId: '',
      folders: new Map<string, FolderSummary>(),
      messages: new Map<string, MessageSummary>()
    }

    const rootFolder = pstFile.getRootFolder()
    indexFolder(session, rootFolder, null, mailboxName, new Set<string>())
    session.stats = computeStats(session)
    return session
  })
}

export function buildSessionSummary(session: ViewerSessionIndex) {
  return {
    fileName: session.fileName,
    mailboxName: session.mailboxName,
    createdAt: session.createdAt,
    rootFolderId: session.rootFolderId,
    stats: session.stats,
    warningCount: session.warnings.length,
    warnings: session.warnings.slice(0, 10)
  }
}

export function buildFolderTree(session: ViewerSessionIndex): FolderTreeNode {
  if (!session.rootFolderId) {
    throw new Error('Session does not contain a root folder')
  }
  return buildFolderTreeNode(session, session.rootFolderId)
}

export function getFolderSummary(session: ViewerSessionIndex, folderId: string): FolderSummary {
  const folder = session.folders.get(folderId)
  if (!folder) {
    throw new Error(`Unknown folder: ${folderId}`)
  }
  return makeFolderSummaryCopy(folder)
}

export function getMessageSummary(session: ViewerSessionIndex, messageId: string): MessageSummary {
  const message = session.messages.get(messageId)
  if (!message) {
    throw new Error(`Unknown message: ${messageId}`)
  }
  return { ...message }
}

export function listFolderMessages(
  session: ViewerSessionIndex,
  folderId: string,
  options: {
    query?: string
    mailOnly?: boolean
    page?: number
    pageSize?: number
    sort?: string
  } = {}
): FolderMessagePage {
  const folder = getFolderSummary(session, folderId)
  const query = safeString(options.query).trim()
  const mailOnly = options.mailOnly !== false
  const sort = options.sort || 'date-desc'
  const pageSize = Math.min(Math.max(options.pageSize || 50, 1), 200)
  const messages = folder.messageIds
    .map((messageId) => session.messages.get(messageId))
    .filter((message): message is MessageSummary => Boolean(message))
    .filter((message) => (mailOnly ? isMailLikeSummary(message) : true))
    .filter((message) => messageMatchesQuery(message, query))
  const sortedMessages = sortMessageSummaries(messages, sort)
  const total = sortedMessages.length
  const totalPages = Math.max(1, Math.ceil(total / pageSize))
  const page = Math.min(Math.max(options.page || 1, 1), totalPages)
  const start = (page - 1) * pageSize
  const items = sortedMessages.slice(start, start + pageSize).map((item) => ({ ...item }))
  return {
    folder,
    items,
    total,
    page,
    pageSize,
    totalPages,
    query,
    mailOnly,
    sort
  }
}

export function buildMessageDetailFromSession(
  session: ViewerSessionIndex,
  messageId: string,
  embeddedDepth = 1
): MessageDetail {
  const summary = getMessageSummary(session, messageId)
  if (summary.parseError) {
    return {
      ...summary,
      sentRepresentingName: '',
      sentRepresentingAddressType: '',
      sentRepresentingEmailAddress: '',
      receivedByName: '',
      receivedByAddressType: '',
      receivedByAddress: '',
      replyRecipientNames: '',
      originalDisplayTo: '',
      originalDisplayCC: '',
      originalDisplayBCC: '',
      bodyPrefix: '',
      bodyText: '',
      bodyHtml: '',
      bodyRtf: '',
      transportMessageHeaders: '',
      conversationTopic: '',
      originalSubject: '',
      internetMessageId: '',
      inReplyToId: '',
      returnPath: '',
      attachments: []
    }
  }

  try {
    return withSessionMessage(session, messageId, (message) => {
      return buildMessageDetail(message, summary, {
        messageId,
        attachmentBaseUrl: '',
        embeddedDepth
      })
    })
  } catch (err) {
    const parseError = err instanceof Error ? err.message : String(err)
    return buildEmptyMessageDetail({
      ...summary,
      parseError
    })
  }
}

export function withSessionMessage<T>(
  session: ViewerSessionIndex,
  messageId: string,
  callback: (message: PSTMessage, summary: MessageSummary) => T
): T {
  const summary = getMessageSummary(session, messageId)
  return withPstFile(session.filePath, (pstFile) => {
    const message = loadMessageById(pstFile, summary)
    return callback(message, summary)
  })
}

export function buildMessageDetail(
  message: PSTMessage,
  summary: MessageSummary,
  options: DetailBuildOptions = {}
): MessageDetail {
  const messageId = options.messageId || summary.id
  const embeddedDepth = options.embeddedDepth ?? 1
  const folderId = options.folderIdOverride || summary.folderId
  const folderPath = options.folderPathOverride || summary.folderPath
  const attachmentBaseUrl = options.attachmentBaseUrl || ''
  const detail: MessageDetail = {
    ...summary,
    id: messageId,
    folderId,
    folderPath,
    sentRepresentingName: safeString(message.sentRepresentingName),
    sentRepresentingAddressType: safeString(message.sentRepresentingAddressType),
    sentRepresentingEmailAddress: safeString(message.sentRepresentingEmailAddress),
    receivedByName: safeString(message.receivedByName),
    receivedByAddressType: safeString(message.receivedByAddressType),
    receivedByAddress: safeString(message.receivedByAddress),
    replyRecipientNames: safeString(message.replyRecipientNames),
    originalDisplayTo: safeString(message.originalDisplayTo),
    originalDisplayCC: safeString(message.originalDisplayCc),
    originalDisplayBCC: safeString(message.originalDisplayBcc),
    bodyPrefix: safeString(message.bodyPrefix),
    bodyText: safeString(message.body),
    bodyHtml: safeString(message.bodyHTML),
    bodyRtf: safeString(message.bodyRTF),
    transportMessageHeaders: safeString(message.transportMessageHeaders),
    conversationTopic: safeString(message.conversationTopic),
    originalSubject: safeString(message.originalSubject),
    internetMessageId: safeString(message.internetMessageId),
    inReplyToId: safeString(message.inReplyToId),
    returnPath: safeString(message.returnPath),
    attachments: []
  }

  let attachmentCount = 0
  try {
    attachmentCount = (message as MessageObjectWithMaybeAttachments).numberOfAttachments || 0
  } catch (err) {
    attachmentCount = 0
  }

  for (let index = 0; index < attachmentCount; index++) {
    let attachmentDetail: AttachmentDetail
    try {
      const attachment = (message as MessageObjectWithMaybeAttachments).getAttachment(index)
      attachmentDetail = buildAttachmentSummary(
        messageId,
        attachment,
        index,
        embeddedDepth,
        folderId,
        folderPath
      )
      if (attachmentBaseUrl) {
        attachmentDetail.downloadUrl = `${attachmentBaseUrl}${index}`
      }
    } catch (err) {
      attachmentDetail = {
        attachmentId: `${messageId}:attachment:${index}`,
        index,
        filename: '',
        longFilename: '',
        downloadFilename: `attachment-${index}`,
        mimeTag: '',
        size: 0,
        attachMethod: 0,
        contentId: '',
        pathname: '',
        longPathname: '',
        isEmbeddedMessage: false,
        embeddedMessage: null,
        downloadUrl: attachmentBaseUrl ? `${attachmentBaseUrl}${index}` : '',
        parseError: err instanceof Error ? err.message : String(err)
      }
    }
    detail.attachments.push(attachmentDetail)
  }

  return detail
}

export function exportMessageAsJson(detail: MessageDetail): string {
  return JSON.stringify(detail, null, 2)
}

export function exportMessageAsEml(detail: MessageDetail): string {
  return buildMessageEmlFromDetail(detail)
}

export function exportMessageAsEmlFromSession(
  session: ViewerSessionIndex,
  messageId: string,
  embeddedDepth = 1
): string {
  return withSessionMessage(session, messageId, (message, summary) =>
    buildMessageEmlFromMessage(message, summary, { embeddedDepth })
  )
}

export function getAttachmentDownloadBuffer(
  session: ViewerSessionIndex,
  messageId: string,
  attachmentIndex: number
): {
  filename: string
  contentType: string
  buffer: Buffer
} {
  return withSessionMessage(session, messageId, (message, summary) => {
    const attachment = (message as MessageObjectWithMaybeAttachments).getAttachment(attachmentIndex)
    const fileName = sanitizeFilename(
      safeString(attachment.longFilename) ||
        safeString(attachment.filename) ||
        `attachment-${attachmentIndex}`,
      `attachment-${attachmentIndex}`
    )
    const embedded = attachment.embeddedPSTMessage
    if (embedded) {
      const embeddedSummary = buildSummaryFromMessage(
        embedded,
        summary.folderId,
        summary.folderPath,
        -1,
        buildEmbeddedDetailId(messageId, attachmentIndex),
        summary.folderId,
        summary.folderPath
      )
      return {
        filename: `${fileName}.eml`,
        contentType: 'message/rfc822',
        buffer: Buffer.from(
          buildMessageEmlFromMessage(embedded, embeddedSummary, { embeddedDepth: 0 }),
          'utf8'
        )
      }
    }
    const stream = attachment.fileInputStream
    if (!stream) {
      return {
        filename: fileName,
        contentType: attachment.mimeTag || guessMimeType(fileName),
        buffer: Buffer.alloc(0)
      }
    }
    return {
      filename: fileName,
      contentType: attachment.mimeTag || guessMimeType(fileName),
      buffer: readNodeInputStreamToBuffer(stream)
    }
  })
}

export function getMessageDetail(
  session: ViewerSessionIndex,
  messageId: string,
  embeddedDepth = 1
): MessageDetail {
  return withSessionMessage(session, messageId, (message, summary) =>
    buildMessageDetail(message, summary, {
      messageId,
      attachmentBaseUrl: '',
      embeddedDepth
    })
  )
}

export function resolveSessionMessageCount(session: ViewerSessionIndex): number {
  return session.messages.size
}

export function resolveSessionFolderCount(session: ViewerSessionIndex): number {
  return session.folders.size
}

export function resolveSessionMailCount(session: ViewerSessionIndex): number {
  return [...session.messages.values()].filter((item) => item.isMailLike).length
}

export function sanitizeFileNameForDownload(name: string, fallback: string): string {
  return sanitizeFilename(name, fallback)
}
