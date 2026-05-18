import express from 'express'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import { API_ROUTES } from './apiRoutes'
import {
  buildFolderExtractionPage,
  buildMessageExtractionRecord,
  buildSummaryExtractionRecord,
  isSummaryOnlyExtraction,
  normalizeExtractionFields,
  type ExtractionFieldGroup
} from './extraction'
import {
  buildReviewContext,
  type ReviewStore
} from './reviewStore'
import type { ReviewState } from './reviewTypes'
import {
  buildSearchIndexDocumentsFromSession,
  refreshSearchIndexFromCatalog,
  type HiddenRuleRecord,
  type SearchIndexPage,
  type SearchIndexStore,
  type SearchScope
} from './searchIndex'
import {
  buildEmptyMessageDetail,
  collectFolderMessages,
  buildFolderTree,
  buildMessageDetail,
  buildMessageDetailFromSession,
  buildSessionSummary,
  exportAppointmentAsIcsFromSession,
  exportMessageAsEml,
  exportMessageAsEmlFromSession,
  exportMessageAsJson,
  getAttachmentDownloadBuffer,
  getFolderSummary,
  getMessageSummary,
  isMailLikeSummary,
  sanitizeFileNameForDownload,
  sortMessageSummaries,
  type FolderMessagePage,
  type MessageDetail,
  type MessageSummary,
  type ViewerSessionIndex
} from './viewer'
import {
  getDefaultPstRootDirectory,
  listPstMailboxFiles,
  openPstMailbox,
  resolvePstMailboxPath
} from './pstCatalog'
import type { ReviewRecord } from './reviewTypes'
import { createZipStreamWriter } from './zipWriter'

export interface CreatePstReviewAppOptions {
  publicDir: string
  reviewStore: ReviewStore
  searchIndexStore: SearchIndexStore
  openApiSpec: Record<string, unknown>
  pstRootDir?: string
  apiSecurity?: ApiSecurityConfig
}

interface RequestInfoLike {
  origin?: string
  referer?: string
  ip?: string
  method?: string
  url?: string
  contentType?: string
  tenantId?: string
}

interface WebChecksLike {
  getRequestInfo?: (req: express.Request) => RequestInfoLike
}

interface M365AuthLike {
  CheckTokens?: (req: express.Request, res: express.Response, next: express.NextFunction) => unknown
}

export interface ApiSecurityConfig {
  bypassIps?: string[]
  allowedOrigins?: string[]
  webChecks?: WebChecksLike
  m365Auth?: M365AuthLike
}

interface SessionRecord {
  id: string
  index: ViewerSessionIndex
  filePath: string
  fileName: string
  scopePath: string
  scopeLabel: string
  mailboxKey: string
}

interface OpenMailboxRequestBody {
  scopePath?: string
  fileName?: string
}

interface HiddenRuleRequestBody {
  kind?: 'address' | 'subject'
  value?: string
  label?: string
}

interface ReviewPatchBody {
  flagged?: boolean
  tags?: unknown
}

interface ListFolderOptions {
  query: string
  mailOnly: boolean
  sort: string
  page: number
  pageSize: number
  reviewFlaggedOnly: boolean
  reviewTaggedOnly: boolean
  reviewTag: string
  mode: 'and' | 'or'
}

interface SessionResponse {
  sessionId: string
  scopePath: string
  scopeLabel: string
  fileName: string
  summary: ReturnType<typeof buildSessionSummary>
  tree: ReturnType<typeof buildFolderTree>
}

interface ReviewedMessageSummary extends MessageSummary {
  review: ReviewState
}

interface ReviewedMessageDetail extends MessageDetail {
  review: ReviewState
}

interface ReviewedFolderPage extends FolderMessagePage {
  items: ReviewedMessageSummary[]
  reviewFilters: {
    flaggedOnly: boolean
    taggedOnly: boolean
    tag: string
  }
}

interface SearchRequestQuery {
  scope?: string
  scopePath?: string
  sessionId?: string
  query?: string
  mode?: string
  mailOnly?: boolean
  sort?: string
  page?: number
  pageSize?: number
  reviewFlagged?: boolean
  reviewTagged?: boolean
  reviewTag?: string
}

interface FlaggedBundleQuery {
  scope?: string
  scopePath?: string
  sessionId?: string
}

interface BundleMailboxDescriptor {
  mailboxKey: string
  fileName: string
  scopePath: string
  scopeLabel: string
  session?: ViewerSessionIndex
}

interface FlaggedBundleManifestItem {
  sourcePstPath: string
  mailboxName: string
  mailboxKey: string
  scopePath: string
  scopeLabel: string
  fileName: string
  folderId: string
  folderPath: string
  messageId: string
  descriptorId: string
  kind: ReviewRecord['kind']
  subject: string
  review: ReviewState
  outputFile: string
  outputType: 'eml' | 'ics'
  status: 'exported' | 'error'
  error?: string
}

interface FlaggedBundleManifest {
  generatedAt: string
  scope: {
    scope: 'all' | 'search' | 'pst'
    scopePath: string
    scopeLabel: string
    sessionId: string
    sessionFileName: string
  }
  counts: {
    total: number
    exported: number
    failed: number
  }
  items: FlaggedBundleManifestItem[]
}

const DEFAULT_PAGE_SIZE = 50
const DEFAULT_DOC_TITLE = 'PST API Documentation'
const DEFAULT_SWAGGER_ASSET_PATH = path.dirname(
  require.resolve('swagger-ui-dist/swagger-ui.css')
)
const DEFAULT_AUTH_BYPASS_IPS = ['127.0.0.1', '::1', '::ffff:127.0.0.1', 'localhost']
const DEFAULT_CORS_ALLOW_HEADERS = [
  'Accept',
  'Authorization',
  'Content-Type',
  'GraphToken',
  'Origin',
  'Referer',
  'X-Graph-Token',
  'X-TenantId',
  'X-Requested-With',
  'graphtoken',
  'x-graph-token',
  'x-tenantid'
].join(', ')
const DEFAULT_CORS_ALLOW_METHODS = 'GET, POST, PATCH, DELETE, OPTIONS'

function createSessionId(): string {
  return randomBytes(12).toString('hex')
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeOrigin(value: unknown): string {
  const text = normalizeText(value)
  if (!text) {
    return ''
  }

  try {
    return new URL(text).origin.toLowerCase()
  } catch {
    return text.toLowerCase()
  }
}

function normalizeRequestInfoField(value: unknown): string {
  const text = normalizeText(value)
  if (!text) {
    return ''
  }

  if (/^\(.+\?\)$/.test(text) || /^[A-Za-z][A-Za-z0-9_-]*\?$/.test(text)) {
    return ''
  }

  return text
}

function normalizeIpAddress(value: unknown): string {
  const text = normalizeText(value)
  if (!text) {
    return ''
  }

  return text
    .replace(/^\[|\]$/g, '')
    .replace(/^::ffff:/i, '')
    .toLowerCase()
}

function parseList(value: string[] | undefined, fallback: string[] = []): string[] {
  const source = Array.isArray(value) ? value : []
  const combined = [...fallback, ...source]
  const values: string[] = []

  for (const item of combined) {
    for (const part of String(item || '').split(/[,\n;]/g)) {
      const normalized = part.trim()
      if (normalized) {
        values.push(normalized)
      }
    }
  }

  return values
}

function canonicalRequestOrigin(req: express.Request): string {
  const host = req.headers.host
  if (!host) {
    return ''
  }

  const protocol = req.protocol || 'http'
  return normalizeOrigin(`${protocol}://${host}`)
}

function getFallbackRequestInfo(req: express.Request): RequestInfoLike {
  return {
    origin: normalizeRequestInfoField(req.headers.origin),
    referer: normalizeRequestInfoField(req.headers.referer),
    ip:
      (typeof req.headers['x-forwarded-for'] === 'string'
        ? req.headers['x-forwarded-for'].split(',')[0]
        : '') || req.ip || req.socket?.remoteAddress || '',
    method: normalizeRequestInfoField(req.method),
    url: normalizeRequestInfoField(req.originalUrl || req.url),
    contentType: normalizeRequestInfoField(req.headers['content-type']),
    tenantId: normalizeRequestInfoField(req.headers['x-tenantid'])
  }
}

function getRequestInfo(req: express.Request, webChecks?: WebChecksLike): RequestInfoLike {
  try {
    if (webChecks?.getRequestInfo) {
      const info = webChecks.getRequestInfo(req)
      if (info && typeof info === 'object') {
        return {
          origin: normalizeRequestInfoField(info.origin),
          referer: normalizeRequestInfoField(info.referer),
          ip: normalizeRequestInfoField(info.ip),
          method: normalizeRequestInfoField(info.method),
          url: normalizeRequestInfoField(info.url),
          contentType: normalizeRequestInfoField(info.contentType),
          tenantId: normalizeRequestInfoField(info.tenantId)
        }
      }
    }
  } catch (error) {
    console.warn('Unable to read request info from webChecks:', error)
  }

  return getFallbackRequestInfo(req)
}

function isPublicApiPath(pathname: string): boolean {
  return pathname === API_ROUTES.openApiJson || pathname === API_ROUTES.docs || pathname.startsWith(`${API_ROUTES.docs}/`)
}

function isProtectedApiPath(pathname: string): boolean {
  return pathname.startsWith('/api/') && !isPublicApiPath(pathname)
}

function buildCorsHeaders(origin: string): Record<string, string> {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Headers': DEFAULT_CORS_ALLOW_HEADERS,
    'Access-Control-Allow-Methods': DEFAULT_CORS_ALLOW_METHODS,
    'Access-Control-Expose-Headers': 'Content-Disposition, Content-Length',
    Vary: 'Origin'
  }
}

function createApiSecurityMiddleware(
  config: ApiSecurityConfig = {}
): express.RequestHandler {
  const bypassIps = new Set(
    parseList(config.bypassIps, DEFAULT_AUTH_BYPASS_IPS).map(normalizeIpAddress).filter(Boolean)
  )
  const allowedOrigins = new Set(
    parseList(config.allowedOrigins).map(normalizeOrigin).filter(Boolean)
  )

  return async (req, res, next) => {
    const pathname = (req.originalUrl || req.url || '').split('?')[0] || ''
    if (!isProtectedApiPath(pathname)) {
      return next()
    }

    const info = getRequestInfo(req, config.webChecks)
    const requestOrigin = normalizeOrigin(info.origin || req.headers.origin || '')
    const requestHostOrigin = canonicalRequestOrigin(req)
    const requestIp = normalizeIpAddress(info.ip || req.ip || req.socket?.remoteAddress || '')
    const isBypassed = bypassIps.has(requestIp)
    const isSameOrigin = Boolean(requestOrigin && requestOrigin === requestHostOrigin)
    const originAllowed = !requestOrigin || isSameOrigin || allowedOrigins.has(requestOrigin)

    if (!originAllowed) {
      return res.status(403).json({
        error: 'CORS origin not allowed',
        origin: requestOrigin || info.origin || ''
      })
    }

    if (requestOrigin && !isSameOrigin && allowedOrigins.has(requestOrigin)) {
      res.set(buildCorsHeaders(requestOrigin))
    }

    if (req.method === 'OPTIONS') {
      return res.status(204).end()
    }

    if (isBypassed) {
      return next()
    }

    const auth = config.m365Auth?.CheckTokens
    if (typeof auth !== 'function') {
      return next(createAppError(500, 'Authentication middleware is not configured'))
    }

    let nextCalled = false
    const wrappedNext: express.NextFunction = (error?: unknown) => {
      nextCalled = true
      return next(error as never)
    }

    try {
      const result = auth(req, res, wrappedNext)
      if (result && typeof (result as Promise<unknown>).then === 'function') {
        await result
      }

      if (!nextCalled && !res.headersSent) {
        return next(createAppError(500, 'Authentication middleware did not complete'))
      }
    } catch (error) {
      if (!nextCalled && !res.headersSent) {
        return next(error as Error)
      }
    }
  }
}

function parsePositiveInt(
  value: string | string[] | undefined,
  fallback: number
): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return fallback
  }
  const parsed = Number.parseInt(raw, 10)
  if (Number.isNaN(parsed) || parsed < 1) {
    return fallback
  }
  return parsed
}

function parseBoolean(value: string | string[] | undefined, fallback = false): boolean {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined) {
    return fallback
  }
  return !['0', 'false', 'no', 'off'].includes(raw.trim().toLowerCase())
}

function parseSort(value: string | string[] | undefined): string {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return 'date-desc'
  }
  const normalized = raw.trim().toLowerCase()
  if (normalized === 'order' || normalized === 'folder-order') {
    return 'order'
  }
  return 'date-desc'
}

function parseSearchMode(value: string | string[] | undefined): 'and' | 'or' {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return 'and'
  }
  return raw.trim().toLowerCase() === 'or' ? 'or' : 'and'
}

function parseQueryText(value: string | string[] | undefined): string {
  const raw = Array.isArray(value) ? value[0] : value
  return normalizeText(raw)
}

function parseSearchModeFromQuery(
  query: string,
  fallbackMode: string | string[] | undefined
): 'and' | 'or' {
  const text = normalizeText(query)
  if (text) {
    const pattern = /"([^"]+)"|(\S+)/g
    let match: RegExpExecArray | null = null
    while ((match = pattern.exec(text))) {
      const token = String(match[1] || match[2] || '').trim().toLowerCase()
      if (!token) {
        continue
      }
      if (token === '|' || token.startsWith('|')) {
        return 'or'
      }
      if (token === '+' || token.startsWith('+')) {
        return 'and'
      }
    }
  }

  return parseSearchMode(fallbackMode)
}

function parseZeroBasedInt(value: string | string[] | undefined): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined) {
    return -1
  }
  const parsed = Number.parseInt(raw, 10)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : -1
}

function parseReviewFilters(value: Record<string, string | string[] | undefined>): ListFolderOptions {
  const queryValue =
    value.q !== undefined && value.q !== null && value.q !== ''
      ? value.q
      : value.query
  const query = parseQueryText(queryValue)
  return {
    query,
    mailOnly: parseBoolean(value.mailOnly, true),
    sort: parseSort(value.sort),
    page: parsePositiveInt(value.page, 1),
    pageSize: parsePositiveInt(value.pageSize, DEFAULT_PAGE_SIZE),
    reviewFlaggedOnly: parseBoolean(value.reviewFlagged, false),
    reviewTaggedOnly: parseBoolean(value.reviewTagged, false),
    reviewTag: normalizeText(value.reviewTag),
    mode: parseSearchModeFromQuery(query, value.mode)
  }
}

function normalizeScopePath(value: unknown): string {
  const text = String(value ?? '')
    .trim()
    .replace(/\\/g, '/')

  if (!text || text === '.') {
    return ''
  }

  const segments = text
    .split('/')
    .map((segment) => segment.trim())
    .filter(Boolean)

  if (!segments.length || segments.some((segment) => segment === '..')) {
    return ''
  }

  return segments.join('/')
}

function parseSearchScope(value: unknown): 'all' | 'search' | 'pst' {
  const normalized = normalizeText(value).toLowerCase()
  if (normalized === 'all' || normalized === 'search' || normalized === 'pst') {
    return normalized
  }
  return 'pst'
}

function getScopeLabel(scopePath: string): string {
  return scopePath ? scopePath.split('/').join(' / ') : 'PST root'
}

function resolveCatalogScopeSelection(
  rootPath: string,
  requestedScopePath: string
): ReturnType<typeof listPstMailboxFiles> {
  const normalizedScopePath = normalizeScopePath(requestedScopePath)
  const catalog = listPstMailboxFiles(rootPath, normalizedScopePath)

  if (normalizedScopePath && !catalog.scopes.some((scope) => scope.scopePath === normalizedScopePath)) {
    throw createAppError(404, 'Search scope not found')
  }

  if (normalizedScopePath && catalog.scopePath !== normalizedScopePath) {
    throw createAppError(404, 'Search scope not found')
  }

  return catalog
}

function parseFlaggedBundleScope(value: unknown): 'all' | 'search' | 'pst' {
  const normalized = normalizeText(value).toLowerCase()
  if (normalized === 'all' || normalized === 'search' || normalized === 'pst') {
    return normalized
  }
  return 'all'
}

function sanitizeBundleSegment(value: string, fallback: string): string {
  return sanitizeFileNameForDownload(value || fallback, fallback) || fallback
}

function buildBundleScopeSegment(scopePath: string): string {
  const normalized = normalizeScopePath(scopePath)
  if (!normalized) {
    return 'PST root'
  }
  return normalized.split('/').map((segment) => sanitizeBundleSegment(segment, 'scope')).join('/')
}

function buildBundleEntryPath(
  kind: ReviewRecord['kind'],
  scopePath: string,
  fileName: string,
  folderPath: string,
  subject: string,
  messageId: string
): string {
  const scopeSegment = buildBundleScopeSegment(scopePath)
  const mailboxSegment = sanitizeBundleSegment(fileName, 'mailbox')
  const folderSegments = normalizeScopePath(folderPath)
    ? normalizeScopePath(folderPath)
        .split('/')
        .map((segment) => sanitizeBundleSegment(segment, 'folder'))
        .filter(Boolean)
    : []
  const itemName = `${sanitizeBundleSegment(subject || 'message', 'message')}-${sanitizeBundleSegment(
    messageId,
    'message'
  )}.${kind === 'appointment' ? 'ics' : 'eml'}`
  return [kind === 'appointment' ? 'calendar' : 'mail', scopeSegment, mailboxSegment, ...folderSegments, itemName]
    .filter(Boolean)
    .join('/')
}

function buildBundleMailboxes(
  rootPath: string,
  scope: 'all' | 'search' | 'pst',
  scopePath: string,
  session: SessionRecord | null
): BundleMailboxDescriptor[] {
  if (scope === 'pst') {
    if (!session) {
      throw createAppError(400, 'Session id is required for selected PST exports')
    }
    return [
      {
        mailboxKey: session.mailboxKey,
        fileName: session.fileName,
        scopePath: session.scopePath,
        scopeLabel: session.scopeLabel,
        session: session.index
      }
    ]
  }

  const catalog =
    scope === 'search'
      ? resolveCatalogScopeSelection(rootPath, scopePath)
      : listPstMailboxFiles(rootPath)
  const scopes = scope === 'search' ? [catalog] : catalog.scopes

  const mailboxes: BundleMailboxDescriptor[] = []
  for (const catalogScope of scopes) {
    for (const file of catalogScope.files) {
      mailboxes.push({
        mailboxKey: resolvePstMailboxPath(rootPath, catalogScope.scopePath, file.fileName),
        fileName: file.fileName,
        scopePath: catalogScope.scopePath,
        scopeLabel: catalogScope.scopeLabel
      })
    }
  }
  return mailboxes
}

function createBundleManifest(scope: {
  scope: 'all' | 'search' | 'pst'
  scopePath: string
  scopeLabel: string
  sessionId: string
  sessionFileName: string
}): FlaggedBundleManifest {
  return {
    generatedAt: new Date().toISOString(),
    scope,
    counts: {
      total: 0,
      exported: 0,
      failed: 0
    },
    items: []
  }
}

function addBundleManifestItem(
  manifest: FlaggedBundleManifest,
  item: FlaggedBundleManifestItem
): void {
  manifest.items.push(item)
  manifest.counts.total += 1
  if (item.status === 'exported') {
    manifest.counts.exported += 1
  } else {
    manifest.counts.failed += 1
  }
}

function isReviewMatch(
  review: ReviewState | null,
  options: {
    flaggedOnly: boolean
    taggedOnly: boolean
    tag: string
  }
): boolean {
  if (options.flaggedOnly && !review?.flagged) {
    return false
  }
  if (options.taggedOnly && (!review || review.tags.length === 0)) {
    return false
  }
  if (options.tag) {
    const needle = options.tag.toLowerCase()
    if (!review || !review.tags.some((value) => value.toLowerCase() === needle)) {
      return false
    }
  }
  return true
}

function createAttachmentBaseUrl(sessionId: string, messageId: string): string {
  return `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(
    messageId
  )}/attachments/`
}

function normalizeReviewState(review: ReviewState | null): ReviewState {
  return (
    review || {
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    }
  )
}

function buildReviewedSummary(
  summary: MessageSummary,
  review: ReviewState | null
): ReviewedMessageSummary {
  return {
    ...summary,
    review: normalizeReviewState(review)
  }
}

function buildReviewedDetail(
  detail: MessageDetail,
  review: ReviewState | null
): ReviewedMessageDetail {
  return {
    ...detail,
    review: normalizeReviewState(review)
  }
}

function buildDocsHtml(): string {
  return `<!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>${DEFAULT_DOC_TITLE}</title>
        <link rel="stylesheet" href="/api/docs/swagger-ui.css" />
        <style>
          html { box-sizing: border-box; }
          *, *::before, *::after { box-sizing: inherit; }
          body { margin: 0; background: #f5f7fb; }
          .swagger-ui .topbar { display: none; }
        </style>
      </head>
      <body>
        <div id="swagger-ui"></div>
        <script src="/api/docs/swagger-ui-bundle.js"></script>
        <script src="/api/docs/swagger-ui-standalone-preset.js"></script>
        <script>
          window.onload = function () {
            window.ui = SwaggerUIBundle({
              url: '/api/openapi.json',
              dom_id: '#swagger-ui',
              deepLinking: true,
              displayRequestDuration: true,
              presets: [SwaggerUIBundle.presets.apis, SwaggerUIBundle.SwaggerUIStandalonePreset],
              layout: 'BaseLayout'
            })
          }
        </script>
      </body>
    </html>`
}

function createAppError(statusCode: number, message: string): Error {
  const error = new Error(message)
  ;(error as Error & { statusCode?: number }).statusCode = statusCode
  return error
}

function toReviewableContext(
  session: SessionRecord,
  summary: MessageSummary
) {
  return buildReviewContext(session.filePath, session.fileName, summary)
}

function getSessionOrThrow(
  sessions: Map<string, SessionRecord>,
  sessionId: string
): SessionRecord {
  const session = sessions.get(sessionId)
  if (!session) {
    throw createAppError(404, 'Session not found')
  }
  return session
}

async function buildMessageDetailResponse(
  session: SessionRecord,
  messageId: string,
  reviewStore: ReviewStore
): Promise<ReviewedMessageDetail> {
  const summary = session.index.messages.get(messageId)
  if (!summary) {
    throw createAppError(404, `Unknown message: ${messageId}`)
  }

  const review = await reviewStore.getReview(session.filePath, messageId)
  if (summary.parseError) {
    return buildReviewedDetail(buildEmptyMessageDetail(summary), review)
  }

  try {
    const detail = await Promise.resolve(
      buildMessageDetailFromSession(session.index, messageId, 1)
    )
    return buildReviewedDetail(
      {
        ...detail,
        attachments: detail.attachments.map((attachment) => ({
          ...attachment,
          downloadUrl: attachment.downloadUrl || createAttachmentBaseUrl(session.id, messageId)
        }))
      },
      review
    )
  } catch (error) {
    const parseError = error instanceof Error ? error.message : String(error)
    return buildReviewedDetail(
      buildEmptyMessageDetail({
        ...summary,
        parseError
      }),
      review
    )
  }
}

async function buildReviewedFolderPage(
  session: SessionRecord,
  folderId: string,
  options: ListFolderOptions,
  reviewStore: ReviewStore,
  hiddenRules: HiddenRuleRecord[]
): Promise<ReviewedFolderPage> {
  const { folder, items: collectedItems } = collectFolderMessages(
    session.index,
    folderId,
    {
      query: options.query,
      mailOnly: options.mailOnly,
      mode: options.mode
    },
    hiddenRules
  )
  const reviewMap = await reviewStore.getMany(
    session.filePath,
    collectedItems.map((message) => message.id)
  )

  let items = [...collectedItems]

  if (options.reviewFlaggedOnly || options.reviewTaggedOnly || options.reviewTag) {
    items = items.filter((message) => {
      const review = reviewMap.get(message.id) || null
      if (options.reviewFlaggedOnly && !review?.flagged) {
        return false
      }
      if (options.reviewTaggedOnly && (!review || review.tags.length === 0)) {
        return false
      }
      if (options.reviewTag) {
        const tag = options.reviewTag.toLowerCase()
        if (!review || !review.tags.some((value) => value.toLowerCase() === tag)) {
          return false
        }
      }
      return true
    })
  }

  const sorted = sortMessageSummaries(items, options.sort)
  const total = sorted.length
  const totalPages = Math.max(1, Math.ceil(total / options.pageSize))
  const page = Math.min(Math.max(options.page, 1), totalPages)
  const start = (page - 1) * options.pageSize
  const pageItems = sorted.slice(start, start + options.pageSize)
  const pageReviewMap = await reviewStore.getMany(
    session.filePath,
    pageItems.map((item) => item.id)
  )

  return {
    folder,
    items: pageItems.map((item) => buildReviewedSummary(item, pageReviewMap.get(item.id) || null)),
    total,
    page,
    pageSize: options.pageSize,
    totalPages,
    query: options.query,
    mailOnly: options.mailOnly,
    sort: options.sort,
    reviewFilters: {
      flaggedOnly: options.reviewFlaggedOnly,
      taggedOnly: options.reviewTaggedOnly,
      tag: options.reviewTag
    }
  }
}

async function buildFolderPageWithReviews(
  session: SessionRecord,
  folderId: string,
  options: ListFolderOptions,
  reviewStore: ReviewStore,
  hiddenRules: HiddenRuleRecord[]
): Promise<ReviewedFolderPage> {
  return buildReviewedFolderPage(session, folderId, options, reviewStore, hiddenRules)
}

function assertReviewableMessage(summary: MessageSummary): void {
  if (!isReviewableSummary(summary)) {
    throw createAppError(400, 'Review state is available for mail and appointment items only')
  }
}

function isReviewableSummary(summary: MessageSummary): boolean {
  return isMailLikeSummary(summary) || summary.kind === 'appointment'
}

function responseJson(
  res: express.Response,
  statusCode: number,
  payload: unknown
): void {
  res.status(statusCode).json(payload)
}

function responseBinary(
  res: express.Response,
  statusCode: number,
  contentType: string,
  fileName: string,
  data: Buffer
): void {
  res.status(statusCode)
    .type(contentType)
    .set('Content-Disposition', `attachment; filename="${fileName}"`)
    .set('Content-Length', String(data.length))
    .send(data)
}

function toOpenApiPath(pathTemplate: string): string {
  return pathTemplate.replace(/:([A-Za-z0-9_]+)/g, '{$1}')
}

function createRouteErrorHandler(
  res: express.Response,
  error: unknown
): void {
  const statusCode =
    typeof error === 'object' && error && 'statusCode' in error
      ? Number((error as { statusCode?: number }).statusCode || 500)
      : isCatalogValidationError(error)
        ? 400
        : 500
  const message = error instanceof Error ? error.message : String(error)
  responseJson(res, statusCode, { error: message })
}

function isCatalogValidationError(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error)
  return [
    'Scope path must stay within the PST folder',
    'Mailbox file name is required',
    'Mailbox file name must not include a path',
    'Only .pst and .ost files are supported'
  ].some((needle) => message.includes(needle))
}

function safeDownloadName(name: string, fallback: string): string {
  return sanitizeFileNameForDownload(name || fallback, fallback)
}

export function createPstReviewApp(options: CreatePstReviewAppOptions): express.Express {
  const app = express()
  const sessions = new Map<string, SessionRecord>()
  const publicDir = options.publicDir
  const pstRootDir = options.pstRootDir || getDefaultPstRootDirectory()
  const reviewStore = options.reviewStore
  const searchIndexStore = options.searchIndexStore

  async function syncSearchIndexForMailbox(session: SessionRecord): Promise<void> {
    const messageIds = [...session.index.messages.keys()]
    const reviewMap = await reviewStore.getMany(session.filePath, messageIds)
    const documents = buildSearchIndexDocumentsFromSession(
      session.index,
      {
        mailboxKey: session.filePath,
        scopePath: session.scopePath,
        scopeLabel: session.scopeLabel,
        fileName: session.fileName,
        mailboxName: session.index.mailboxName
      },
      reviewMap
    )
    await searchIndexStore.replaceMailboxDocuments(session.filePath, documents)
  }

  app.disable('x-powered-by')
  app.use(express.json({ limit: '2mb' }))
  app.use(express.static(publicDir))
  app.use(createApiSecurityMiddleware(options.apiSecurity))

  app.get(API_ROUTES.openApiJson, (_req, res) => {
    responseJson(res, 200, options.openApiSpec)
  })

  app.use(
    `${API_ROUTES.docs}/`,
    express.static(DEFAULT_SWAGGER_ASSET_PATH, {
      index: false
    })
  )
  app.get(API_ROUTES.docs, (_req, res) => {
    res.status(200).type('html').send(buildDocsHtml())
  })

  app.get(API_ROUTES.pstCatalog, (req, res) => {
    try {
      const scopePath = typeof req.query.scopePath === 'string' ? normalizeText(req.query.scopePath) : ''
      responseJson(res, 200, listPstMailboxFiles(pstRootDir, scopePath))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.pstOpen, async (req, res) => {
    try {
      const body = (req.body || {}) as OpenMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      const index = openPstMailbox(pstRootDir, scopePath, fileName)
      const sessionId = createSessionId()
      const scopeLabel = scopePath ? scopePath.split('/').join(' / ') : 'PST root'
      const record: SessionRecord = {
        id: sessionId,
        index,
        filePath: index.filePath,
        fileName: index.fileName,
        scopePath,
        scopeLabel,
        mailboxKey: index.filePath
      }
      sessions.set(sessionId, record)

      try {
        await syncSearchIndexForMailbox(record)
      } catch (error) {
        console.warn(
          `Unable to refresh search index for ${index.fileName}:`,
          error instanceof Error ? error.message : error
        )
      }

      const summary = buildSessionSummary(index)
      responseJson(res, 200, {
        sessionId,
        scopePath,
        scopeLabel,
        fileName,
        summary,
        tree: buildFolderTree(index)
      } satisfies SessionResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.search, async (req, res) => {
    try {
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const scope = parseSearchScope(req.query.scope) as SearchScope
      const requestedScopePath = normalizeScopePath(req.query.scopePath)
      const sessionId = normalizeText(req.query.sessionId)
      let scopePath = ''
      let scopeLabel = 'All cases/searches'
      let mailboxKey = ''

      if (scope === 'pst') {
        const session = getSessionOrThrow(sessions, sessionId)
        scopePath = session.scopePath
        scopeLabel = session.scopeLabel || getScopeLabel(scopePath)
        mailboxKey = session.filePath
      } else if (scope === 'search') {
        const catalog = resolveCatalogScopeSelection(pstRootDir, requestedScopePath)
        scopePath = catalog.scopePath
        scopeLabel = catalog.scopeLabel
      }

      const page = await searchIndexStore.search({
        scope,
        scopePath,
        mailboxKey,
        query: filters.query,
        mode: filters.mode,
        mailOnly: filters.mailOnly,
        sort: filters.sort,
        page: filters.page,
        pageSize: filters.pageSize,
        reviewFlaggedOnly: filters.reviewFlaggedOnly,
        reviewTaggedOnly: filters.reviewTaggedOnly,
        reviewTag: filters.reviewTag
      })
      responseJson(res, 200, {
        scope,
        scopePath: page.scopePath || scopePath,
        scopeLabel: page.scopeLabel || scopeLabel,
        page
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.searchIndexRefresh, async (_req, res) => {
    try {
      const summary = await refreshSearchIndexFromCatalog(pstRootDir, reviewStore, searchIndexStore)
      responseJson(res, 200, {
        summary,
        hiddenRules: await searchIndexStore.listHiddenRules()
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.searchFilters, async (_req, res) => {
    try {
      responseJson(res, 200, {
        items: await searchIndexStore.listHiddenRules()
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.searchFilters, async (req, res) => {
    try {
      const body = (req.body || {}) as HiddenRuleRequestBody
      const kind = body.kind
      const value = normalizeText(body.value)
      if (kind !== 'address' && kind !== 'subject') {
        throw createAppError(400, 'Filter kind must be address or subject')
      }
      if (!value) {
        throw createAppError(400, 'Filter value is required')
      }
      const rule = await searchIndexStore.upsertHiddenRule({
        kind,
        value,
        label: normalizeText(body.label || value)
      })
      responseJson(res, 200, { rule })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.searchFilter, async (req, res) => {
    try {
      const filterId = normalizeText(req.params.filterId)
      if (!filterId) {
        throw createAppError(400, 'Filter id is required')
      }
      const deleted = await searchIndexStore.deleteHiddenRule(filterId)
      responseJson(res, 200, { deleted })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.sessionSummary, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      responseJson(res, 200, {
        sessionId: session.id,
        summary: buildSessionSummary(session.index)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.sessionTree, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      responseJson(res, 200, {
        sessionId: session.id,
        tree: buildFolderTree(session.index)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.folderMessages, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const folderId = normalizeText(req.params.folderId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const hiddenRules = await searchIndexStore.listHiddenRules()
      const page = await buildFolderPageWithReviews(
        session,
        folderId,
        filters,
        reviewStore,
        hiddenRules
      )
      responseJson(res, 200, {
        sessionId: session.id,
        page
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageDetail, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      responseJson(res, 200, {
        sessionId: session.id,
        detail
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExtract, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      const fields = normalizeExtractionFields(req.query.fields)
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        fields,
        record: buildMessageExtractionRecord(detail, detail.review, fields)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.folderExtract, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const folderId = normalizeText(req.params.folderId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const fields = normalizeExtractionFields(req.query.fields)
      const hiddenRules = await searchIndexStore.listHiddenRules()
      const page = await buildFolderPageWithReviews(
        session,
        folderId,
        filters,
        reviewStore,
        hiddenRules
      )
      const folder = getFolderSummary(session.index, folderId)
      const extraction = buildFolderExtractionPage(
        page,
        new Map(page.items.map((item) => [item.id, item.review])),
        fields,
        (summary, review, fieldList) => {
          if (isSummaryOnlyExtraction(fieldList)) {
            return buildSummaryExtractionRecord(summary, review, fieldList)
          }
          const detail = buildMessageDetailFromSession(session.index, summary.id, 1)
          return buildMessageExtractionRecord(detail, review, fieldList)
        }
      )

      responseJson(res, 200, {
        sessionId: session.id,
        folder: {
          id: folder.id,
          descriptorId: folder.descriptorId,
          displayName: folder.displayName,
          path: folder.path
        },
        fields: extraction.fields,
        paging: extraction.paging,
        items: extraction.items
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExportJson, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      const fileName = `${safeDownloadName(detail.subject || 'message', 'message')}.json`
      responseBinary(
        res,
        200,
        'application/json; charset=utf-8',
        fileName,
        Buffer.from(exportMessageAsJson(detail), 'utf8')
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExportEml, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const summary = getMessageSummary(session.index, messageId)
      const fileName = `${safeDownloadName(summary.subject || 'message', 'message')}.eml`
      const eml = summary.parseError
        ? exportMessageAsEml(buildReviewedDetail(buildEmptyMessageDetail(summary), null))
        : exportMessageAsEmlFromSession(session.index, messageId)
      responseBinary(
        res,
        200,
        'message/rfc822; charset=utf-8',
        fileName,
        Buffer.from(eml, 'utf8')
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.flaggedBundleExport, async (req, res) => {
    try {
      const query = req.query as FlaggedBundleQuery
      const scope = parseFlaggedBundleScope(query.scope)
      const scopePath = normalizeScopePath(query.scopePath)
      const sessionId = normalizeText(query.sessionId)
      const session = scope === 'pst' ? getSessionOrThrow(sessions, sessionId) : null
      const bundleScope =
        scope === 'search'
          ? resolveCatalogScopeSelection(pstRootDir, scopePath)
          : scope === 'all'
            ? listPstMailboxFiles(pstRootDir)
            : null
      const scopeLabel =
        scope === 'all'
          ? 'All cases/searches'
          : scope === 'search'
            ? bundleScope?.scopeLabel || getScopeLabel(scopePath)
            : session?.scopeLabel || 'Selected PST'
      const writer = createZipStreamWriter(res)
      const manifest = createBundleManifest({
        scope,
        scopePath: scope === 'all' ? '' : scope === 'search' ? normalizeText(bundleScope?.scopePath || scopePath) : session?.scopePath || '',
        scopeLabel,
        sessionId: session?.id || '',
        sessionFileName: session?.fileName || ''
      })
      const mailboxes = buildBundleMailboxes(
        pstRootDir,
        scope,
        scope === 'search' ? normalizeText(bundleScope?.scopePath || scopePath) : scopePath,
        session
      )

      res.status(200)
        .type('application/zip')
        .set('Content-Disposition', 'attachment; filename="flagged-bundle.zip"')
      res.flushHeaders?.()

      for (const mailbox of mailboxes) {
        let mailboxSession = mailbox.session || null
        let flaggedReviews = await reviewStore.listReviews(mailbox.mailboxKey, {
          flaggedOnly: true
        })
        if (!flaggedReviews.length) {
          continue
        }

        if (!mailboxSession) {
          try {
            mailboxSession = openPstMailbox(pstRootDir, mailbox.scopePath, mailbox.fileName)
          } catch (error) {
            const message = error instanceof Error ? error.message : String(error)
            for (const review of flaggedReviews) {
              const outputType = review.kind === 'appointment' ? 'ics' : 'eml'
              const outputFile = buildBundleEntryPath(
                review.kind,
                mailbox.scopePath,
                mailbox.fileName,
                review.folderPath,
                review.subject,
                review.messageId
              )
              addBundleManifestItem(manifest, {
                sourcePstPath: mailbox.mailboxKey,
                mailboxName: mailbox.scopeLabel || mailbox.fileName,
                mailboxKey: mailbox.mailboxKey,
                scopePath: mailbox.scopePath,
                scopeLabel: mailbox.scopeLabel,
                fileName: mailbox.fileName,
                folderId: review.folderId,
                folderPath: review.folderPath,
                messageId: review.messageId,
                descriptorId: review.descriptorId,
                kind: review.kind,
                subject: review.subject,
                review: {
                  flagged: review.flagged,
                  tags: [...review.tags],
                  createdAt: review.createdAt,
                  updatedAt: review.updatedAt
                },
                outputFile,
                outputType,
                status: 'error',
                error: message
              })
            }
            continue
          }
        }

        const activeMailboxSession = mailboxSession
        if (!activeMailboxSession) {
          continue
        }

        for (const review of flaggedReviews) {
          const summary = activeMailboxSession.messages.get(review.messageId) || null
          const outputType = review.kind === 'appointment' ? 'ics' : 'eml'
          const outputFile = buildBundleEntryPath(
            review.kind,
            mailbox.scopePath,
            mailbox.fileName,
            review.folderPath,
            review.subject,
            review.messageId
          )

          if (!summary) {
            addBundleManifestItem(manifest, {
              sourcePstPath: mailbox.mailboxKey,
              mailboxName: activeMailboxSession.mailboxName || mailbox.fileName,
              mailboxKey: mailbox.mailboxKey,
              scopePath: mailbox.scopePath,
              scopeLabel: mailbox.scopeLabel,
              fileName: mailbox.fileName,
              folderId: review.folderId,
              folderPath: review.folderPath,
              messageId: review.messageId,
              descriptorId: review.descriptorId,
              kind: review.kind,
              subject: review.subject,
              review: {
                flagged: review.flagged,
                tags: [...review.tags],
                createdAt: review.createdAt,
                updatedAt: review.updatedAt
              },
              outputFile,
              outputType,
              status: 'error',
              error: 'Message not found in mailbox session'
            })
            continue
          }

          try {
            const payload =
              outputType === 'ics'
                ? exportAppointmentAsIcsFromSession(activeMailboxSession, review.messageId)
                : exportMessageAsEmlFromSession(activeMailboxSession, review.messageId)
            const mtime = summary.modificationTime
              ? new Date(summary.modificationTime)
              : summary.creationTime
                ? new Date(summary.creationTime)
                : new Date()
            await writer.addText(outputFile, payload, { mtime })
            addBundleManifestItem(manifest, {
              sourcePstPath: mailbox.mailboxKey,
              mailboxName: activeMailboxSession.mailboxName || mailbox.fileName,
              mailboxKey: mailbox.mailboxKey,
              scopePath: mailbox.scopePath,
              scopeLabel: mailbox.scopeLabel,
              fileName: mailbox.fileName,
              folderId: review.folderId,
              folderPath: review.folderPath,
              messageId: review.messageId,
              descriptorId: review.descriptorId,
              kind: review.kind,
              subject: review.subject,
              review: {
                flagged: review.flagged,
                tags: [...review.tags],
                createdAt: review.createdAt,
                updatedAt: review.updatedAt
              },
              outputFile,
              outputType,
              status: 'exported'
            })
          } catch (error) {
            addBundleManifestItem(manifest, {
              sourcePstPath: mailbox.mailboxKey,
              mailboxName: activeMailboxSession.mailboxName || mailbox.fileName,
              mailboxKey: mailbox.mailboxKey,
              scopePath: mailbox.scopePath,
              scopeLabel: mailbox.scopeLabel,
              fileName: mailbox.fileName,
              folderId: review.folderId,
              folderPath: review.folderPath,
              messageId: review.messageId,
              descriptorId: review.descriptorId,
              kind: review.kind,
              subject: review.subject,
              review: {
                flagged: review.flagged,
                tags: [...review.tags],
                createdAt: review.createdAt,
                updatedAt: review.updatedAt
              },
              outputFile,
              outputType,
              status: 'error',
              error: error instanceof Error ? error.message : String(error)
            })
          }
        }
      }

      await writer.addText('manifest.json', JSON.stringify(manifest, null, 2), {
        mtime: new Date(manifest.generatedAt)
      })
      await writer.finalize()
      res.end()
    } catch (error) {
      if (!res.headersSent) {
        createRouteErrorHandler(res, error)
      } else {
        res.destroy(error instanceof Error ? error : undefined)
      }
    }
  })

  app.get(API_ROUTES.messageAttachment, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const attachmentIndex = parseZeroBasedInt(req.params.attachmentIndex)
      if (attachmentIndex < 0) {
        throw createAppError(404, 'Attachment not found')
      }
      const payload = getAttachmentDownloadBuffer(session.index, messageId, attachmentIndex)
      const fileName = safeDownloadName(payload.filename, 'attachment')
      responseBinary(res, 200, payload.contentType, fileName, payload.buffer)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const review = normalizeReviewState(await reviewStore.getReview(session.filePath, messageId))
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.patch(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const body = (req.body || {}) as ReviewPatchBody
      if (body.flagged === undefined && body.tags === undefined) {
        throw createAppError(400, 'Provide flagged or tags to update review state')
      }

      const summary = getMessageSummary(session.index, messageId)
      assertReviewableMessage(summary)
      const review = await reviewStore.upsertReview({
        ...toReviewableContext(session, summary),
        flagged: body.flagged,
        tags: body.tags
      })
      await searchIndexStore.updateReviewState(session.filePath, messageId, review)

      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review: normalizeReviewState(review)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      await reviewStore.deleteReview(session.filePath, messageId)
      await searchIndexStore.updateReviewState(session.filePath, messageId, null)
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review: normalizeReviewState(null)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.mailboxReviewQueue, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const pageSize = filters.pageSize
      const page = filters.page
      const records = await reviewStore.listReviews(session.filePath, {
        query: filters.query,
        flaggedOnly: filters.reviewFlaggedOnly,
        taggedOnly: filters.reviewTaggedOnly,
        tag: filters.reviewTag
      })
      const total = records.length
      const totalPages = Math.max(1, Math.ceil(total / pageSize))
      const currentPage = Math.min(Math.max(page, 1), totalPages)
      const start = (currentPage - 1) * pageSize
      const items = records.slice(start, start + pageSize)

      responseJson(res, 200, {
        sessionId: session.id,
        mailboxKey: session.filePath,
        total,
        page: currentPage,
        pageSize,
        totalPages,
        filters: {
          query: filters.query,
          flaggedOnly: filters.reviewFlaggedOnly,
          taggedOnly: filters.reviewTaggedOnly,
          tag: filters.reviewTag
        },
        items
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.use(
    (err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
      createRouteErrorHandler(res, err)
    }
  )

  return app
}
