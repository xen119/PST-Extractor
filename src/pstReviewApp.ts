import express from 'express'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import nodemailer from 'nodemailer'
import QRCode from 'qrcode'
import { API_ROUTES } from './apiRoutes'
import {
  type AuditActor,
  type AuditLogEntry,
  type AuditLogStore,
  createFileAuditLogStore
} from './auditLog'
import {
  buildSmtpSettingsView,
  type AppSettingsStore,
  type SmtpSettingsInput,
  type SmtpSettingsRecord,
  createMemoryAppSettingsStore,
  mergeSmtpSettings,
  normalizeSmtpSettingsInput
} from './appSettings'
import {
  type AuthInviteResult,
  type AuthMfaCompletionResult,
  type AuthMfaSetupResult,
  type AuthUserListItem,
  type AuthUserStore,
  createMemoryAuthUserStore
} from './authUsers'
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
  refreshSearchIndexFromCatalog,
  type HiddenRuleRecord,
  type SearchIndexDocument,
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
  listRemovedPstMailboxFiles,
  movePstMailboxToRemoved,
  openPstMailbox,
  resolvePstMailboxPath,
  restorePstMailboxFromRemoved
} from './pstCatalog'
import type { ReviewRecord } from './reviewTypes'
import { createZipStreamWriter } from './zipWriter'

export interface CreatePstReviewAppOptions {
  publicDir: string
  reviewStore: ReviewStore
  searchIndexStore: SearchIndexStore
  openApiSpec: Record<string, unknown>
  pstRootDir?: string
  auth?: AppAuthConfig
  authUserStore?: AuthUserStore
  appSettingsStore?: AppSettingsStore
  auditLogDir?: string
  apiSecurity?: ApiSecurityConfig
  smtpTransportFactory?: SmtpTransportFactory
}

export interface AppAuthConfig {
  username: string
  password: string
  sessionTtlMinutes?: number
  cookieName?: string
  inviteTtlMinutes?: number
  mfaIssuer?: string
  publicBaseUrl?: string
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
  hasLocalAuthSession?: (req: express.Request) => boolean
}

interface AuthSessionRecord {
  token: string
  username: string
  expiresAt: number
  mfaEnabled: boolean
}

interface AuthMfaChallengeRecord {
  token: string
  username: string
  expiresAt: number
}

interface AuthStatusResponse {
  authenticated: boolean
  enabled: boolean
  canManageUsers: boolean
  mfaEnabled: boolean
  mfaEnforced: boolean
  mfaRequired: boolean
  mfaChallengeExpiresAt: string | null
  user: {
    username: string
  } | null
  expiresAt: string | null
}

interface AuthUsersResponse {
  users: AuthUserListItem[]
}

interface AuthUserCreateResponse {
  user: AuthUserListItem
  inviteUrl?: string
  emailSent?: boolean
  inviteExpiresAt?: string
}

interface AuthUserDeleteResponse {
  user: AuthUserListItem
}

interface AuthMfaEnforceRequestBody {
  enforced?: boolean
}

interface AuthInviteLookupResponse {
  invite: AuthUserListItem
}

interface AuthInviteAcceptResponse {
  user: AuthUserListItem
  mfaAvailable: boolean
}

interface AuthMfaStartResponse {
  user: AuthUserListItem
  secret: string
  otpauthUri: string
  qrCodeDataUrl: string
}

interface AuthMfaCompleteResponse {
  user: AuthUserListItem
  recoveryCodes: string[]
}

interface ActivityLogResponse {
  entries: AuditLogEntry[]
}

interface SmtpSettingsResponse {
  settings: ReturnType<typeof buildSmtpSettingsView>
}

interface SmtpTestRequestBody {
  recipient?: string
  enabled?: boolean
  host?: string
  port?: number | string
  secure?: boolean
  username?: string
  password?: string
  fromName?: string
  fromAddress?: string
  replyTo?: string
}

interface SmtpTestResponse {
  success: boolean
  recipient: string
  messageId: string
  accepted: string[]
  rejected: string[]
}

interface SmtpTransportSendResult {
  messageId?: string
  accepted?: string[]
  rejected?: string[]
}

interface SmtpTransportLike {
  sendMail(message: Record<string, unknown>): Promise<SmtpTransportSendResult>
  close?: () => void | Promise<void>
}

type SmtpTransportFactory = (settings: SmtpSettingsRecord) => SmtpTransportLike

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

interface CatalogScopeRequestQuery {
  scopePath?: string
}

interface MoveMailboxRequestBody {
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
const DEFAULT_AUTH_SESSION_COOKIE = 'pst-review-session'
const DEFAULT_AUTH_MFA_CHALLENGE_COOKIE_SUFFIX = '-mfa-challenge'
const DEFAULT_AUTH_SESSION_TTL_MINUTES = 180
const DEFAULT_AUTH_INVITE_TTL_MINUTES = 24 * 60
const DEFAULT_AUTH_MFA_ISSUER = 'PST Mail Explorer'
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

function normalizeAuthUsername(value: unknown): string {
  return normalizeText(value)
}

function normalizeAuthUsernameKey(value: unknown): string {
  return normalizeAuthUsername(value).toLowerCase()
}

function isAdminAuthSession(
  session: AuthSessionRecord | null,
  auth: { enabled: boolean; username: string }
): boolean {
  if (!auth.enabled || !session) {
    return false
  }

  return normalizeAuthUsernameKey(session.username) === normalizeAuthUsernameKey(auth.username)
}

function normalizeAuthConfig(auth?: AppAuthConfig): {
  enabled: boolean
  username: string
  sessionTtlMinutes: number
  cookieName: string
  inviteTtlMinutes: number
  mfaIssuer: string
  publicBaseUrl: string
  seedUsers: Array<{ username: string; password: string }>
} {
  if (!auth) {
    return {
      enabled: false,
      username: '',
      sessionTtlMinutes: DEFAULT_AUTH_SESSION_TTL_MINUTES,
      cookieName: DEFAULT_AUTH_SESSION_COOKIE,
      inviteTtlMinutes: DEFAULT_AUTH_INVITE_TTL_MINUTES,
      mfaIssuer: DEFAULT_AUTH_MFA_ISSUER,
      publicBaseUrl: '',
      seedUsers: []
    }
  }

  const username = normalizeAuthUsername(auth.username)
  const password = String(auth.password ?? '')
  if (!username || !password.trim()) {
    throw createAppError(500, 'Authentication requires both a username and password')
  }

  const sessionTtlMinutes =
    Number.isFinite(auth.sessionTtlMinutes) && Number(auth.sessionTtlMinutes) > 0
      ? Number(auth.sessionTtlMinutes)
      : DEFAULT_AUTH_SESSION_TTL_MINUTES

  return {
    enabled: true,
    username,
    sessionTtlMinutes,
    cookieName: normalizeText(auth.cookieName) || DEFAULT_AUTH_SESSION_COOKIE,
    inviteTtlMinutes:
      Number.isFinite(auth.inviteTtlMinutes) && Number(auth.inviteTtlMinutes) > 0
        ? Number(auth.inviteTtlMinutes)
        : DEFAULT_AUTH_INVITE_TTL_MINUTES,
    mfaIssuer: normalizeText(auth.mfaIssuer) || DEFAULT_AUTH_MFA_ISSUER,
    publicBaseUrl: normalizeOrigin(auth.publicBaseUrl),
    seedUsers: [{ username, password }]
  }
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

function parseCookieHeader(value: string | undefined): Map<string, string> {
  const cookies = new Map<string, string>()
  for (const part of String(value || '').split(';')) {
    const index = part.indexOf('=')
    if (index < 0) {
      continue
    }
    const name = part.slice(0, index).trim()
    if (!name) {
      continue
    }
    const rawValue = part.slice(index + 1).trim()
    try {
      cookies.set(name, decodeURIComponent(rawValue))
    } catch {
      cookies.set(name, rawValue)
    }
  }
  return cookies
}

function getCookieValue(req: express.Request, cookieName: string): string {
  if (!cookieName) {
    return ''
  }
  return parseCookieHeader(req.headers.cookie).get(cookieName) || ''
}

function isPublicApiPath(pathname: string): boolean {
  return (
    pathname === API_ROUTES.openApiJson ||
    pathname === API_ROUTES.docs ||
    pathname.startsWith(`${API_ROUTES.docs}/`) ||
    pathname === API_ROUTES.authLogin ||
    pathname === API_ROUTES.authMe ||
    pathname === API_ROUTES.authLogout ||
    pathname.startsWith(`${API_ROUTES.authInviteLookup.split('/:token')[0]}`) ||
    pathname.startsWith(`${API_ROUTES.authInviteAccept.split('/:token')[0]}`) ||
    pathname === API_ROUTES.authMfaChallenge
  )
}

function isAuthApiPath(pathname: string): boolean {
  return (
    pathname === API_ROUTES.authLogin ||
    pathname === API_ROUTES.authMe ||
    pathname === API_ROUTES.authLogout ||
    pathname === API_ROUTES.authMfaChallenge
  )
}

function isProtectedApiPath(pathname: string): boolean {
  return pathname.startsWith('/api/') && !isPublicApiPath(pathname) && !isAuthApiPath(pathname)
}

function buildCorsHeaders(origin: string): Record<string, string> {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Credentials': 'true',
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
    const isAuthRoute = isAuthApiPath(pathname)
    const isProtectedRoute = isProtectedApiPath(pathname)
    if (!isProtectedRoute && !isAuthRoute) {
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

    if (isAuthRoute) {
      return next()
    }

    if (isBypassed) {
      return next()
    }

    if (config.hasLocalAuthSession?.(req)) {
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
  requestedScopePath: string,
  loader: (rootPath: string, scopePath?: string) => ReturnType<typeof listPstMailboxFiles> = listPstMailboxFiles
): ReturnType<typeof listPstMailboxFiles> {
  const normalizedScopePath = normalizeScopePath(requestedScopePath)
  const catalog = loader(rootPath, normalizedScopePath)

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

function getReviewOwnerUsername(session: AuthSessionRecord | null | undefined): string {
  return normalizeText(session?.username || '') || 'anonymous'
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
  summary: MessageSummary,
  reviewerUsername: string
) {
  return {
    ...buildReviewContext(session.filePath, session.fileName, summary),
    reviewerUsername: normalizeText(reviewerUsername) || 'anonymous'
  }
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
  reviewStore: ReviewStore,
  reviewerUsername: string
): Promise<ReviewedMessageDetail> {
  const summary = session.index.messages.get(messageId)
  if (!summary) {
    throw createAppError(404, `Unknown message: ${messageId}`)
  }

  const review = await reviewStore.getReview(session.filePath, messageId, reviewerUsername)
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
          downloadUrl:
            attachment.isDownloadable
              ? attachment.downloadUrl || `${createAttachmentBaseUrl(session.id, messageId)}${attachment.index}`
              : ''
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
  hiddenRules: HiddenRuleRecord[],
  reviewerUsername: string
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
    collectedItems.map((message) => message.id),
    reviewerUsername
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
    pageItems.map((item) => item.id),
    reviewerUsername
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
  hiddenRules: HiddenRuleRecord[],
  reviewerUsername: string
): Promise<ReviewedFolderPage> {
  return buildReviewedFolderPage(
    session,
    folderId,
    options,
    reviewStore,
    hiddenRules,
    reviewerUsername
  )
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

function responseText(
  res: express.Response,
  statusCode: number,
  contentType: string,
  fileName: string,
  content: string
): void {
  const body = Buffer.from(content, 'utf8')
  res.status(statusCode)
    .type(contentType)
    .set('Content-Disposition', `attachment; filename="${fileName}"`)
    .set('Content-Length', String(body.length))
    .send(content)
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

function buildSmtpFromAddress(settings: SmtpSettingsRecord): string {
  const fromAddress = normalizeText(settings.fromAddress)
  const fromName = normalizeText(settings.fromName).replace(/"/g, '\\"')
  if (!fromName) {
    return fromAddress
  }

  return `${fromName} <${fromAddress}>`
}

function createDefaultSmtpTransport(settings: SmtpSettingsRecord): SmtpTransportLike {
  return nodemailer.createTransport({
    host: normalizeText(settings.host),
    port: Number(settings.port) || 587,
    secure: Boolean(settings.secure),
    auth:
      normalizeText(settings.username) || normalizeText(settings.password)
        ? {
            user: normalizeText(settings.username),
            pass: normalizeText(settings.password)
          }
        : undefined
  }) as unknown as SmtpTransportLike
}

export function createPstReviewApp(options: CreatePstReviewAppOptions): express.Express {
  const app = express()
  const sessions = new Map<string, SessionRecord>()
  let searchIndexRefreshInProgress = false
  const publicDir = options.publicDir
  const pstRootDir = options.pstRootDir || getDefaultPstRootDirectory()
  const reviewStore = options.reviewStore
  const searchIndexStore = options.searchIndexStore
  const authConfig = normalizeAuthConfig(options.auth)
  const authUserStore = options.authUserStore || createMemoryAuthUserStore(authConfig.seedUsers)
  const appSettingsStore = options.appSettingsStore || createMemoryAppSettingsStore()
  const auditLogStore = options.auditLogDir ? createFileAuditLogStore(options.auditLogDir) : null
  const smtpTransportFactory: SmtpTransportFactory =
    options.smtpTransportFactory || createDefaultSmtpTransport
  const authSessions = new Map<string, AuthSessionRecord>()
  const authMfaChallenges = new Map<string, AuthMfaChallengeRecord>()

  function getRequestPathname(req: express.Request): string {
    return (req.originalUrl || req.url || '').split('?')[0] || ''
  }

  function buildAuditActor(
    session: AuthSessionRecord | null,
    fallbackUsername = 'anonymous'
  ): AuditActor {
    return {
      username: normalizeText(session?.username || fallbackUsername) || 'anonymous',
      authenticated: Boolean(session),
      admin: isAdminAuthSession(session, authConfig)
    }
  }

  function buildAuditRequest(req: express.Request): AuditLogEntry['request'] {
    const info = getRequestInfo(req, options.apiSecurity?.webChecks)
    return {
      method: normalizeText(info.method || req.method || ''),
      path: getRequestPathname(req),
      origin: normalizeOrigin(info.origin || req.headers.origin || ''),
      ip: normalizeIpAddress(info.ip || req.ip || req.socket?.remoteAddress || '')
    }
  }

  function recordAuditEvent(input: {
    req: express.Request
    actor?: AuditActor
    session?: AuthSessionRecord | null
    action: string
    target: string
    outcome: AuditLogEntry['outcome']
    metadata?: Record<string, unknown>
  }): void {
    if (!auditLogStore) {
      return
    }

    const actor = input.actor || buildAuditActor(input.session || null)
    void auditLogStore
      .append({
        timestamp: new Date().toISOString(),
        actor,
        action: normalizeText(input.action),
        target: normalizeText(input.target),
        outcome: input.outcome,
        request: buildAuditRequest(input.req),
        metadata: input.metadata && typeof input.metadata === 'object' ? input.metadata : {}
      })
      .catch((error) => {
        console.warn('Unable to write activity log entry:', error)
      })
  }

  async function listRecentAuditEntries(
    limit: number,
    actorUsername = ''
  ): Promise<AuditLogEntry[]> {
    if (!auditLogStore) {
      return []
    }

    return auditLogStore.listRecent(limit, actorUsername)
  }

  async function listAllAuditEntries(actorUsername = ''): Promise<AuditLogEntry[]> {
    if (!auditLogStore) {
      return []
    }

    return auditLogStore.listAll(actorUsername)
  }

  function csvCell(value: unknown): string {
    const normalized = String(value ?? '')
    const safeValue = normalized && /^[=+\-@]/.test(normalized) ? `'${normalized}` : normalized
    return `"${safeValue.replace(/"/g, '""')}"`
  }

  function buildActivityLogCsv(entries: AuditLogEntry[]): string {
    const header = [
      'timestamp',
      'actorUsername',
      'actorAuthenticated',
      'actorAdmin',
      'action',
      'target',
      'outcome',
      'requestMethod',
      'requestPath',
      'requestOrigin',
      'requestIp',
      'metadataJson'
    ]

    const lines = entries.map((entry) =>
      [
        entry.timestamp,
        entry.actor.username,
        entry.actor.authenticated,
        entry.actor.admin,
        entry.action,
        entry.target,
        entry.outcome,
        entry.request.method,
        entry.request.path,
        entry.request.origin,
        entry.request.ip,
        JSON.stringify(entry.metadata || {})
      ]
        .map(csvCell)
        .join(',')
    )

    return [header.join(','), ...lines].join('\r\n')
  }

  function buildActivityLogCsvFileName(actorUsername = ''): string {
    const normalized = normalizeText(actorUsername)
    if (!normalized) {
      return 'activity-log.csv'
    }

    const safeSegment = normalized.replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
    return `activity-log-${safeSegment || 'user'}.csv`
  }

  function getAuthMfaChallengeCookieName(): string {
    return `${authConfig.cookieName}${DEFAULT_AUTH_MFA_CHALLENGE_COOKIE_SUFFIX}`
  }

  function cleanupExpiredAuthSessions(): void {
    if (!authConfig.enabled) {
      return
    }

    const now = Date.now()
    for (const [token, session] of authSessions.entries()) {
      if (session.expiresAt <= now) {
        authSessions.delete(token)
      }
    }
    for (const [token, challenge] of authMfaChallenges.entries()) {
      if (challenge.expiresAt <= now) {
        authMfaChallenges.delete(token)
      }
    }
  }

  function getAuthSessionFromRequest(req: express.Request): AuthSessionRecord | null {
    if (!authConfig.enabled) {
      return null
    }

    cleanupExpiredAuthSessions()
    const token = getCookieValue(req, authConfig.cookieName)
    if (!token) {
      return null
    }

    const session = authSessions.get(token) || null
    if (!session) {
      return null
    }

    if (session.expiresAt <= Date.now()) {
      authSessions.delete(token)
      return null
    }

    return session
  }

  function getAuthMfaChallengeFromRequest(req: express.Request): AuthMfaChallengeRecord | null {
    if (!authConfig.enabled) {
      return null
    }

    cleanupExpiredAuthSessions()
    const token = getCookieValue(req, getAuthMfaChallengeCookieName())
    if (!token) {
      return null
    }

    const challenge = authMfaChallenges.get(token) || null
    if (!challenge) {
      return null
    }

    if (challenge.expiresAt <= Date.now()) {
      authMfaChallenges.delete(token)
      return null
    }

    return challenge
  }

  function buildAuthStatus(
    session: AuthSessionRecord | null,
    challenge: AuthMfaChallengeRecord | null = null,
    mfaEnabled = false,
    mfaEnforced = false
  ): AuthStatusResponse {
    const canManageUsers = isAdminAuthSession(session, authConfig)
    if (!authConfig.enabled) {
      return {
        authenticated: true,
        enabled: false,
        canManageUsers: false,
        mfaEnabled: false,
        mfaEnforced: false,
        mfaRequired: false,
        mfaChallengeExpiresAt: null,
        user: null,
        expiresAt: null
      }
    }

    if (!session) {
      if (challenge) {
        return {
          authenticated: false,
          enabled: true,
          canManageUsers: false,
          mfaEnabled: false,
          mfaEnforced: Boolean(mfaEnforced),
          mfaRequired: true,
          mfaChallengeExpiresAt: new Date(challenge.expiresAt).toISOString(),
          user: {
            username: challenge.username
          },
          expiresAt: null
        }
      }

      return {
        authenticated: false,
        enabled: true,
        canManageUsers,
        mfaEnabled: false,
        mfaEnforced: false,
        mfaRequired: false,
        mfaChallengeExpiresAt: null,
        user: null,
        expiresAt: null
      }
    }

    return {
      authenticated: true,
      enabled: true,
      canManageUsers,
      mfaEnabled: Boolean(mfaEnabled),
      mfaEnforced: Boolean(mfaEnforced),
      mfaRequired: false,
      mfaChallengeExpiresAt: null,
      user: {
        username: session.username
      },
      expiresAt: new Date(session.expiresAt).toISOString()
    }
  }

  function createAuthSession(username: string, mfaEnabled = false): AuthSessionRecord {
    const session: AuthSessionRecord = {
      token: randomBytes(24).toString('hex'),
      username,
      expiresAt: Date.now() + authConfig.sessionTtlMinutes * 60 * 1000,
      mfaEnabled: Boolean(mfaEnabled)
    }
    authSessions.set(session.token, session)
    return session
  }

  function updateAuthSessionsMfaEnabled(username: string, mfaEnabled: boolean): number {
    if (!authConfig.enabled) {
      return 0
    }

    const normalizedUsername = normalizeAuthUsernameKey(username)
    if (!normalizedUsername) {
      return 0
    }

    let updatedCount = 0
    for (const session of authSessions.values()) {
      if (normalizeAuthUsernameKey(session.username) !== normalizedUsername) {
        continue
      }
      session.mfaEnabled = Boolean(mfaEnabled)
      updatedCount += 1
    }

    return updatedCount
  }

  function createAuthMfaChallenge(username: string): AuthMfaChallengeRecord {
    const challenge: AuthMfaChallengeRecord = {
      token: randomBytes(24).toString('hex'),
      username,
      expiresAt: Date.now() + 10 * 60 * 1000
    }
    authMfaChallenges.set(challenge.token, challenge)
    return challenge
  }

  function clearAuthSession(req: express.Request): void {
    if (!authConfig.enabled) {
      return
    }

    const token = getCookieValue(req, authConfig.cookieName)
    if (token) {
      authSessions.delete(token)
    }
  }

  function clearAuthMfaChallenge(req: express.Request): void {
    if (!authConfig.enabled) {
      return
    }

    const token = getCookieValue(req, getAuthMfaChallengeCookieName())
    if (token) {
      authMfaChallenges.delete(token)
    }
  }

  function revokeAuthSessionsForUsername(username: string): number {
    if (!authConfig.enabled) {
      return 0
    }

    const normalizedUsername = normalizeAuthUsernameKey(username)
    if (!normalizedUsername) {
      return 0
    }

    let revokedCount = 0
    for (const [token, session] of authSessions.entries()) {
      if (normalizeAuthUsernameKey(session.username) !== normalizedUsername) {
        continue
      }
      authSessions.delete(token)
      revokedCount += 1
    }

    return revokedCount
  }

  function revokeAuthMfaChallengesForUsername(username: string): number {
    if (!authConfig.enabled) {
      return 0
    }

    const normalizedUsername = normalizeAuthUsernameKey(username)
    if (!normalizedUsername) {
      return 0
    }

    let revokedCount = 0
    for (const [token, challenge] of authMfaChallenges.entries()) {
      if (normalizeAuthUsernameKey(challenge.username) !== normalizedUsername) {
        continue
      }
      authMfaChallenges.delete(token)
      revokedCount += 1
    }

    return revokedCount
  }

  function setAuthCookie(res: express.Response, req: express.Request, session: AuthSessionRecord): void {
    if (!authConfig.enabled) {
      return
    }

    res.cookie(authConfig.cookieName, session.token, {
      httpOnly: true,
      sameSite: 'lax',
      secure: Boolean(req.secure || String(req.headers['x-forwarded-proto'] || '').split(',')[0] === 'https'),
      path: '/',
      maxAge: authConfig.sessionTtlMinutes * 60 * 1000
    })
  }

  function clearAuthCookie(res: express.Response): void {
    if (!authConfig.enabled) {
      return
    }

    res.clearCookie(authConfig.cookieName, {
      path: '/',
      sameSite: 'lax'
    })
  }

  function setAuthMfaChallengeCookie(
    res: express.Response,
    req: express.Request,
    challenge: AuthMfaChallengeRecord
  ): void {
    if (!authConfig.enabled) {
      return
    }

    res.cookie(getAuthMfaChallengeCookieName(), challenge.token, {
      httpOnly: true,
      sameSite: 'lax',
      secure: Boolean(req.secure || String(req.headers['x-forwarded-proto'] || '').split(',')[0] === 'https'),
      path: '/',
      maxAge: 10 * 60 * 1000
    })
  }

  function clearAuthMfaChallengeCookie(res: express.Response): void {
    if (!authConfig.enabled) {
      return
    }

    res.clearCookie(getAuthMfaChallengeCookieName(), {
      path: '/',
      sameSite: 'lax'
    })
  }

  function getRequestBaseOrigin(req: express.Request): string {
    const info = getRequestInfo(req, options.apiSecurity?.webChecks)
    const requestOrigin = normalizeOrigin(info.origin || req.headers.origin || '')
    if (requestOrigin) {
      return requestOrigin
    }

    const forwardedProto = normalizeText(req.headers['x-forwarded-proto'] || '')
      .split(',')[0]
      .trim()
    const protocol = forwardedProto || req.protocol || 'http'
    const forwardedHost = normalizeText(req.headers['x-forwarded-host'] || '')
    const host = forwardedHost || normalizeText(req.headers.host || '')
    if (!host) {
      return ''
    }

    return normalizeOrigin(`${protocol}://${host}`)
  }

  function buildInvitePath(token: string): string {
    return `/invite/${encodeURIComponent(token)}`
  }

  function buildInviteUrl(req: express.Request, token: string): string {
    const origin =
      normalizeOrigin(authConfig.publicBaseUrl) || getRequestBaseOrigin(req) || canonicalRequestOrigin(req)
    const path = buildInvitePath(token)
    return origin ? `${origin}${path}` : path
  }

  function buildInviteEmailText(input: {
    username: string
    inviteUrl: string
    inviteExpiresAt: string
  }): string {
    return [
      `You have been invited to DV PST Mail Explorer as ${input.username}.`,
      '',
      'Click here to setup your access to DV PST Mail Explorer:',
      input.inviteUrl,
      '',
      `This invite expires at ${input.inviteExpiresAt}.`
    ].join('\n')
  }

  function escapeHtml(value: string): string {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
  }

  function buildInviteEmailHtml(input: {
    username: string
    inviteUrl: string
    inviteExpiresAt: string
  }): string {
    const inviteLinkText = 'Click here to setup your access to DV PST Mail Explorer'
    return [
      '<!doctype html>',
      '<html>',
      '<body style="margin:0;padding:0;background:#f4f7fb;color:#1f2937;font-family:Arial,Helvetica,sans-serif;">',
      '<div style="max-width:640px;margin:0 auto;padding:32px;">',
      '<div style="background:#ffffff;border:1px solid #d7e0ee;border-radius:20px;padding:28px;">',
      `<p style="margin:0 0 12px;font-size:16px;line-height:24px;">You have been invited to <strong>DV PST Mail Explorer</strong> as ${escapeHtml(
        input.username
      )}.</p>`,
      `<p style="margin:0 0 16px;font-size:16px;line-height:24px;"><a href="${escapeHtml(
        input.inviteUrl
      )}" style="color:#2f6feb;text-decoration:underline;">${escapeHtml(inviteLinkText)}</a></p>`,
      `<p style="margin:0;font-size:14px;line-height:22px;color:#4b5563;">This invite expires at ${escapeHtml(
        input.inviteExpiresAt
      )}.</p>`,
      '</div>',
      '</div>',
      '</body>',
      '</html>'
    ].join('')
  }

  async function sendInviteEmail(input: {
    recipientEmail: string
    username: string
    inviteUrl: string
    inviteExpiresAt: string
    req: express.Request
  }): Promise<{ emailSent: boolean; error?: string }> {
    try {
      const settings = await appSettingsStore.getSmtpSettings()
      if (
        !settings.enabled ||
        !normalizeText(settings.host) ||
        !normalizeText(settings.fromAddress)
      ) {
        return { emailSent: false, error: 'SMTP is not configured' }
      }

      const transporter = smtpTransportFactory(settings)
      try {
        await transporter.sendMail({
          from: buildSmtpFromAddress(settings),
          to: input.recipientEmail,
          subject: 'PST Mail Explorer invitation',
          text: buildInviteEmailText({
            username: input.username,
            inviteUrl: input.inviteUrl,
            inviteExpiresAt: input.inviteExpiresAt
          }),
          html: buildInviteEmailHtml({
            username: input.username,
            inviteUrl: input.inviteUrl,
            inviteExpiresAt: input.inviteExpiresAt
          })
        })
        return { emailSent: true }
      } finally {
        try {
          await transporter.close?.()
        } catch {
          // Ignore transport close failures.
        }
      }
    } catch (error) {
      return {
        emailSent: false,
        error: error instanceof Error ? error.message : String(error)
      }
    }
  }

  function createAuthGateMiddleware(): express.RequestHandler {
    const allowedOrigins = new Set(
      parseList(options.apiSecurity?.allowedOrigins).map(normalizeOrigin).filter(Boolean)
    )

    return (req, res, next) => {
      const pathname = (req.originalUrl || req.url || '').split('?')[0] || ''
      if (!isProtectedApiPath(pathname)) {
        return next()
      }

      if (!authConfig.enabled) {
        return next()
      }

      const info = getRequestInfo(req, options.apiSecurity?.webChecks)
      const session = getAuthSessionFromRequest(req)
      const requestOrigin = normalizeOrigin(info.origin || req.headers.origin || '')
      const requestHostOrigin = canonicalRequestOrigin(req)
      const isSameOrigin = Boolean(requestOrigin && requestOrigin === requestHostOrigin)
      const originAllowed = !requestOrigin || isSameOrigin || allowedOrigins.has(requestOrigin)

      if (!originAllowed) {
        recordAuditEvent({
          req,
          session,
          action: 'access.denied',
          target: pathname,
          outcome: 'denied',
          metadata: {
            reason: 'CORS origin not allowed'
          }
        })
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

      if (!session) {
        recordAuditEvent({
          req,
          actor: buildAuditActor(null),
          action: 'access.denied',
          target: pathname,
          outcome: 'denied',
          metadata: {
            reason: 'Authentication required'
          }
        })
        res.set('Cache-Control', 'no-store')
        return res.status(401).json({
          error: 'Authentication required'
        })
      }

      return next()
    }
  }

  function closeSessionsForMailboxKey(mailboxKey: string): string[] {
    const closedSessionIds: string[] = []
    const normalizedMailboxKey = normalizeText(mailboxKey)
    for (const [sessionId, session] of sessions.entries()) {
      if (session.filePath === normalizedMailboxKey) {
        sessions.delete(sessionId)
        closedSessionIds.push(sessionId)
      }
    }
    return closedSessionIds
  }

  app.disable('x-powered-by')
  app.use(express.json({ limit: '2mb' }))
  app.use(express.static(publicDir))
  app.use(createAuthGateMiddleware())
  app.use(
    createApiSecurityMiddleware({
      ...options.apiSecurity,
      hasLocalAuthSession: (req) => Boolean(authConfig.enabled && getAuthSessionFromRequest(req))
    })
  )

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

  app.get('/invite/:token', (_req, res) => {
    res.status(200).sendFile(path.join(publicDir, 'index.html'))
  })

  app.get(API_ROUTES.authMe, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const session = getAuthSessionFromRequest(req)
      const challenge = session ? null : getAuthMfaChallengeFromRequest(req)
      if (!authConfig.enabled && !session && !challenge) {
        responseJson(res, 200, buildAuthStatus(null))
        return
      }

      if (!session && !challenge) {
        responseJson(res, 401, {
          ...buildAuthStatus(null),
          error: 'Authentication required'
        })
        return
      }

      const currentUser = session
        ? await authUserStore.getUser(session.username)
        : challenge
          ? await authUserStore.getUser(challenge.username)
          : null
      responseJson(
        res,
        200,
        buildAuthStatus(
          session,
          challenge,
          session?.mfaEnabled ?? Boolean(currentUser?.mfaEnabled),
          Boolean(currentUser?.mfaEnforced)
        )
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authLogin, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 200, buildAuthStatus(null))
        return
      }

      const body = (req.body || {}) as { username?: string; password?: string }
      const username = normalizeAuthUsername(body.username)
      const password = String(body.password ?? '')
      const user = await authUserStore.authenticate(username, password)
      if (!user) {
        recordAuditEvent({
          req,
          actor: {
            username: username || 'anonymous',
            authenticated: false,
            admin: false
          },
          action: 'auth.login',
          target: username || 'anonymous',
          outcome: 'denied',
          metadata: {
            reason: 'Invalid username or password'
          }
        })
        responseJson(res, 401, {
          ...buildAuthStatus(null),
          error: 'Invalid username or password'
        })
        return
      }

      if (user.mfaEnabled) {
        const challenge = createAuthMfaChallenge(user.username)
        setAuthMfaChallengeCookie(res, req, challenge)
        recordAuditEvent({
          req,
          actor: buildAuditActor(null, user.username),
          action: 'auth.login',
          target: user.username,
          outcome: 'success',
          metadata: {
            mfaRequired: true,
            mfaEnforced: Boolean(user.mfaEnforced)
          }
        })
        responseJson(res, 200, buildAuthStatus(null, challenge, false, Boolean(user.mfaEnforced)))
        return
      }

      const session = createAuthSession(user.username, Boolean(user.mfaEnabled))
      setAuthCookie(res, req, session)
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, user.username),
        action: 'auth.login',
        target: user.username,
        outcome: 'success',
        metadata: {
          mfaEnforced: Boolean(user.mfaEnforced)
        }
      })
      responseJson(res, 200, buildAuthStatus(session, null, Boolean(user.mfaEnabled), Boolean(user.mfaEnforced)))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authLogout, (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (authConfig.enabled) {
        const session = getAuthSessionFromRequest(req)
        const challenge = getAuthMfaChallengeFromRequest(req)
        if (session) {
          recordAuditEvent({
            req,
            session,
            action: 'auth.logout',
            target: session.username,
            outcome: 'success'
          })
        }
        clearAuthSession(req)
        clearAuthMfaChallenge(req)
        clearAuthCookie(res)
        clearAuthMfaChallengeCookie(res)
      }
      responseJson(res, 200, buildAuthStatus(null))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.authUsers, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 200, { users: [] } satisfies AuthUsersResponse)
        return
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.list',
          target: 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const users = await authUserStore.listUsers()
      recordAuditEvent({
        req,
        session,
        action: 'auth.users.list',
        target: 'local users',
        outcome: 'success',
        metadata: {
          count: users.length
        }
      })
      responseJson(res, 200, {
        users
      } satisfies AuthUsersResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authUsers, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.create',
          target: 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const body = (req.body || {}) as {
        username?: string
        password?: string
        email?: string
        recipientEmail?: string
      }
      const username = normalizeAuthUsername(body.username)
      const recipientEmail = normalizeText(body.recipientEmail || body.email)
      if (recipientEmail) {
        const result = await authUserStore.createInvite(
          username,
          recipientEmail,
          authConfig.inviteTtlMinutes
        )
        const inviteUrl = buildInviteUrl(req, result.inviteToken)
        const emailResult = await sendInviteEmail({
          recipientEmail,
          username: result.user.username,
          inviteUrl,
          inviteExpiresAt: result.inviteExpiresAt,
          req
        })
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.invite.create',
          target: result.user.username,
          outcome: 'success',
          metadata: {
            recipientEmail,
            inviteExpiresAt: result.inviteExpiresAt,
            emailSent: emailResult.emailSent
          }
        })
        responseJson(res, 200, {
          user: result.user,
          inviteUrl,
          emailSent: emailResult.emailSent,
          inviteExpiresAt: result.inviteExpiresAt
        } satisfies AuthUserCreateResponse)
        return
      }

      const user = await authUserStore.addUser(username, String(body.password ?? ''))
      recordAuditEvent({
        req,
        session,
        action: 'auth.users.create',
        target: user.username,
        outcome: 'success'
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserCreateResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.authUser, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const targetUsername = normalizeAuthUsername(req.params.username)
      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.delete',
          target: targetUsername || 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      if (!targetUsername) {
        responseJson(res, 400, {
          error: 'Username is required'
        })
        return
      }

      if (normalizeAuthUsernameKey(targetUsername) === normalizeAuthUsernameKey(authConfig.username)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.delete',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'Admin account cannot be deleted'
          }
        })
        responseJson(res, 400, {
          error: 'Admin account cannot be deleted'
        })
        return
      }

      const user = await authUserStore.deleteUser(targetUsername)
      if (!user) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.delete',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found'
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }

      const revokedSessions = revokeAuthSessionsForUsername(user.username)
      const revokedChallenges = revokeAuthMfaChallengesForUsername(user.username)
      recordAuditEvent({
        req,
        session,
        action: 'auth.users.delete',
        target: user.username,
        outcome: 'success',
        metadata: {
          revokedSessions,
          revokedChallenges
        }
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authUserInviteResend, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.invite.resend',
          target: normalizeAuthUsername(req.params.username) || 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const targetUsername = normalizeAuthUsername(req.params.username)
      if (!targetUsername) {
        responseJson(res, 400, {
          error: 'Username is required'
        })
        return
      }

      const result = await authUserStore.resendInvite(targetUsername, authConfig.inviteTtlMinutes)
      if (!result) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.invite.resend',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found'
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }

      const inviteUrl = buildInviteUrl(req, result.inviteToken)
      const emailResult = await sendInviteEmail({
        recipientEmail: result.user.recipientEmail,
        username: result.user.username,
        inviteUrl,
        inviteExpiresAt: result.inviteExpiresAt,
        req
      })
      recordAuditEvent({
        req,
        session,
        action: 'auth.users.invite.resend',
        target: result.user.username,
        outcome: 'success',
        metadata: {
          recipientEmail: result.user.recipientEmail,
          inviteExpiresAt: result.inviteExpiresAt,
          emailSent: emailResult.emailSent
        }
      })
      responseJson(res, 200, {
        user: result.user,
        inviteUrl,
        emailSent: emailResult.emailSent,
        inviteExpiresAt: result.inviteExpiresAt
      } satisfies AuthUserCreateResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.authUserInvite, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.invite.revoke',
          target: normalizeAuthUsername(req.params.username) || 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const targetUsername = normalizeAuthUsername(req.params.username)
      if (!targetUsername) {
        responseJson(res, 400, {
          error: 'Username is required'
        })
        return
      }

      const user = await authUserStore.revokeInvite(targetUsername)
      if (!user) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.invite.revoke',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found'
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }

      recordAuditEvent({
        req,
        session,
        action: 'auth.users.invite.revoke',
        target: user.username,
        outcome: 'success'
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authUserMfaReset, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.mfa.reset',
          target: normalizeAuthUsername(req.params.username) || 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const targetUsername = normalizeAuthUsername(req.params.username)
      if (!targetUsername) {
        responseJson(res, 400, {
          error: 'Username is required'
        })
        return
      }

      const user = await authUserStore.resetMfa(targetUsername)
      if (!user) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.mfa.reset',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found'
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }

      const revokedSessions = revokeAuthSessionsForUsername(user.username)
      const revokedChallenges = revokeAuthMfaChallengesForUsername(user.username)
      recordAuditEvent({
        req,
        session,
        action: 'auth.users.mfa.reset',
        target: user.username,
        outcome: 'success',
        metadata: {
          revokedSessions,
          revokedChallenges
        }
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authUserMfaEnforce, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.mfa.enforce',
          target: normalizeAuthUsername(req.params.username) || 'local users',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const targetUsername = normalizeAuthUsername(req.params.username)
      if (!targetUsername) {
        responseJson(res, 400, {
          error: 'Username is required'
        })
        return
      }

      const body = (req.body || {}) as AuthMfaEnforceRequestBody
      const enforced = Boolean(body.enforced)
      const user = await authUserStore.setMfaEnforced(targetUsername, enforced)
      if (!user) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.mfa.enforce',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found',
            enforced
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }

      recordAuditEvent({
        req,
        session,
        action: 'auth.users.mfa.enforce',
        target: user.username,
        outcome: 'success',
        metadata: {
          enforced,
          mfaEnabled: Boolean(user.mfaEnabled)
        }
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.authInviteLookup, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 400, {
          error: 'Authentication is disabled'
        })
        return
      }

      const token = normalizeText(req.params.token)
      if (!token) {
        responseJson(res, 400, {
          error: 'Invite token is required'
        })
        return
      }

      const invite = await authUserStore.getInviteByToken(token)
      if (!invite) {
        responseJson(res, 404, {
          error: 'Invite not found'
        })
        return
      }

      responseJson(res, 200, {
        invite
      } satisfies AuthInviteLookupResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authInviteAccept, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const token = normalizeText(req.params.token)
      if (!token) {
        responseJson(res, 400, {
          error: 'Invite token is required'
        })
        return
      }

      const body = (req.body || {}) as { password?: string; confirmPassword?: string }
      const password = String(body.password ?? '')
      const confirmPassword = String(body.confirmPassword ?? '')
      if (confirmPassword && password !== confirmPassword) {
        responseJson(res, 400, {
          error: 'Passwords do not match'
        })
        return
      }

      const user = await authUserStore.acceptInvite(token, password)
      const session = createAuthSession(user.username, Boolean(user.mfaEnabled))
      setAuthCookie(res, req, session)
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, user.username),
        action: 'auth.invite.accept',
        target: user.username,
        outcome: 'success',
        metadata: {
          mfaEnabled: Boolean(user.mfaEnabled),
          mfaEnforced: Boolean(user.mfaEnforced)
        }
      })
      responseJson(res, 200, {
        user,
        mfaAvailable: !Boolean(user.mfaEnabled)
      } satisfies AuthInviteAcceptResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authMfaChallenge, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 200, buildAuthStatus(null))
        return
      }

      const challenge = getAuthMfaChallengeFromRequest(req)
      if (!challenge) {
        responseJson(res, 401, {
          error: 'MFA challenge required'
        })
        return
      }

      const body = (req.body || {}) as { code?: string }
      const code = String(body.code ?? '')
      const user = await authUserStore.verifyMfaChallenge(challenge.username, code)
      if (!user) {
        recordAuditEvent({
          req,
          actor: {
            username: challenge.username,
            authenticated: false,
            admin: false
          },
          action: 'auth.mfa.challenge',
          target: challenge.username,
          outcome: 'denied',
          metadata: {
            reason: 'Invalid verification code'
          }
        })
        responseJson(res, 401, {
          error: 'Invalid verification code'
        })
        return
      }

      const session = createAuthSession(user.username, Boolean(user.mfaEnabled))
      setAuthCookie(res, req, session)
      clearAuthMfaChallenge(req)
      clearAuthMfaChallengeCookie(res)
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, user.username),
        action: 'auth.mfa.challenge',
        target: user.username,
        outcome: 'success'
      })
      responseJson(res, 200, buildAuthStatus(session, null, Boolean(user.mfaEnabled), Boolean(user.mfaEnforced)))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authMfaEnrollmentStart, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const session = getAuthSessionFromRequest(req)
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const result = await authUserStore.startMfaEnrollment(session.username, authConfig.mfaIssuer)
      const qrCodeDataUrl = await QRCode.toDataURL(result.otpauthUri, {
        margin: 1,
        width: 240
      })
      recordAuditEvent({
        req,
        session,
        action: 'auth.mfa.enroll.start',
        target: session.username,
        outcome: 'success'
      })
      responseJson(res, 200, {
        user: result.user,
        secret: result.secret,
        otpauthUri: result.otpauthUri,
        qrCodeDataUrl
      } satisfies AuthMfaStartResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authMfaEnrollmentComplete, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const session = getAuthSessionFromRequest(req)
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const body = (req.body || {}) as { code?: string }
      const code = String(body.code ?? '')
      const result = await authUserStore.completeMfaEnrollment(session.username, code)
      updateAuthSessionsMfaEnabled(session.username, true)
      recordAuditEvent({
        req,
        session,
        action: 'auth.mfa.enroll.complete',
        target: session.username,
        outcome: 'success',
        metadata: {
          recoveryCodeCount: result.recoveryCodes.length
        }
      })
      responseJson(res, 200, {
        user: result.user,
        recoveryCodes: result.recoveryCodes
      } satisfies AuthMfaCompleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.smtpSettings, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 400, {
          error: 'Authentication is disabled'
        })
        return
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.read',
          target: 'smtp settings',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      responseJson(res, 200, {
        settings: buildSmtpSettingsView(await appSettingsStore.getSmtpSettings())
      } satisfies SmtpSettingsResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.put(API_ROUTES.smtpSettings, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 400, {
          error: 'Authentication is disabled'
        })
        return
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.update',
          target: 'smtp settings',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const body = normalizeSmtpSettingsInput((req.body || {}) as SmtpSettingsInput)
      const updated = await appSettingsStore.updateSmtpSettings(body)
      recordAuditEvent({
        req,
        session,
        action: 'settings.smtp.update',
        target: updated.fromAddress || updated.host || 'smtp settings',
        outcome: 'success',
        metadata: {
          enabled: updated.enabled,
          host: updated.host,
          port: updated.port,
          secure: updated.secure,
          username: updated.username,
          fromName: updated.fromName,
          fromAddress: updated.fromAddress,
          replyTo: updated.replyTo,
          hasPassword: Boolean(updated.password)
        }
      })
      responseJson(res, 200, {
        settings: buildSmtpSettingsView(updated)
      } satisfies SmtpSettingsResponse)
    } catch (error) {
      const session = getAuthSessionFromRequest(req)
      if (session) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.update',
          target: 'smtp settings',
          outcome: 'failure',
          metadata: {
            reason: error instanceof Error ? error.message : String(error)
          }
        })
      }
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.smtpSettingsTest, async (req, res) => {
    let session: AuthSessionRecord | null = null
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 400, {
          error: 'Authentication is disabled'
        })
        return
      }

      session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.test',
          target: 'smtp settings',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const rawBody = (req.body || {}) as SmtpTestRequestBody
      const body = normalizeSmtpSettingsInput(rawBody)
      const recipient = normalizeText(rawBody.recipient)
      if (!recipient) {
        throw createAppError(400, 'Test recipient is required')
      }

      const baseSettings = await appSettingsStore.getSmtpSettings()
      const settings = mergeSmtpSettings(baseSettings, body)
      if (!normalizeText(settings.host)) {
        throw createAppError(400, 'SMTP host is required')
      }
      if (!normalizeText(settings.fromAddress)) {
        throw createAppError(400, 'SMTP from address is required')
      }

      const transporter = smtpTransportFactory(settings)
      try {
        let result: SmtpTransportSendResult | null = null
        try {
          result = await transporter.sendMail({
            from: buildSmtpFromAddress(settings),
            to: recipient,
            subject: 'PST Mail Explorer SMTP test',
            text:
              `This is a test email from PST Mail Explorer.\n\n` +
              `SMTP host: ${settings.host}\n` +
              `SMTP port: ${settings.port}\n` +
              `SMTP secure: ${settings.secure ? 'true' : 'false'}\n` +
              `Sender: ${settings.fromAddress}\n` +
              `Recipient: ${recipient}\n` +
              `Sent at: ${new Date().toISOString()}\n`,
            replyTo: normalizeText(settings.replyTo) || undefined
          })
        } catch (error) {
          throw createAppError(
            502,
            `Unable to send test email: ${error instanceof Error ? error.message : String(error)}`
          )
        }

        if (!result) {
          throw createAppError(502, 'Unable to send test email: no transport response')
        }

        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.test',
          target: recipient,
          outcome: 'success',
          metadata: {
            host: settings.host,
            port: settings.port,
            secure: settings.secure,
            username: settings.username,
            fromAddress: settings.fromAddress,
            replyTo: settings.replyTo,
            messageId: normalizeText(result.messageId),
            accepted: result.accepted || [],
            rejected: result.rejected || []
          }
        })
        responseJson(res, 200, {
          success: true,
          recipient,
          messageId: normalizeText(result.messageId),
          accepted: result.accepted || [],
          rejected: result.rejected || []
        } satisfies SmtpTestResponse)
      } finally {
        if (transporter && typeof transporter.close === 'function') {
          try {
            await Promise.resolve(transporter.close())
          } catch {
            // Ignore transport shutdown failures after the email attempt completes.
          }
        }
      }
    } catch (error) {
      if (session) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.smtp.test',
          target: normalizeText(((req.body || {}) as SmtpTestRequestBody).recipient) || 'smtp settings',
          outcome: 'failure',
          metadata: {
            reason: error instanceof Error ? error.message : String(error)
          }
        })
      }
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.activityLog, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 200, {
          entries: []
        } satisfies ActivityLogResponse)
        return
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'access.denied',
          target: getRequestPathname(req),
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const limit = Math.min(parsePositiveInt(req.query.limit as string | string[] | undefined, 50), 200)
      const actorUsername =
        typeof req.query.username === 'string' ? normalizeAuthUsername(req.query.username) : ''
      responseJson(res, 200, {
        entries: await listRecentAuditEntries(limit, actorUsername)
      } satisfies ActivityLogResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.activityLogCsv, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        const csv = buildActivityLogCsv([])
        responseText(res, 200, 'text/csv; charset=utf-8', 'activity-log.csv', csv)
        return
      }

      const session = getAuthSessionFromRequest(req)
      if (!session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      if (!isAdminAuthSession(session, authConfig)) {
        recordAuditEvent({
          req,
          session,
          action: 'access.denied',
          target: getRequestPathname(req),
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const actorUsername =
        typeof req.query.username === 'string' ? normalizeAuthUsername(req.query.username) : ''
      const entries = await listAllAuditEntries(actorUsername)
      const csv = buildActivityLogCsv(entries)
      responseText(
        res,
        200,
        'text/csv; charset=utf-8',
        buildActivityLogCsvFileName(actorUsername),
        csv
      )
      recordAuditEvent({
        req,
        session,
        action: 'activity.log.export',
        target: actorUsername || 'all activity',
        outcome: 'success',
        metadata: {
          username: actorUsername || '',
          entryCount: entries.length
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.pstCatalog, (req, res) => {
    try {
      const scopePath =
        typeof req.query.scopePath === 'string' ? normalizeText(req.query.scopePath) : ''
      responseJson(res, 200, listPstMailboxFiles(pstRootDir, scopePath))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.pstRemovedCatalog, (req, res) => {
    try {
      const scopePath =
        typeof req.query.scopePath === 'string' ? normalizeText(req.query.scopePath) : ''
      responseJson(res, 200, listRemovedPstMailboxFiles(pstRootDir, scopePath))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.pstOpen, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
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

      const summary = buildSessionSummary(index)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'mailbox.open',
        target: `${scopePath ? `${scopePath}/` : ''}${fileName}`,
        outcome: 'success',
        metadata: {
          scopePath,
          scopeLabel,
          fileName,
          messageCount: index.messages.size
        }
      })
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

  app.post(API_ROUTES.pstRemove, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const body = (req.body || {}) as MoveMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      const removal = movePstMailboxToRemoved(pstRootDir, scopePath, fileName)
      await searchIndexStore.deleteMailboxDocuments(removal.sourcePath)
      const closedSessionIds = closeSessionsForMailboxKey(removal.sourcePath)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'mailbox.remove',
        target: `${removal.scopePath ? `${removal.scopePath}/` : ''}${removal.fileName}`,
        outcome: 'success',
        metadata: {
          scopePath: removal.scopePath,
          scopeLabel: getScopeLabel(removal.scopePath),
          fileName: removal.fileName,
          closedSessionCount: closedSessionIds.length
        }
      })

      responseJson(res, 200, {
        removed: {
          sourcePath: removal.sourcePath,
          destinationPath: removal.destinationPath,
          scopePath: removal.scopePath,
          scopeLabel: getScopeLabel(removal.scopePath),
          fileName: removal.fileName
        },
        closedSessionIds
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.pstRestore, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const body = (req.body || {}) as MoveMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      const restore = restorePstMailboxFromRemoved(pstRootDir, scopePath, fileName)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'mailbox.restore',
        target: `${restore.scopePath ? `${restore.scopePath}/` : ''}${restore.fileName}`,
        outcome: 'success',
        metadata: {
          scopePath: restore.scopePath,
          scopeLabel: getScopeLabel(restore.scopePath),
          fileName: restore.fileName
        }
      })

      responseJson(res, 200, {
        restored: {
          sourcePath: restore.sourcePath,
          destinationPath: restore.destinationPath,
          scopePath: restore.scopePath,
          scopeLabel: getScopeLabel(restore.scopePath),
          fileName: restore.fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.search, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const scope = parseSearchScope(req.query.scope) as SearchScope
      const requestedScopePath = normalizeScopePath(req.query.scopePath)
      const sessionId = normalizeText(req.query.sessionId)
      let scopePath = ''
      let scopeLabel = 'All cases/searches'
      let mailboxKey = ''
      let allowedMailboxKeys: string[] = []

      if (scope === 'pst') {
        const session = getSessionOrThrow(sessions, sessionId)
        scopePath = session.scopePath
        scopeLabel = session.scopeLabel || getScopeLabel(scopePath)
        mailboxKey = session.filePath
        allowedMailboxKeys = [mailboxKey]
      } else if (scope === 'search') {
        const catalog = resolveCatalogScopeSelection(pstRootDir, requestedScopePath)
        scopePath = catalog.scopePath
        scopeLabel = catalog.scopeLabel
        const selectedCatalog = listPstMailboxFiles(pstRootDir, scopePath)
        allowedMailboxKeys = selectedCatalog.files
          .map((file) => resolvePstMailboxPath(pstRootDir, selectedCatalog.scopePath, file.fileName))
      } else {
        const activeCatalog = listPstMailboxFiles(pstRootDir)
        allowedMailboxKeys = activeCatalog.scopes.flatMap((entry) =>
          entry.files.map((file) => resolvePstMailboxPath(pstRootDir, entry.scopePath, file.fileName))
        )
      }

      const page = await searchIndexStore.search({
        scope,
        scopePath,
        mailboxKey,
        allowedMailboxKeys,
        reviewerUsername,
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
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.execute',
        target: scopeLabel,
        outcome: 'success',
        metadata: {
          scope,
          scopePath,
          queryLength: filters.query.length,
          mode: filters.mode,
          mailOnly: filters.mailOnly,
          sort: filters.sort,
          page: filters.page,
          pageSize: filters.pageSize,
          resultCount: page.total
        }
      })
      responseJson(res, 200, {
        scope,
        scopePath: page.scopePath || scopePath,
        scopeLabel: page.scopeLabel || scopeLabel,
        page: {
          ...page,
          items: page.items.map((item) => {
            const { reviewStates: _reviewStates, ...rest } = item as SearchIndexDocument & {
              reviewStates?: unknown
            }
            return rest
          })
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.searchIndexRefresh, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !isAdminAuthSession(authSession, authConfig)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'search.index.refresh',
          target: 'PST catalog',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required'
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      if (searchIndexRefreshInProgress) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'search.index.refresh',
          target: 'PST catalog',
          outcome: 'denied',
          metadata: {
            reason: 'Search index refresh already in progress'
          }
        })
        responseJson(res, 409, {
          error: 'Search index refresh already in progress'
        })
        return
      }

      searchIndexRefreshInProgress = true
      try {
        const summary = await refreshSearchIndexFromCatalog(pstRootDir, reviewStore, searchIndexStore)
        recordAuditEvent({
          req,
          session: authSession,
          action: 'search.index.refresh',
          target: 'PST catalog',
          outcome: 'success',
          metadata: {
            mailboxCount: summary.mailboxCount,
            messageCount: summary.messageCount
          }
        })
        responseJson(res, 200, {
          summary,
          hiddenRules: await searchIndexStore.listHiddenRules()
        })
      } finally {
        searchIndexRefreshInProgress = false
      }
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
      const authSession = getAuthSessionFromRequest(req)
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
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.filter.create',
        target: rule.filterId,
        outcome: 'success',
        metadata: {
          kind: rule.kind
        }
      })
      responseJson(res, 200, { rule })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.searchFilter, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const filterId = normalizeText(req.params.filterId)
      if (!filterId) {
        throw createAppError(400, 'Filter id is required')
      }
      const deleted = await searchIndexStore.deleteHiddenRule(filterId)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.filter.delete',
        target: filterId,
        outcome: deleted ? 'success' : 'failure'
      })
      responseJson(res, 200, { deleted })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.sessionSummary, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'mailbox.summary.view',
        target: session.fileName,
        outcome: 'success',
        metadata: {
          scopePath: session.scopePath,
          scopeLabel: session.scopeLabel
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'mailbox.tree.view',
        target: session.fileName,
        outcome: 'success',
        metadata: {
          scopePath: session.scopePath,
          scopeLabel: session.scopeLabel
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const folderId = normalizeText(req.params.folderId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const hiddenRules = await searchIndexStore.listHiddenRules()
      const page = await buildFolderPageWithReviews(
        session,
        folderId,
        filters,
        reviewStore,
        hiddenRules,
        reviewerUsername
      )
      recordAuditEvent({
        req,
        session: authSession,
        action: 'folder.view',
        target: folderId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          scopePath: session.scopePath,
          scopeLabel: session.scopeLabel,
          page: page.page,
          pageSize: page.pageSize,
          resultCount: page.items.length
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(
        session,
        messageId,
        reviewStore,
        reviewerUsername
      )
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.view',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          folderId: detail.folderId,
          folderPath: detail.folderPath
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(
        session,
        messageId,
        reviewStore,
        reviewerUsername
      )
      const fields = normalizeExtractionFields(req.query.fields)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.extract',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          fields: [...fields]
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
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
        hiddenRules,
        reviewerUsername
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
      recordAuditEvent({
        req,
        session: authSession,
        action: 'folder.extract',
        target: folderId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          fields: [...fields],
          page: page.page,
          pageSize: page.pageSize,
          resultCount: extraction.items.length
        }
      })

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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(
        session,
        messageId,
        reviewStore,
        reviewerUsername
      )
      const fileName = `${safeDownloadName(detail.subject || 'message', 'message')}.json`
      responseBinary(
        res,
        200,
        'application/json; charset=utf-8',
        fileName,
        Buffer.from(exportMessageAsJson(detail), 'utf8')
      )
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.export.json',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          downloadName: fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExportEml, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
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
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.export.eml',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          downloadName: fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.flaggedBundleExport, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
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
          flaggedOnly: true,
          reviewerUsername
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
      recordAuditEvent({
        req,
        session: authSession,
        action: 'bundle.export',
        target: scopeLabel,
        outcome: 'success',
        metadata: {
          scope,
          scopePath: manifest.scope.scopePath,
          scopeLabel: manifest.scope.scopeLabel,
          reviewerUsername,
          exported: manifest.counts.exported,
          failed: manifest.counts.failed
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const attachmentIndex = parseZeroBasedInt(req.params.attachmentIndex)
      if (attachmentIndex < 0) {
        throw createAppError(404, 'Attachment not found')
      }
      const payload = getAttachmentDownloadBuffer(session.index, messageId, attachmentIndex)
      const fileName = safeDownloadName(payload.filename, 'attachment')
      responseBinary(res, 200, payload.contentType, fileName, payload.buffer)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.attachment.download',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          attachmentIndex,
          downloadName: fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageReview, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const review = normalizeReviewState(
        await reviewStore.getReview(session.filePath, messageId, reviewerUsername)
      )
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const body = (req.body || {}) as ReviewPatchBody
      if (body.flagged === undefined && body.tags === undefined) {
        throw createAppError(400, 'Provide flagged or tags to update review state')
      }

      const summary = getMessageSummary(session.index, messageId)
      assertReviewableMessage(summary)
      const review = await reviewStore.upsertReview({
        ...toReviewableContext(session, summary, reviewerUsername),
        flagged: body.flagged,
        tags: body.tags
      })
      await searchIndexStore.updateReviewState(
        session.filePath,
        messageId,
        reviewerUsername,
        review
      )
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.review.update',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          reviewerUsername,
          flagged: review ? review.flagged : false,
          tagCount: review ? review.tags.length : 0
        }
      })

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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      await reviewStore.deleteReview(session.filePath, messageId, reviewerUsername)
      await searchIndexStore.updateReviewState(session.filePath, messageId, reviewerUsername, null)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'message.review.delete',
        target: messageId,
        outcome: 'success',
        metadata: {
          fileName: session.fileName,
          reviewerUsername
        }
      })
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
      const authSession = getAuthSessionFromRequest(req)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const pageSize = filters.pageSize
      const page = filters.page
      const records = await reviewStore.listReviews(session.filePath, {
        query: filters.query,
        flaggedOnly: filters.reviewFlaggedOnly,
        taggedOnly: filters.reviewTaggedOnly,
        tag: filters.reviewTag,
        reviewerUsername
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
