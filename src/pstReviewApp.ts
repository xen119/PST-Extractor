import express from 'express'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { createHash, createPublicKey, createVerify, randomBytes } from 'crypto'
import { once } from 'events'
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
  buildEntraSettingsView,
  type AppSettingsStore,
  type EntraSettingsInput,
  type EntraSettingsRecord,
  type EntraSettingsView,
  type SmtpSettingsInput,
  type SmtpSettingsRecord,
  createMemoryAppSettingsStore,
  mergeSmtpSettings,
  normalizeEntraSettingsInput,
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
  buildPasswordPolicyDefaultsFromEnv,
  normalizePasswordPolicyInput,
  validatePasswordAgainstPolicy,
  type PasswordPolicyInput,
  type PasswordPolicyRecord
} from './passwordPolicy'
import {
  buildReviewContext,
  type ReviewStore
} from './reviewStore'
import type { ReviewState } from './reviewTypes'
import {
  type HiddenRuleRecord,
  buildMailboxSearchDocumentId,
  type SearchIndexDocument,
  type SearchIndexFileFingerprint,
  type SearchIndexPage,
  type SearchIndexSearchOptions,
  type SearchIndexRefreshSource,
  type SearchIndexStore,
  type SearchScope
} from './searchIndex'
import {
  buildFlaggedBundleWorkspaceKey,
  createFlaggedBundleJobInProgressError,
  createMemoryFlaggedBundleStore,
  type FlaggedBundleArtifactDownload,
  type FlaggedBundleArtifactRecord,
  type FlaggedBundleExportScope,
  type FlaggedBundleGroupRecord,
  type FlaggedBundleGroupType,
  type FlaggedBundleJobRecord,
  type FlaggedBundleJobStage,
  type FlaggedBundleJobStatus,
  type FlaggedBundleProgress,
  type FlaggedBundleScope,
  type FlaggedBundleStore
} from './flaggedBundleStore'
import {
  createSearchIndexRefreshCoordinator,
  type SearchIndexRefreshCoordinator,
  type SearchIndexRefreshStatus
} from './searchIndexRefresh'
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
import {
  listArchiveBundleFiles,
  readArchiveBundleItemContent
} from './archiveBundles'
import {
  buildOfficePreview,
  isOfficePreviewable
} from './officePreview'
import type { ReviewRecord } from './reviewTypes'
import { createZipStreamWriter, estimateZipEntrySize } from './zipWriter'

export interface CreatePstReviewAppOptions {
  publicDir: string
  reviewStore: ReviewStore
  searchIndexStore: SearchIndexStore
  flaggedBundleStore?: FlaggedBundleStore
  openApiSpec: Record<string, unknown>
  pstRootDir?: string
  auth?: AppAuthConfig
  authUserStore?: AuthUserStore
  appSettingsStore?: AppSettingsStore
  auditLogDir?: string
  apiSecurity?: ApiSecurityConfig
  smtpTransportFactory?: SmtpTransportFactory
  searchIndexRefreshCoordinator?: SearchIndexRefreshCoordinator
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

interface AuthPasswordChangeChallengeRecord {
  token: string
  username: string
  expiresAt: number
}

interface AuthStatusResponse {
  authenticated: boolean
  enabled: boolean
  canManageUsers: boolean
  entraEnabled: boolean
  mfaEnabled: boolean
  mfaEnforced: boolean
  mfaRequired: boolean
  mfaChallengeExpiresAt: string | null
  lockedUntil: string | null
  loginFailedCount: number
  passwordResetAvailable: boolean
  passwordChangeRequired: boolean
  passwordChangeChallengeExpiresAt: string | null
  user: {
    username: string
    assignedCasePaths: string[]
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

interface AuthUserPasswordResetResponse {
  user: AuthUserListItem
  mode: 'link' | 'temporary'
  resetUrl?: string
  resetExpiresAt?: string
  emailSent?: boolean
  emailError?: string
  temporaryPassword?: string
}

interface EntraSettingsResponse {
  settings: EntraSettingsView
  redirectUri: string
}

interface ActivityLogResponse {
  entries: AuditLogEntry[]
}

interface SmtpSettingsResponse {
  settings: ReturnType<typeof buildSmtpSettingsView>
}

interface PasswordPolicyResponse {
  settings: PasswordPolicyRecord
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
  messageDetailCache: Map<string, Promise<ReviewedMessageDetail>>
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
  sourceType?: string
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
  collapseDuplicates?: boolean
}

interface WorkspaceItemsRequestQuery extends SearchRequestQuery {
  workspaceMode?: 'folder' | 'search'
  folderId?: string
}

interface FlaggedBundleQuery {
  scope?: string
  scopePath?: string
  sessionId?: string
}

interface FlaggedBundlePrepareBody {
  scope?: string
  scopePath?: string
  sessionId?: string
  maxSizeBytes?: number
}

interface FlaggedBundleScopeDetails extends FlaggedBundleExportScope {
  workspaceKey: string
}

interface FlaggedBundleExportArtifactResponse {
  artifactId: string
  fileName: string
  downloadUrl: string
  partNumber: number
  partCount: number
  itemCount: number
  sizeBytes: number
  exceedsMaxSize: boolean
}

interface FlaggedBundleExportGroupResponse {
  groupType: 'mailbox' | 'archive'
  label: string
  itemCount: number
  failedCount: number
  artifactCount: number
  artifacts: FlaggedBundleExportArtifactResponse[]
}

interface FlaggedBundleExportProgressResponse {
  stage: 'collecting' | 'mailbox' | 'archive' | 'finalizing' | 'succeeded' | 'failed'
  totalItems: number
  processedItems: number
  failedItems: number
  percent: number
  currentGroup: 'mailbox' | 'archive' | null
  currentLabel: string
}

interface FlaggedBundleExportJobResponse {
  exportId: string
  ownerUsername: string
  workspaceKey: string
  generatedAt: string
  startedAt: string
  completedAt: string | null
  updatedAt: string
  status: 'running' | 'succeeded' | 'failed'
  scope: FlaggedBundleExportScope
  maxSizeBytes: number
  progress: FlaggedBundleExportProgressResponse
  error: string | null
  groups: FlaggedBundleExportGroupResponse[]
}

interface FlaggedBundleExportHistoryResponse {
  scope: FlaggedBundleExportScope
  workspaceKey: string
  jobs: FlaggedBundleExportJobResponse[]
}

interface FlaggedBundleExportDeleteResponse {
  deleted: boolean
  exportId: string
}

interface FlaggedBundleExportArtifactRecord {
  artifactId: string
  fileName: string
  filePath: string
  partNumber: number
  partCount: number
  itemCount: number
  sizeBytes: number
  exceedsMaxSize: boolean
}

interface FlaggedBundleExportGroupRecord {
  groupType: 'mailbox' | 'archive'
  label: string
  itemCount: number
  failedCount: number
  artifactCount: number
  artifacts: Array<Omit<FlaggedBundleExportArtifactRecord, 'filePath'> & { downloadUrl: string }>
}

interface FlaggedBundleExportManifest {
  exportId: string
  generatedAt: string
  expiresAt: string
  scope: {
    scope: 'all' | 'search' | 'pst'
    scopePath: string
    scopeLabel: string
    sessionId: string
    sessionFileName: string
  }
  maxSizeBytes: number
  groups: FlaggedBundleExportGroupRecord[]
}

interface FlaggedBundleExportJob {
  exportId: string
  generatedAt: string
  expiresAt: string
  rootDir: string
  manifest: FlaggedBundleExportManifest
  artifactsById: Map<string, FlaggedBundleExportArtifactRecord>
}

interface BundleMailboxDescriptor {
  mailboxKey: string
  fileName: string
  scopePath: string
  scopeLabel: string
  session?: ViewerSessionIndex
}

interface ArchiveBundleDescriptor {
  bundlePath: string
  fileName: string
  scopePath: string
  scopeLabel: string
}

interface FlaggedBundleManifestItem {
  sourcePstPath: string
  mailboxName: string
  mailboxKey: string
  sourceType: 'mailbox' | 'archive'
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
  outputType: 'eml' | 'ics' | 'raw'
  status: 'exported' | 'error'
  error?: string
  archivePath?: string
  archiveEntryPath?: string
  archiveEntryName?: string
  contentType?: string
  downloadFilename?: string
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
const DEFAULT_AUTH_PASSWORD_CHANGE_CHALLENGE_COOKIE_SUFFIX = '-password-change-challenge'
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

interface LoadedReviewableItem {
  sourceType: SearchIndexDocument['sourceType']
  mailboxKey: string
  scopePath: string
  scopeLabel: string
  fileName: string
  folderId: string
  folderPath: string
  messageId: string
  descriptorId: string
  messageClass: string
  kind: ReviewRecord['kind']
  isMailLike: boolean
  subject: string
  senderName: string
  senderEmailAddress: string
  displayTo: string
  displayCC: string
  displayBCC: string
  resolvedDisplayTo: string
  resolvedDisplayCC: string
  resolvedDisplayBCC: string
  sortDate: string
  review: ReviewState
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
    pathname === API_ROUTES.authEntraStart ||
    pathname === API_ROUTES.authEntraCallback ||
    pathname === API_ROUTES.authMe ||
    pathname === API_ROUTES.authLogout ||
    pathname === API_ROUTES.authPasswordResetRequest ||
    pathname.startsWith(`${API_ROUTES.authPasswordResetLookup.split('/:token')[0]}`) ||
    pathname === API_ROUTES.authPasswordChangeConfirm ||
    pathname.startsWith(`${API_ROUTES.authInviteLookup.split('/:token')[0]}`) ||
    pathname.startsWith(`${API_ROUTES.authInviteAccept.split('/:token')[0]}`) ||
    pathname === API_ROUTES.authMfaChallenge
  )
}

function isAuthApiPath(pathname: string): boolean {
  return (
    pathname === API_ROUTES.authLogin ||
    pathname === API_ROUTES.authEntraStart ||
    pathname === API_ROUTES.authEntraCallback ||
    pathname === API_ROUTES.authMe ||
    pathname === API_ROUTES.authLogout ||
    pathname === API_ROUTES.authPasswordResetRequest ||
    pathname.startsWith(`${API_ROUTES.authPasswordResetLookup.split('/:token')[0]}`) ||
    pathname === API_ROUTES.authPasswordChangeConfirm ||
    pathname === API_ROUTES.authMfaChallenge
  )
}

function isProtectedApiPath(pathname: string): boolean {
  return pathname.startsWith('/api/') && !isPublicApiPath(pathname) && !isAuthApiPath(pathname)
}

interface EntraLoginStateRecord {
  state: string
  nonce: string
  codeVerifier: string
  returnTo: string
  expiresAt: number
}

interface EntraTokenResponse {
  token_type?: string
  scope?: string
  expires_in?: number
  ext_expires_in?: number
  access_token?: string
  refresh_token?: string
  id_token?: string
  error?: string
  error_description?: string
}

interface EntraIdTokenClaims {
  aud?: string | string[]
  iss?: string
  exp?: number
  nbf?: number
  iat?: number
  nonce?: string
  email?: string
  preferred_username?: string
  upn?: string
  oid?: string
  tid?: string
  name?: string
}

function base64UrlEncode(buffer: Buffer): string {
  return buffer
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '')
}

function base64UrlDecode(value: string): Buffer {
  const normalized = String(value || '')
    .replace(/-/g, '+')
    .replace(/_/g, '/')
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4)
  return Buffer.from(padded, 'base64')
}

function safeJsonParse<T>(value: string): T | null {
  try {
    return JSON.parse(value) as T
  } catch {
    return null
  }
}

function decodeJwt(token: string): { header: Record<string, unknown>; payload: EntraIdTokenClaims; signingInput: string; signature: Buffer } {
  const parts = String(token || '').split('.')
  if (parts.length !== 3) {
    throw createAppError(400, 'Invalid identity token')
  }

  const header = safeJsonParse<Record<string, unknown>>(base64UrlDecode(parts[0]).toString('utf8'))
  const payload = safeJsonParse<EntraIdTokenClaims>(base64UrlDecode(parts[1]).toString('utf8'))
  if (!header || !payload) {
    throw createAppError(400, 'Invalid identity token')
  }

  return {
    header,
    payload,
    signingInput: `${parts[0]}.${parts[1]}`,
    signature: base64UrlDecode(parts[2])
  }
}

function normalizeClaimValue(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

function sanitizeReturnTo(value: unknown, fallback = '/'): string {
  const text = normalizeText(value)
  if (!text) {
    return fallback
  }

  if (text.startsWith('//')) {
    return fallback
  }

  if (text.startsWith('/')) {
    return text
  }

  try {
    const parsed = new URL(text)
    if (parsed.origin && parsed.pathname) {
      return `${parsed.pathname}${parsed.search}${parsed.hash}` || fallback
    }
  } catch {
    return fallback
  }

  return fallback
}

function buildPkceCodeChallenge(codeVerifier: string): string {
  return base64UrlEncode(createHash('sha256').update(codeVerifier, 'utf8').digest())
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
  value: unknown,
  fallback: number
): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return fallback
  }
  const parsed = Number.parseInt(String(raw), 10)
  if (Number.isNaN(parsed) || parsed < 1) {
    return fallback
  }
  return parsed
}

function parseBoolean(value: unknown, fallback = false): boolean {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined || raw === null) {
    return fallback
  }
  if (typeof raw === 'boolean') {
    return raw
  }
  return !['0', 'false', 'no', 'off'].includes(String(raw).trim().toLowerCase())
}

function parseSort(value: unknown): string {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return 'date-desc'
  }
  const normalized = String(raw).trim().toLowerCase()
  if (normalized === 'order' || normalized === 'folder-order') {
    return 'order'
  }
  return 'date-desc'
}

function parseSearchMode(value: unknown): 'and' | 'or' {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return 'and'
  }
  return String(raw).trim().toLowerCase() === 'or' ? 'or' : 'and'
}

function parseQueryText(value: unknown): string {
  const raw = Array.isArray(value) ? value[0] : value
  return normalizeText(raw)
}

function parseSearchModeFromQuery(
  query: string,
  fallbackMode: unknown
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

function parseZeroBasedInt(value: unknown): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined) {
    return -1
  }
  const parsed = Number.parseInt(String(raw), 10)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : -1
}

function parseReviewFilters(value: Record<string, unknown>): ListFolderOptions {
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

function normalizeAssignedCasePath(value: unknown): string {
  const normalized = normalizeScopePath(value)
  if (!normalized) {
    return ''
  }

  return normalized.split('/').filter(Boolean)[0] || ''
}

function normalizeAssignedCasePaths(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return []
  }

  return [...new Set(value.map((entry) => normalizeAssignedCasePath(entry)).filter(Boolean))]
}

function isScopePathAllowed(scopePath: string, allowedCasePaths: string[], allowAll = false): boolean {
  if (allowAll) {
    return true
  }
  if (!allowedCasePaths.length) {
    return false
  }

  const normalizedScopePath = normalizeScopePath(scopePath)
  if (!normalizedScopePath) {
    return false
  }

  return allowedCasePaths.some(
    (allowedCasePath) =>
      normalizedScopePath === allowedCasePath || normalizedScopePath.startsWith(`${allowedCasePath}/`)
  )
}

function getAccessibleCasePaths(user: AuthUserListItem | null): string[] {
  return normalizeAssignedCasePaths(user?.assignedCasePaths || [])
}

function parseSearchScope(value: unknown): 'all' | 'search' | 'pst' {
  const normalized = normalizeText(value).toLowerCase()
  if (normalized === 'all' || normalized === 'search' || normalized === 'pst') {
    return normalized
  }
  return 'pst'
}

function parseSearchSourceType(value: unknown): 'mailbox' | 'teams' | 'sharepoint' | 'all' {
  const normalized = normalizeText(value).toLowerCase()
  if (normalized === 'teams' || normalized === 'sharepoint' || normalized === 'all') {
    return normalized
  }
  return 'mailbox'
}

function parseRefreshSource(value: unknown): SearchIndexRefreshSource {
  return normalizeText(value).toLowerCase() === 'items' ? 'items' : 'mailboxes'
}

function uniqueTextValues(values: string[]): string[] {
  return [...new Set(values.map((value) => normalizeText(value)).filter(Boolean))]
}

function collectCatalogMailboxKeys(
  rootPath: string,
  catalog: ReturnType<typeof listPstMailboxFiles>,
  resolveFilePath: (rootPath: string, scopePath: string, fileName: string) => string
): string[] {
  return catalog.scopes.flatMap((scope) =>
    scope.files.map((file) => resolveFilePath(rootPath, scope.scopePath, file.fileName))
  )
}

function resolveArchiveBundlePath(rootPath: string, scopePath: string, fileName: string): string {
  return path.resolve(rootPath, scopePath || '', fileName)
}

function buildArchiveMailboxKeys(rootPath: string, catalog: ReturnType<typeof listArchiveBundleFiles>): string[] {
  return collectCatalogMailboxKeys(rootPath, catalog, resolveArchiveBundlePath)
}

interface FingerprintScopeEntry {
  scopePath: string
  scopeLabel: string
  fileCount: number
  files: SearchIndexFileFingerprint[]
}

interface FingerprintCatalogSelection {
  scopes: FingerprintScopeEntry[]
  scopePath: string
  scopeLabel: string
  files: SearchIndexFileFingerprint[]
}

function getScopeLabel(scopePath: string): string {
  return scopePath ? scopePath.split('/').join(' / ') : 'PST root'
}

function buildFingerprintScopeEntries(fingerprints: SearchIndexFileFingerprint[]): FingerprintScopeEntry[] {
  const scopes = new Map<string, FingerprintScopeEntry>()

  for (const fingerprint of fingerprints) {
    const scopePath = normalizeScopePath(fingerprint.scopePath)
    const existingScope = scopes.get(scopePath)
    const normalizedScopeLabel = normalizeText(fingerprint.scopeLabel) || getScopeLabel(scopePath)
    if (!existingScope) {
      scopes.set(scopePath, {
        scopePath,
        scopeLabel: normalizedScopeLabel,
        fileCount: 1,
        files: [
          {
            ...fingerprint,
            scopePath,
            scopeLabel: normalizedScopeLabel
          }
        ]
      })
      continue
    }

    existingScope.fileCount += 1
    existingScope.files.push({
      ...fingerprint,
      scopePath,
      scopeLabel: existingScope.scopeLabel || normalizedScopeLabel
    })
  }

  return [...scopes.values()]
    .map((scope) => ({
      ...scope,
      files: scope.files.sort((left, right) =>
        left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
      )
    }))
    .sort((left, right) => {
      if (left.scopePath === right.scopePath) {
        return 0
      }
      if (left.scopePath === '') {
        return -1
      }
      if (right.scopePath === '') {
        return 1
      }
      return left.scopeLabel.localeCompare(right.scopeLabel, undefined, { sensitivity: 'base' })
    })
}

function chooseFingerprintScopeEntry(
  scopes: FingerprintScopeEntry[],
  requestedScopePath: unknown
): FingerprintScopeEntry | null {
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  if (normalizedRequestedScopePath) {
    const requestedScope = scopes.find((scope) => scope.scopePath === normalizedRequestedScopePath)
    if (requestedScope) {
      return requestedScope
    }
  }

  const rootScope = scopes.find((scope) => scope.scopePath === '')
  return rootScope || scopes[0] || null
}

function collectFingerprintMailboxKeys(scopes: FingerprintScopeEntry[]): string[] {
  return uniqueTextValues(scopes.flatMap((scope) => scope.files.map((file) => file.mailboxKey)))
}

interface FingerprintMailboxSelection {
  scopePath: string
  scopeLabel: string
  mailboxKeys: string[]
}

function selectFingerprintMailboxKeys(
  requestedCasePath: string,
  requestedScopePath: string,
  allowedCasePaths: string[],
  fingerprints: SearchIndexFileFingerprint[],
  allowAll = false
): FingerprintMailboxSelection {
  const normalizedRequestedCasePath = normalizeScopePath(requestedCasePath)
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  const accessibleScopes = buildFingerprintScopeEntries(fingerprints).filter((scope) =>
    isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAll)
  )

  let selectedScopes = accessibleScopes
  let scopePath = ''
  let scopeLabel = 'All cases/searches'

  if (normalizedRequestedScopePath) {
    selectedScopes = accessibleScopes.filter((scope) => scope.scopePath === normalizedRequestedScopePath)
    scopePath = normalizedRequestedScopePath
    scopeLabel = selectedScopes[0]?.scopeLabel || getScopeLabel(normalizedRequestedScopePath)
  } else if (normalizedRequestedCasePath) {
    selectedScopes = accessibleScopes.filter(
      (scope) =>
        scope.scopePath === normalizedRequestedCasePath ||
        scope.scopePath.startsWith(`${normalizedRequestedCasePath}/`)
    )
    scopeLabel = getScopeLabel(normalizedRequestedCasePath)
  }

  return {
    scopePath,
    scopeLabel,
    mailboxKeys: uniqueTextValues(
      selectedScopes.flatMap((scope) => scope.files.map((file) => file.mailboxKey))
    )
  }
}

function mergeFingerprintMailboxSelections(
  ...selections: FingerprintMailboxSelection[]
): FingerprintMailboxSelection {
  return {
    scopePath: selections.find((selection) => selection.scopePath)?.scopePath || '',
    scopeLabel:
      selections.find((selection) => selection.scopePath)?.scopeLabel ||
      selections.find((selection) => selection.scopeLabel && selection.scopeLabel !== 'All cases/searches')
        ?.scopeLabel ||
      'All cases/searches',
    mailboxKeys: uniqueTextValues(selections.flatMap((selection) => selection.mailboxKeys))
  }
}

function resolveAccessibleFingerprintCatalogSelection(
  requestedScopePath: string,
  allowedCasePaths: string[],
  fingerprints: SearchIndexFileFingerprint[],
  allowAll = false
): FingerprintCatalogSelection | null {
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  if (normalizedRequestedScopePath && !isScopePathAllowed(normalizedRequestedScopePath, allowedCasePaths, allowAll)) {
    throw createAppError(403, 'Case access required')
  }

  if (!allowAll && !allowedCasePaths.length) {
    return {
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: []
    }
  }

  if (!fingerprints.length) {
    return null
  }

  const scopes = buildFingerprintScopeEntries(fingerprints)
  const accessibleScopes = scopes.filter((scope) => isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAll))
  if (!accessibleScopes.length) {
    return {
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: []
    }
  }

  if (normalizedRequestedScopePath) {
    const requestedScope = accessibleScopes.find((scope) => scope.scopePath === normalizedRequestedScopePath)
    if (!requestedScope) {
      return null
    }
    return {
      scopes: accessibleScopes,
      scopePath: requestedScope.scopePath,
      scopeLabel: requestedScope.scopeLabel || getScopeLabel(requestedScope.scopePath),
      files: requestedScope.files
    }
  }

  const effectiveScopePath = normalizedRequestedScopePath || (allowAll ? '' : allowedCasePaths[0] || '')
  const selectedScopeCandidate = chooseFingerprintScopeEntry(scopes, effectiveScopePath)
  const selectedScope =
    accessibleScopes.find((scope) => scope.scopePath === selectedScopeCandidate?.scopePath) ||
    accessibleScopes[0] ||
    null

  if (!selectedScope) {
    return {
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: []
    }
  }

  return {
    scopes: accessibleScopes,
    scopePath: selectedScope.scopePath,
    scopeLabel: selectedScope.scopeLabel || getScopeLabel(selectedScope.scopePath),
    files: selectedScope.files
  }
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

function resolveAccessibleCatalogSelection(
  rootPath: string,
  requestedScopePath: string,
  allowedCasePaths: string[],
  loader: (rootPath: string, scopePath?: string) => ReturnType<typeof listPstMailboxFiles> = listPstMailboxFiles,
  allowAll = false
): ReturnType<typeof listPstMailboxFiles> {
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  if (normalizedRequestedScopePath && !isScopePathAllowed(normalizedRequestedScopePath, allowedCasePaths, allowAll)) {
    throw createAppError(403, 'Case access required')
  }

  if (!allowAll && !allowedCasePaths.length) {
    return {
      rootPath,
      rootExists: fs.existsSync(rootPath),
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: [],
      message: 'No cases assigned.'
    }
  }

  const effectiveScopePath = normalizedRequestedScopePath || (allowAll ? '' : allowedCasePaths[0] || '')
  const catalog = loader(rootPath, effectiveScopePath)
  if (allowAll) {
    return catalog
  }

  const scopes = catalog.scopes.filter((scope) => isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAll))
  if (!scopes.length) {
    return {
      ...catalog,
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: [],
      message: catalog.message
    }
  }

  const selectedScope =
    scopes.find((scope) => scope.scopePath === catalog.scopePath) || scopes[0] || null

  return {
    ...catalog,
    scopes,
    scopePath: selectedScope?.scopePath || '',
    scopeLabel: selectedScope?.scopeLabel || '',
    files: selectedScope?.files || []
  }
}

function resolveAccessibleArchiveCatalogSelection(
  rootPath: string,
  requestedScopePath: string,
  allowedCasePaths: string[],
  allowAll = false
): ReturnType<typeof listArchiveBundleFiles> {
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  if (normalizedRequestedScopePath && !isScopePathAllowed(normalizedRequestedScopePath, allowedCasePaths, allowAll)) {
    throw createAppError(403, 'Case access required')
  }

  if (!allowAll && !allowedCasePaths.length) {
    return {
      rootPath,
      rootExists: fs.existsSync(rootPath),
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: [],
      message: 'No cases assigned.'
    }
  }

  const catalog = listArchiveBundleFiles(rootPath, '')
  const scopes = catalog.scopes.filter((scope) => isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAll))
  if (!scopes.length) {
    return {
      ...catalog,
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: [],
      message: catalog.message
    }
  }

  if (!normalizedRequestedScopePath) {
    const selectedScope = scopes.find((scope) => scope.scopePath === '') || scopes[0] || null
    return {
      ...catalog,
      scopes,
      scopePath: selectedScope?.scopePath || '',
      scopeLabel: selectedScope?.scopeLabel || '',
      files: selectedScope?.files || [],
      message: catalog.message
    }
  }

  const selectedScope = scopes.find((scope) => scope.scopePath === normalizedRequestedScopePath)
  if (!selectedScope) {
    return {
      ...catalog,
      scopes,
      scopePath: normalizedRequestedScopePath,
      scopeLabel: normalizedRequestedScopePath.split('/').join(' / '),
      files: [],
      message: `No Items*.zip bundles were found in ${normalizedRequestedScopePath.split('/').join(' / ')}.`
    }
  }

  return {
    ...catalog,
    scopes,
    scopePath: selectedScope.scopePath,
    scopeLabel: selectedScope.scopeLabel,
    files: selectedScope.files,
    message: catalog.message
  }
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

function buildArchiveBundleEntryPath(
  scopePath: string,
  bundleFileName: string,
  archiveEntryPath: string,
  downloadFilename: string,
  messageId: string
): string {
  const scopeSegment = buildBundleScopeSegment(scopePath)
  const bundleSegment = sanitizeBundleSegment(bundleFileName, 'bundle')
  const entrySegments = normalizeScopePath(archiveEntryPath)
    ? normalizeScopePath(archiveEntryPath)
        .split('/')
        .map((segment) => sanitizeBundleSegment(segment, 'entry'))
        .filter(Boolean)
    : []
  const fileSegment = sanitizeBundleSegment(downloadFilename || messageId, 'item')
  if (entrySegments.length) {
    entrySegments.pop()
  }
  return ['archive', scopeSegment, bundleSegment, ...entrySegments, fileSegment].filter(Boolean).join('/')
}

function buildBundleMailboxes(
  rootPath: string,
  scope: 'all' | 'search' | 'pst',
  scopePath: string,
  session: SessionRecord | null,
  allowedCasePaths: string[],
  allowAll = false
): BundleMailboxDescriptor[] {
  if (scope === 'pst') {
    if (!session) {
      throw createAppError(400, 'Session id is required for selected PST exports')
    }
    if (!isScopePathAllowed(session.scopePath, allowedCasePaths, allowAll)) {
      throw createAppError(403, 'Case access required')
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
      ? resolveAccessibleCatalogSelection(rootPath, scopePath, allowedCasePaths, listPstMailboxFiles, allowAll)
      : resolveAccessibleCatalogSelection(rootPath, '', allowedCasePaths, listPstMailboxFiles, allowAll)
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

function buildBundleArchiveDescriptors(
  rootPath: string,
  scope: 'all' | 'search' | 'pst',
  scopePath: string,
  allowedCasePaths: string[],
  allowAll = false
): ArchiveBundleDescriptor[] {
  if (scope === 'pst') {
    return []
  }

  const catalog =
    scope === 'search'
      ? resolveAccessibleArchiveCatalogSelection(rootPath, scopePath, allowedCasePaths, allowAll)
      : resolveAccessibleCatalogSelection(rootPath, '', allowedCasePaths, listArchiveBundleFiles, allowAll)
  const scopes = scope === 'search' ? [catalog] : catalog.scopes

  const bundles: ArchiveBundleDescriptor[] = []
  for (const catalogScope of scopes) {
    for (const file of catalogScope.files) {
      bundles.push({
        bundlePath: resolveArchiveBundlePath(rootPath, catalogScope.scopePath, file.fileName),
        fileName: file.fileName,
        scopePath: catalogScope.scopePath,
        scopeLabel: catalogScope.scopeLabel
      })
    }
  }
  return bundles
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

interface FlaggedBundleZipEntry {
  entryName: string
  content: Buffer
  mtime: Date
}

function buildFlaggedBundleGroupLabel(groupType: 'mailbox' | 'archive'): string {
  return groupType === 'mailbox' ? 'Mailbox' : 'Teams / SharePoint'
}

function buildFlaggedBundleArtifactId(groupType: 'mailbox' | 'archive', partNumber: number): string {
  return `${groupType}-${partNumber}`
}

function buildFlaggedBundleArtifactFileName(groupType: 'mailbox' | 'archive', partNumber: number): string {
  const normalizedGroup = groupType === 'mailbox' ? 'mailbox' : 'teams-sharepoint'
  return `flagged-${normalizedGroup}-part-${partNumber}.zip`
}

function toFlaggedBundleBuffer(content: Buffer | string): Buffer {
  return Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8')
}

async function writeFlaggedBundleZipFile(
  filePath: string,
  entries: FlaggedBundleZipEntry[]
): Promise<number> {
  const stream = fs.createWriteStream(filePath)
  const writer = createZipStreamWriter(stream)
  try {
    for (const entry of entries) {
      await writer.addFile(entry.entryName, entry.content, { mtime: entry.mtime })
    }
    await writer.finalize()
    stream.end()
    await once(stream, 'finish')
    const stats = await fs.promises.stat(filePath)
    return stats.size
  } catch (error) {
    stream.destroy()
    try {
      await fs.promises.rm(filePath, { force: true })
    } catch {
      // ignore cleanup failures
    }
    throw error
  }
}

function buildReviewContextFromSearchIndexDocument(
  item: SearchIndexDocument,
  reviewerUsername: string
): Parameters<ReviewStore['upsertReview']>[0] {
  return {
    mailboxKey: item.mailboxKey,
    reviewerUsername,
    fileName: item.fileName,
    messageId: item.messageId,
    descriptorId: item.descriptorId,
    folderId: item.folderId,
    folderPath: item.folderPath,
    messageClass: item.messageClass,
    kind: item.kind,
    isMailLike: item.isMailLike,
    subject: item.subject,
    senderName: item.senderName,
    senderEmailAddress: item.senderEmailAddress,
    displayTo: item.displayTo,
    displayCC: item.displayCC,
    displayBCC: item.displayBCC,
    resolvedDisplayTo: item.resolvedDisplayTo,
    resolvedDisplayCC: item.resolvedDisplayCC,
    resolvedDisplayBCC: item.resolvedDisplayBCC
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

function createMailboxAttachmentBaseUrl(sessionId: string, messageId: string): string {
  return `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(
    messageId
  )}/attachments/`
}

function createItemAttachmentBaseUrl(itemId: string): string {
  return `/api/items/${encodeURIComponent(itemId)}/attachments/`
}

function applyAttachmentDownloadUrls(detail: MessageDetail, attachmentBaseUrl: string): MessageDetail {
  return {
    ...detail,
    attachments: (detail.attachments || []).map((attachment) => ({
      ...attachment,
      downloadUrl:
        attachment.isDownloadable && attachmentBaseUrl
          ? `${attachmentBaseUrl}${attachment.index}`
          : '',
      embeddedMessage: attachment.embeddedMessage
        ? applyAttachmentDownloadUrls(attachment.embeddedMessage, attachmentBaseUrl)
        : null
    }))
  }
}

function applyMailboxAttachmentDownloadUrls(
  detail: MessageDetail,
  sessionId: string
): MessageDetail {
  const detailId = normalizeText(detail.id)
  const attachmentBaseUrl = detailId ? createMailboxAttachmentBaseUrl(sessionId, detailId) : ''
  return applyAttachmentDownloadUrls(detail, attachmentBaseUrl)
}

function applyItemAttachmentDownloadUrls(detail: MessageDetail, itemId: string): MessageDetail {
  const detailId = normalizeText(itemId || detail.id || '')
  const attachmentBaseUrl = detailId ? createItemAttachmentBaseUrl(detailId) : ''
  return applyAttachmentDownloadUrls(detail, attachmentBaseUrl)
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

  function buildMailboxSearchPreviewDetail(item: SearchIndexDocument): MessageDetail | undefined {
    if (item.sourceType !== 'mailbox' || !item.mailboxDetail) {
      return undefined
    }

    const itemId = buildMailboxSearchDocumentId(item.mailboxKey, item.messageId)
    return applyItemAttachmentDownloadUrls(
      {
        ...buildReviewedDetail(item.mailboxDetail, item.review),
        id: itemId
      },
      itemId
    )
  }

  async function resolveMailboxMessageDetail(
    session: SessionRecord,
    messageId: string
  ): Promise<MessageDetail> {
  const summary = session.index.messages.get(messageId) || null
  if (!summary) {
    throw createAppError(404, `Unknown message: ${messageId}`)
  }

  return applyMailboxAttachmentDownloadUrls(
    buildMessageDetailFromSession(session.index, messageId, 1),
    session.id
  )
}

async function resolveMailboxItemDetail(
  item: SearchIndexDocument,
  reviewStore: ReviewStore,
  reviewerUsername: string
): Promise<MessageDetail> {
  if (!item.mailboxDetail) {
    throw createAppError(500, 'Mailbox preview data unavailable; rebuild the search index')
  }

  const review = await reviewStore.getReview(item.mailboxKey, item.messageId, reviewerUsername)
  const itemId = buildMailboxSearchDocumentId(item.mailboxKey, item.messageId)
  return applyItemAttachmentDownloadUrls(
    {
      ...buildReviewedDetail(item.mailboxDetail, review),
      id: itemId
    },
    itemId
  )
}

async function resolveSearchItemDetail(
  item: SearchIndexDocument,
  reviewStore: ReviewStore,
  reviewerUsername: string
): Promise<MessageDetail> {
  if (item.sourceType === 'mailbox') {
    return resolveMailboxItemDetail(item, reviewStore, reviewerUsername)
  }

  const review = await reviewStore.getReview(item.mailboxKey, item.messageId, reviewerUsername)
  return buildReviewedDetail(
    ({
      id: item.id || item.messageId,
      subject: item.subject,
      senderName: item.senderName,
      senderEmailAddress: item.senderEmailAddress,
      displayTo: item.displayTo,
      displayCC: item.displayCC,
      displayBCC: item.displayBCC,
      resolvedDisplayTo: item.resolvedDisplayTo,
      resolvedDisplayCC: item.resolvedDisplayCC,
      resolvedDisplayBCC: item.resolvedDisplayBCC,
      clientSubmitTime: item.clientSubmitTime,
      creationTime: item.creationTime,
      modificationTime: item.modificationTime,
      messageDeliveryTime: item.messageDeliveryTime,
      sortDate: item.sortDate,
      bodyHtml: item.previewHtml || '',
      bodyText: item.previewText || '',
      bodyPrefix: item.previewText || item.previewHtml ? 'Preview from archive item' : '',
      parseError: '',
      attachments: [],
      folderId: item.folderId,
      folderPath: item.folderPath,
      mailboxName: item.mailboxName,
      archivePath: item.archivePath,
      archiveEntryPath: item.archiveEntryPath,
      archiveEntryChain: item.archiveEntryChain,
      archiveEntryName: item.archiveEntryName,
      contentType: item.contentType,
      downloadFilename: item.downloadFilename,
      previewKind: item.previewKind,
      previewText: item.previewText,
      previewHtml: item.previewHtml,
      previewUrl: item.archivePath ? `/api/items/${encodeURIComponent(item.id || item.messageId)}/preview` : '',
      downloadUrl: item.archivePath ? `/api/items/${encodeURIComponent(item.id || item.messageId)}/content` : ''
    } as unknown as MessageDetail),
    review
  )
}

function buildMailboxDetailCacheKey(messageId: string, reviewerUsername: string): string {
  return `${normalizeText(reviewerUsername) || 'anonymous'}::${normalizeText(messageId)}`
}

function invalidateMailboxDetailCache(
  session: SessionRecord,
  messageId?: string,
  reviewerUsername?: string
): void {
  if (!messageId && !reviewerUsername) {
    session.messageDetailCache.clear()
    return
  }

  if (messageId && reviewerUsername) {
    session.messageDetailCache.delete(buildMailboxDetailCacheKey(messageId, reviewerUsername))
    return
  }

  const normalizedMessageId = normalizeText(messageId || '')
  if (!normalizedMessageId) {
    return
  }

  for (const cacheKey of session.messageDetailCache.keys()) {
    if (cacheKey.endsWith(`::${normalizedMessageId}`)) {
      session.messageDetailCache.delete(cacheKey)
    }
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
  const cacheKey = buildMailboxDetailCacheKey(messageId, reviewerUsername)
  const cached = session.messageDetailCache.get(cacheKey)
  if (cached) {
    return cached
  }

  const detailPromise = (async () => {
    const [review, detail] = await Promise.all([
      reviewStore.getReview(session.filePath, messageId, reviewerUsername),
      resolveMailboxMessageDetail(session, messageId)
    ])
    return buildReviewedDetail(detail, review)
  })()

  session.messageDetailCache.set(cacheKey, detailPromise)

  try {
    return await detailPromise
  } catch (error) {
    if (session.messageDetailCache.get(cacheKey) === detailPromise) {
      session.messageDetailCache.delete(cacheKey)
    }
    throw error
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
  const reusableMailboxSessions = new Map<string, string>()
  const openingMailboxSessions = new Map<string, Promise<SessionResponse>>()
  const publicDir = options.publicDir
  const pstRootDir = options.pstRootDir || getDefaultPstRootDirectory()
  const reviewStore = options.reviewStore
  const searchIndexStore = options.searchIndexStore
  const flaggedBundleStore = options.flaggedBundleStore || createMemoryFlaggedBundleStore()
  const authConfig = normalizeAuthConfig(options.auth)
  const authUserStore = options.authUserStore || createMemoryAuthUserStore(authConfig.seedUsers)
  const appSettingsStore = options.appSettingsStore || createMemoryAppSettingsStore()
  const auditLogStore = options.auditLogDir ? createFileAuditLogStore(options.auditLogDir) : null
  const smtpTransportFactory: SmtpTransportFactory =
    options.smtpTransportFactory || createDefaultSmtpTransport
  const authSessions = new Map<string, AuthSessionRecord>()
  const authMfaChallenges = new Map<string, AuthMfaChallengeRecord>()
  const authPasswordChangeChallenges = new Map<string, AuthPasswordChangeChallengeRecord>()
  const entraLoginStates = new Map<string, EntraLoginStateRecord>()

  function cleanupExpiredEntraLoginStates(): void {
    const now = Date.now()
    for (const [state, record] of entraLoginStates.entries()) {
      if (record.expiresAt <= now) {
        entraLoginStates.delete(state)
      }
    }
  }

  function createEntraLoginState(returnTo: string): EntraLoginStateRecord {
    const state = randomBytes(24).toString('hex')
    const nonce = randomBytes(24).toString('hex')
    const codeVerifier = randomBytes(32).toString('hex')
    const record: EntraLoginStateRecord = {
      state,
      nonce,
      codeVerifier,
      returnTo: sanitizeReturnTo(returnTo),
      expiresAt: Date.now() + 10 * 60 * 1000
    }
    entraLoginStates.set(state, record)
    return record
  }

  function consumeEntraLoginState(state: string): EntraLoginStateRecord | null {
    cleanupExpiredEntraLoginStates()
    const normalizedState = normalizeText(state)
    if (!normalizedState) {
      return null
    }

    const record = entraLoginStates.get(normalizedState) || null
    if (!record) {
      return null
    }

    entraLoginStates.delete(normalizedState)
    if (record.expiresAt <= Date.now()) {
      return null
    }

    return record
  }

  function getEntraAuthorityBase(settings: EntraSettingsRecord): string {
    const tenantId = encodeURIComponent(normalizeText(settings.tenantId) || 'common')
    return `https://login.microsoftonline.com/${tenantId}`
  }

  function buildEntraAuthorizeUrl(
    settings: EntraSettingsRecord,
    redirectUri: string,
    state: string,
    nonce: string,
    codeVerifier: string
  ): string {
    const authorizeUrl = new URL(`${getEntraAuthorityBase(settings)}/oauth2/v2.0/authorize`)
    authorizeUrl.searchParams.set('client_id', normalizeText(settings.clientId))
    authorizeUrl.searchParams.set('response_type', 'code')
    authorizeUrl.searchParams.set('redirect_uri', redirectUri)
    authorizeUrl.searchParams.set('response_mode', 'query')
    authorizeUrl.searchParams.set('scope', 'openid profile email')
    authorizeUrl.searchParams.set('state', state)
    authorizeUrl.searchParams.set('nonce', nonce)
    authorizeUrl.searchParams.set('code_challenge', buildPkceCodeChallenge(codeVerifier))
    authorizeUrl.searchParams.set('code_challenge_method', 'S256')
    authorizeUrl.searchParams.set('prompt', 'select_account')
    return authorizeUrl.toString()
  }

  function buildEntraTokenUrl(settings: EntraSettingsRecord): string {
    return `${getEntraAuthorityBase(settings)}/oauth2/v2.0/token`
  }

  function buildEntraJwksUrl(settings: EntraSettingsRecord): string {
    return `${getEntraAuthorityBase(settings)}/discovery/v2.0/keys`
  }

  function buildEntraRedirectUri(req: express.Request): string {
    const origin =
      normalizeOrigin(authConfig.publicBaseUrl) || getRequestBaseOrigin(req) || canonicalRequestOrigin(req)
    return `${origin}${API_ROUTES.authEntraCallback}`
  }

  function buildFlaggedBundleExportDownloadUrl(exportId: string, artifactId: string): string {
    return `${API_ROUTES.flaggedBundleArtifact
      .replace(':exportId', encodeURIComponent(exportId))
      .replace(':artifactId', encodeURIComponent(artifactId))}`
  }

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

  function buildSessionResponse(session: SessionRecord): SessionResponse {
    return {
      sessionId: session.id,
      scopePath: session.scopePath,
      scopeLabel: session.scopeLabel,
      fileName: session.fileName,
      summary: buildSessionSummary(session.index),
      tree: buildFolderTree(session.index)
    }
  }

  function getLiveSessionForMailboxKey(mailboxKey: string): SessionRecord | null {
    const normalizedMailboxKey = normalizeText(mailboxKey)
    if (!normalizedMailboxKey) {
      return null
    }

    const cachedSessionId = reusableMailboxSessions.get(normalizedMailboxKey)
    if (cachedSessionId) {
      const cachedSession = sessions.get(cachedSessionId)
      if (cachedSession) {
        return cachedSession
      }
      reusableMailboxSessions.delete(normalizedMailboxKey)
    }

    for (const session of sessions.values()) {
      if (session.filePath === normalizedMailboxKey) {
        return session
      }
    }

    return null
  }

  function registerMailboxSession(
    sessionIndex: ViewerSessionIndex,
    mailboxKey: string,
    scopePath: string,
    fileName: string
  ): SessionRecord {
    const normalizedMailboxKey = normalizeText(mailboxKey)
    const sessionId = createSessionId()
    const scopeLabel = scopePath ? scopePath.split('/').join(' / ') : 'PST root'
    const record: SessionRecord = {
      id: sessionId,
      index: sessionIndex,
      filePath: sessionIndex.filePath,
      fileName: sessionIndex.fileName || fileName,
      scopePath,
      scopeLabel,
      mailboxKey: normalizedMailboxKey || sessionIndex.filePath,
      messageDetailCache: new Map<string, Promise<ReviewedMessageDetail>>()
    }
    sessions.set(sessionId, record)
    if (record.mailboxKey) {
      reusableMailboxSessions.set(record.mailboxKey, sessionId)
    }
    return record
  }

  function clearReusableMailboxSessions(mailboxKey?: string): void {
    if (!mailboxKey) {
      reusableMailboxSessions.clear()
      openingMailboxSessions.clear()
      return
    }

    const normalizedMailboxKey = normalizeText(mailboxKey)
    if (!normalizedMailboxKey) {
      return
    }

    reusableMailboxSessions.delete(normalizedMailboxKey)
    openingMailboxSessions.delete(normalizedMailboxKey)
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

  function buildLoadedItemsCsvFileName(): string {
    return 'loaded-items.csv'
  }

  function buildAllItemsCsvFileName(sourceType: string): string {
    const normalized = normalizeText(sourceType).toLowerCase()
    if (normalized === 'teams') {
      return 'all-items-teams.csv'
    }
    if (normalized === 'sharepoint') {
      return 'all-items-sharepoint.csv'
    }
    if (normalized === 'mailbox') {
      return 'all-items-mailbox.csv'
    }
    return 'all-items.csv'
  }

  function buildLoadedItemsCsvHeader(): string {
    return [
      'sourceType',
      'scopePath',
      'scopeLabel',
      'fileName',
      'mailboxKey',
      'folderPath',
      'messageId',
      'descriptorId',
      'kind',
      'subject',
      'senderName',
      'senderEmailAddress',
      'sortDate',
      'flagged',
      'tags',
      'reviewUpdatedAt'
    ]
      .map(csvCell)
      .join(',')
  }

  function buildLoadedItemsCsvRow(item: LoadedReviewableItem): string {
    return [
      item.sourceType,
      item.scopePath,
      item.scopeLabel,
      item.fileName,
      item.mailboxKey,
      item.folderPath,
      item.messageId,
      item.descriptorId,
      item.kind,
      item.subject,
      item.senderName,
      item.senderEmailAddress,
      item.sortDate,
      item.review.flagged,
      item.review.tags.join('; '),
      item.review.updatedAt
    ]
      .map(csvCell)
      .join(',')
  }

  function buildLoadedReviewableItemFromSearchDocument(
    item: SearchIndexDocument
  ): LoadedReviewableItem {
    return {
      sourceType: item.sourceType,
      mailboxKey: item.mailboxKey,
      scopePath: item.scopePath,
      scopeLabel: item.scopeLabel,
      fileName: item.fileName,
      folderId: item.folderId,
      folderPath: item.folderPath,
      messageId: item.messageId,
      descriptorId: item.descriptorId,
      messageClass: item.messageClass,
      kind: item.kind,
      isMailLike: item.isMailLike,
      subject: item.subject,
      senderName: item.senderName,
      senderEmailAddress: item.senderEmailAddress,
      displayTo: item.displayTo,
      displayCC: item.displayCC,
      displayBCC: item.displayBCC,
      resolvedDisplayTo: item.resolvedDisplayTo,
      resolvedDisplayCC: item.resolvedDisplayCC,
      resolvedDisplayBCC: item.resolvedDisplayBCC,
      sortDate: item.sortDate || '',
      review: normalizeReviewState(item.review)
    }
  }

  function buildSearchResultSummary(item: SearchIndexDocument) {
    const id =
      item.sourceType === 'mailbox'
        ? buildMailboxSearchDocumentId(item.mailboxKey, item.messageId)
        : normalizeText(item.id || item.messageId)
    const messageId = normalizeText(item.messageId || id)
    return {
      id,
      messageId,
      sourceType: item.sourceType,
      descriptorId: item.descriptorId,
      folderId: item.folderId,
      folderPath: item.folderPath,
      order: item.order,
      messageClass: item.messageClass,
      kind: item.kind,
      subject: item.subject,
      senderName: item.senderName,
      senderEmailAddress: item.senderEmailAddress,
      recipientText: item.recipientText,
      displayTo: item.displayTo,
      displayCC: item.displayCC,
      displayBCC: item.displayBCC,
      resolvedDisplayTo: item.resolvedDisplayTo,
      resolvedDisplayCC: item.resolvedDisplayCC,
      resolvedDisplayBCC: item.resolvedDisplayBCC,
      originalSubject: item.originalSubject,
      clientSubmitTime: item.clientSubmitTime,
      creationTime: item.creationTime,
      modificationTime: item.modificationTime,
      messageDeliveryTime: item.messageDeliveryTime,
      sortDate: item.sortDate,
      sortDateMs: item.sortDateMs,
      importance: item.importance,
      hasAttachments: item.hasAttachments,
      isRead: item.isRead,
      isMailLike: item.isMailLike,
      size: item.size,
      review: item.review,
      scopePath: item.scopePath,
      scopeLabel: item.scopeLabel,
      fileName: item.fileName,
      mailboxName: item.mailboxName,
      mailboxDetail: item.sourceType === 'mailbox' ? buildMailboxSearchPreviewDetail(item) : undefined,
      contentType: item.contentType,
      downloadFilename: item.downloadFilename
    }
  }

  function buildLoadedReviewableItemFromFolderSummary(
    session: SessionRecord,
    item: ReviewedMessageSummary
  ): LoadedReviewableItem {
    return {
      sourceType: 'mailbox',
      mailboxKey: session.filePath,
      scopePath: session.scopePath,
      scopeLabel: session.scopeLabel || getScopeLabel(session.scopePath),
      fileName: session.fileName,
      folderId: item.folderId,
      folderPath: item.folderPath,
      messageId: item.id,
      descriptorId: item.descriptorId,
      messageClass: item.messageClass,
      kind: item.kind,
      isMailLike: item.isMailLike,
      subject: item.subject,
      senderName: item.senderName,
      senderEmailAddress: item.senderEmailAddress,
      displayTo: item.displayTo,
      displayCC: item.displayCC,
      displayBCC: item.displayBCC,
      resolvedDisplayTo: item.resolvedDisplayTo,
      resolvedDisplayCC: item.resolvedDisplayCC,
      resolvedDisplayBCC: item.resolvedDisplayBCC,
      sortDate: item.sortDate || '',
      review: normalizeReviewState(item.review)
    }
  }

  function buildReviewPatchInputFromLoadedItem(
    item: LoadedReviewableItem,
    reviewerUsername: string
  ): Parameters<ReviewStore['upsertReview']>[0] {
    return {
      mailboxKey: item.mailboxKey,
      reviewerUsername: normalizeText(reviewerUsername) || 'anonymous',
      fileName: item.fileName,
      messageId: item.messageId,
      descriptorId: item.descriptorId,
      folderId: item.folderId,
      folderPath: item.folderPath,
      messageClass: item.messageClass,
      kind: item.kind,
      isMailLike: item.isMailLike,
      subject: item.subject,
      senderName: item.senderName,
      senderEmailAddress: item.senderEmailAddress,
      displayTo: item.displayTo,
      displayCC: item.displayCC,
      displayBCC: item.displayBCC,
      resolvedDisplayTo: item.resolvedDisplayTo,
      resolvedDisplayCC: item.resolvedDisplayCC,
      resolvedDisplayBCC: item.resolvedDisplayBCC
    }
  }

  function buildLoadedItemsCsv(items: LoadedReviewableItem[]): string {
    const lines = items.map((item) => buildLoadedItemsCsvRow(item))
    return [buildLoadedItemsCsvHeader(), ...lines].join('\r\n')
  }

  function dedupeTextValues(values: string[]): string[] {
    const seen = new Set<string>()
    const deduped: string[] = []
    for (const value of values) {
      const normalized = normalizeText(value)
      if (!normalized || seen.has(normalized)) {
        continue
      }
      seen.add(normalized)
      deduped.push(normalized)
    }
    return deduped
  }

  async function resolveCsvExportMailboxKeys(authSession: AuthSessionRecord | null): Promise<string[] | undefined> {
    const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
    if (allowAllCases) {
      return undefined
    }

    const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
    const allowedCasePaths = getAccessibleCasePaths(currentUser)
    const mailCatalog = listPstMailboxFiles(pstRootDir)
    const archiveCatalog = listArchiveBundleFiles(pstRootDir)
    const mailboxKeys = mailCatalog.scopes
      .filter((scope) => isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAllCases))
      .flatMap((scope) => scope.files.map((file) => resolvePstMailboxPath(pstRootDir, scope.scopePath, file.fileName)))
    const archiveMailboxKeys = archiveCatalog.scopes
      .filter((scope) => isScopePathAllowed(scope.scopePath, allowedCasePaths, allowAllCases))
      .flatMap((scope) =>
        scope.files.map((file) => resolveArchiveBundlePath(pstRootDir, scope.scopePath, file.fileName))
      )
    return dedupeTextValues([...mailboxKeys, ...archiveMailboxKeys])
  }

  async function streamLoadedItemsCsv(
    res: express.Response,
    authSession: AuthSessionRecord | null
  ): Promise<number> {
    const reviewerUsername = getReviewOwnerUsername(authSession)
    const allowedMailboxKeys = await resolveCsvExportMailboxKeys(authSession)
    const pageSize = 1000
    let page = 1
    let totalPages = 1
    let itemCount = 0

    res.status(200)
      .type('text/csv; charset=utf-8')
      .set('Content-Disposition', `attachment; filename="${buildLoadedItemsCsvFileName()}"`)
    res.write(`${buildLoadedItemsCsvHeader()}\r\n`)
    res.flushHeaders?.()

    while (page <= totalPages) {
      const pageResult = await searchIndexStore.search({
        scope: 'all',
        scopePath: '',
        mailboxKey: '',
        allowedMailboxKeys,
        reviewerUsername,
        sourceType: 'all',
        query: '',
        mode: 'and',
        mailOnly: false,
        sort: 'date-desc',
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: '',
        page,
        pageSize
      })
      totalPages = Math.max(1, pageResult.totalPages || 1)
      for (const item of pageResult.items) {
        itemCount += 1
        res.write(`${buildLoadedItemsCsvRow(buildLoadedReviewableItemFromSearchDocument(item))}\r\n`)
      }
      if (pageResult.items.length < pageSize) {
        break
      }
      page += 1
    }

    res.end()
    return itemCount
  }

  async function collectAllSearchIndexDocuments(
    options: Omit<SearchIndexSearchOptions, 'page' | 'pageSize'>,
    pageSize = 500
  ): Promise<SearchIndexDocument[]> {
    const documents: SearchIndexDocument[] = []
    let page = 1
    let totalPages = 1

    while (page <= totalPages) {
      const result = await searchIndexStore.search({
        ...options,
        page,
        pageSize
      })
      documents.push(...result.items)
      totalPages = Math.max(1, result.totalPages || 1)
      if (result.items.length < pageSize) {
        break
      }
      page += 1
    }

    return documents
  }

  async function collectLoadedReviewableItems(
    query: WorkspaceItemsRequestQuery,
    authSession: AuthSessionRecord | null
  ): Promise<{
    scopePath: string
    scopeLabel: string
    sourceType: SearchIndexDocument['sourceType']
    items: LoadedReviewableItem[]
  }> {
    const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
    const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
    const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
    const reviewerUsername = getReviewOwnerUsername(authSession)
    const filters = parseReviewFilters(query as Record<string, string | string[] | undefined>)
    const collapseDuplicates = parseBoolean(query.collapseDuplicates, false)
    const workspaceMode = normalizeText(query.workspaceMode).toLowerCase() === 'search' ? 'search' : 'folder'

    if (workspaceMode === 'folder') {
      const sessionId = normalizeText(query.sessionId)
      const folderId = normalizeText(query.folderId)
      if (!sessionId) {
        throw createAppError(400, 'Session id is required for folder exports')
      }
      if (!folderId) {
        throw createAppError(400, 'Folder id is required for folder exports')
      }

      const session = getSessionOrThrow(sessions, sessionId)
      if (!isScopePathAllowed(session.scopePath, allowedCasePaths, allowAllCases)) {
        throw createAppError(403, 'Case access required')
      }

      const hiddenRules = await searchIndexStore.listHiddenRules()
      const page = await buildReviewedFolderPage(
        session,
        folderId,
        {
          query: filters.query,
          mailOnly: filters.mailOnly,
          sort: filters.sort,
          page: 1,
          pageSize: Number.MAX_SAFE_INTEGER,
          reviewFlaggedOnly: filters.reviewFlaggedOnly,
          reviewTaggedOnly: filters.reviewTaggedOnly,
          reviewTag: filters.reviewTag,
          mode: filters.mode
        },
        reviewStore,
        hiddenRules,
        reviewerUsername
      )

      return {
        scopePath: session.scopePath,
        scopeLabel: session.scopeLabel || getScopeLabel(session.scopePath),
        sourceType: 'mailbox',
        items: page.items.map((item) => buildLoadedReviewableItemFromFolderSummary(session, item))
      }
    }

    const sourceType = parseSearchSourceType(query.sourceType)
    const requestedScope = parseSearchScope(query.scope) as SearchScope
    const requestedScopePath = normalizeScopePath(query.scopePath)
    const sessionId = normalizeText(query.sessionId)
    const scope = sourceType === 'mailbox' ? requestedScope : 'search'
    let scopePath = ''
    let scopeLabel = 'All cases/searches'
    let mailboxKey = ''
    let allowedMailboxKeys: string[] = []

    if (sourceType === 'mailbox' && scope === 'pst') {
      if (!sessionId) {
        throw createAppError(400, 'Session id is required for selected PST search')
      }
      const session = getSessionOrThrow(sessions, sessionId)
      if (!isScopePathAllowed(session.scopePath, allowedCasePaths, allowAllCases)) {
        throw createAppError(403, 'Case access required')
      }
      scopePath = session.scopePath
      scopeLabel = session.scopeLabel || getScopeLabel(scopePath)
      mailboxKey = session.filePath
      allowedMailboxKeys = [mailboxKey]
    } else if (sourceType === 'mailbox' && scope === 'search') {
      const catalog = resolveAccessibleCatalogSelection(
        pstRootDir,
        requestedScopePath,
        allowedCasePaths,
        listPstMailboxFiles,
        allowAllCases
      )
      scopePath = catalog.scopePath
      scopeLabel = catalog.scopeLabel
      const selectedCatalog = scopePath
        ? resolveAccessibleCatalogSelection(
            pstRootDir,
            scopePath,
            allowedCasePaths,
            listPstMailboxFiles,
            allowAllCases
          )
        : resolveAccessibleCatalogSelection(pstRootDir, '', allowedCasePaths, listPstMailboxFiles, allowAllCases)
      allowedMailboxKeys = selectedCatalog.files.map((file) =>
        resolvePstMailboxPath(pstRootDir, selectedCatalog.scopePath, file.fileName)
      )
    } else if (sourceType === 'mailbox' && scope === 'all') {
      const activeCatalog = resolveAccessibleCatalogSelection(
        pstRootDir,
        '',
        allowedCasePaths,
        listPstMailboxFiles,
        allowAllCases
      )
      allowedMailboxKeys = activeCatalog.scopes.flatMap((entry) =>
        entry.files.map((file) => resolvePstMailboxPath(pstRootDir, entry.scopePath, file.fileName))
      )
    } else {
      const archiveCatalog = resolveAccessibleArchiveCatalogSelection(
        pstRootDir,
        requestedScopePath,
        allowedCasePaths,
        allowAllCases
      )
      scopePath = archiveCatalog.scopePath
      scopeLabel = archiveCatalog.scopeLabel
      allowedMailboxKeys = buildArchiveMailboxKeys(pstRootDir, archiveCatalog)
    }

    const documents = await collectAllSearchIndexDocuments({
      scope,
      scopePath,
      mailboxKey,
      allowedMailboxKeys,
      reviewerUsername,
      sourceType,
      query: filters.query,
      mode: filters.mode,
      mailOnly: filters.mailOnly,
      sort: filters.sort,
      reviewFlaggedOnly: filters.reviewFlaggedOnly,
      reviewTaggedOnly: filters.reviewTaggedOnly,
      reviewTag: filters.reviewTag,
      collapseDuplicates
    })

    return {
      scopePath,
      scopeLabel,
      sourceType: sourceType === 'all' ? 'mailbox' : sourceType,
      items: documents.map((item) => buildLoadedReviewableItemFromSearchDocument(item))
    }
  }

  const searchIndexRefreshCoordinator =
    options.searchIndexRefreshCoordinator ||
    createSearchIndexRefreshCoordinator({
      pstRootDir,
      reviewStore,
      searchIndexStore,
      onJobComplete: (refreshStatus: SearchIndexRefreshStatus) => {
        recordAuditEvent({
          req: {
            headers: {},
            method: 'POST',
            originalUrl: API_ROUTES.searchIndexRefresh,
            url: API_ROUTES.searchIndexRefresh,
            ip: '',
            socket: { remoteAddress: '' }
          } as express.Request,
          session: null,
          action: `search.index.refresh.${refreshStatus.source}`,
          target: refreshStatus.source === 'items' ? 'Items index' : 'Mailbox index',
          outcome: refreshStatus.status === 'failed' ? 'failure' : 'success',
          metadata: {
            jobId: refreshStatus.jobId,
            trigger: refreshStatus.trigger,
            startedAt: refreshStatus.startedAt,
            completedAt: refreshStatus.completedAt,
            source: refreshStatus.source,
            mailboxCount: refreshStatus.summary?.mailboxCount || 0,
            messageCount: refreshStatus.summary?.messageCount || 0,
            error: refreshStatus.error || undefined
          }
        })
        if (refreshStatus.source === 'mailboxes' && refreshStatus.status === 'succeeded') {
          clearOpenMailboxDetailCaches()
        }
      }
    })

  app.set('searchIndexRefreshCoordinator', searchIndexRefreshCoordinator)

  function getAuthMfaChallengeCookieName(): string {
    return `${authConfig.cookieName}${DEFAULT_AUTH_MFA_CHALLENGE_COOKIE_SUFFIX}`
  }

  function getAuthPasswordChangeChallengeCookieName(): string {
    return `${authConfig.cookieName}${DEFAULT_AUTH_PASSWORD_CHANGE_CHALLENGE_COOKIE_SUFFIX}`
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
    for (const [token, challenge] of authPasswordChangeChallenges.entries()) {
      if (challenge.expiresAt <= now) {
        authPasswordChangeChallenges.delete(token)
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

  function getAuthPasswordChangeChallengeFromRequest(
    req: express.Request
  ): AuthPasswordChangeChallengeRecord | null {
    if (!authConfig.enabled) {
      return null
    }

    cleanupExpiredAuthSessions()
    const token = getCookieValue(req, getAuthPasswordChangeChallengeCookieName())
    if (!token) {
      return null
    }

    const challenge = authPasswordChangeChallenges.get(token) || null
    if (!challenge) {
      return null
    }

    if (challenge.expiresAt <= Date.now()) {
      authPasswordChangeChallenges.delete(token)
      return null
    }

    return challenge
  }

  function buildAuthStatus(
    session: AuthSessionRecord | null,
    challenge: AuthMfaChallengeRecord | null = null,
    user: AuthUserListItem | null = null,
    mfaEnabled = false,
    mfaEnforced = false,
    lockedUntil: string | null = null,
    loginFailedCount = 0,
    passwordResetAvailable = false,
    passwordChangeRequired = false,
    passwordChangeChallengeExpiresAt: string | null = null,
    entraEnabled = false
  ): AuthStatusResponse {
    const canManageUsers = isAdminAuthSession(session, authConfig)
    const authUser = user
      ? {
          username: user.username,
          assignedCasePaths: [...(user.assignedCasePaths || [])]
        }
      : null
    if (!authConfig.enabled) {
      return {
        authenticated: true,
        enabled: false,
        canManageUsers: false,
        entraEnabled: Boolean(entraEnabled),
        mfaEnabled: false,
        mfaEnforced: false,
        mfaRequired: false,
        mfaChallengeExpiresAt: null,
        lockedUntil: null,
        loginFailedCount: 0,
        passwordResetAvailable: false,
        passwordChangeRequired: false,
        passwordChangeChallengeExpiresAt: null,
        user: authUser,
        expiresAt: null
      }
    }

    if (!session) {
      if (passwordChangeRequired) {
        return {
          authenticated: false,
          enabled: true,
          canManageUsers: false,
          entraEnabled: Boolean(entraEnabled),
          mfaEnabled: false,
          mfaEnforced: Boolean(mfaEnforced),
          mfaRequired: false,
          mfaChallengeExpiresAt: null,
          lockedUntil: null,
          loginFailedCount: 0,
          passwordResetAvailable: false,
          passwordChangeRequired: true,
          passwordChangeChallengeExpiresAt: passwordChangeChallengeExpiresAt || null,
          user: authUser || {
            username: user?.username || '',
            assignedCasePaths: []
          },
          expiresAt: null
        }
      }

      if (challenge) {
        return {
          authenticated: false,
          enabled: true,
          canManageUsers: false,
          entraEnabled: Boolean(entraEnabled),
          mfaEnabled: false,
          mfaEnforced: Boolean(mfaEnforced),
          mfaRequired: true,
          mfaChallengeExpiresAt: new Date(challenge.expiresAt).toISOString(),
          lockedUntil: null,
          loginFailedCount: 0,
          passwordResetAvailable: false,
          passwordChangeRequired: false,
          passwordChangeChallengeExpiresAt: null,
          user: authUser || {
            username: challenge.username,
            assignedCasePaths: []
          },
          expiresAt: null
        }
      }

      return {
        authenticated: false,
        enabled: true,
        canManageUsers,
        entraEnabled: Boolean(entraEnabled),
        mfaEnabled: false,
        mfaEnforced: false,
        mfaRequired: false,
        mfaChallengeExpiresAt: null,
        lockedUntil: null,
        loginFailedCount: 0,
        passwordResetAvailable: Boolean(passwordResetAvailable),
        passwordChangeRequired: false,
        passwordChangeChallengeExpiresAt: null,
        user: null,
        expiresAt: null
      }
    }

    return {
      authenticated: true,
      enabled: true,
      canManageUsers,
      entraEnabled: Boolean(entraEnabled),
      mfaEnabled: Boolean(mfaEnabled),
      mfaEnforced: Boolean(mfaEnforced),
      mfaRequired: false,
      mfaChallengeExpiresAt: null,
      lockedUntil,
      loginFailedCount: Math.max(0, Math.floor(loginFailedCount || 0)),
      passwordResetAvailable: Boolean(passwordResetAvailable),
      passwordChangeRequired: false,
      passwordChangeChallengeExpiresAt: null,
      user: authUser || {
        username: session.username,
        assignedCasePaths: []
      },
      expiresAt: new Date(session.expiresAt).toISOString()
    }
  }

  async function getPasswordPolicy(): Promise<PasswordPolicyRecord> {
    try {
      return await appSettingsStore.getPasswordPolicy()
    } catch {
      return buildPasswordPolicyDefaultsFromEnv()
    }
  }

  async function getEntraSettings(): Promise<EntraSettingsRecord> {
    try {
      return await appSettingsStore.getEntraSettings()
    } catch {
      return {
        enabled: false,
        tenantId: '',
        clientId: '',
        clientSecret: ''
      }
    }
  }

  async function isEntraEnabled(): Promise<boolean> {
    const settings = await getEntraSettings()
    return Boolean(settings.enabled && normalizeText(settings.tenantId) && normalizeText(settings.clientId) && normalizeText(settings.clientSecret))
  }

  async function verifyEntraIdToken(
    idToken: string,
    settings: EntraSettingsRecord,
    expectedNonce: string
  ): Promise<EntraIdTokenClaims> {
    const { header, payload, signingInput, signature } = decodeJwt(idToken)
    const algorithm = normalizeText(header.alg)
    if (algorithm !== 'RS256') {
      throw createAppError(400, 'Unsupported identity token algorithm')
    }

    const kid = normalizeText(header.kid)
    if (!kid) {
      throw createAppError(400, 'Identity token key id is missing')
    }

    const jwksResponse = await fetch(buildEntraJwksUrl(settings), {
      headers: { Accept: 'application/json' }
    })
    if (!jwksResponse.ok) {
      throw createAppError(502, 'Unable to load Microsoft identity keys')
    }

    const jwksPayload = (await jwksResponse.json()) as { keys?: Array<Record<string, unknown>> }
    const jwk = (jwksPayload.keys || []).find((entry) => normalizeText(entry.kid) === kid) || null
    if (!jwk) {
      throw createAppError(400, 'Identity token key not found')
    }

    const publicKey = createPublicKey({ key: jwk as never, format: 'jwk' })
    const verifier = createVerify('RSA-SHA256')
    verifier.update(signingInput, 'utf8')
    verifier.end()
    if (!verifier.verify(publicKey, signature)) {
      throw createAppError(400, 'Invalid identity token signature')
    }

    const audience = Array.isArray(payload.aud)
      ? payload.aud.map((entry) => normalizeText(entry)).filter(Boolean)
      : [normalizeText(payload.aud)]
    if (!audience.includes(normalizeText(settings.clientId))) {
      throw createAppError(400, 'Identity token audience mismatch')
    }

    const now = Date.now()
    const expiresAt = Number(payload.exp || 0) * 1000
    if (!Number.isFinite(expiresAt) || expiresAt <= now) {
      throw createAppError(400, 'Identity token has expired')
    }

    const notBefore = Number(payload.nbf || 0) * 1000
    if (Number.isFinite(notBefore) && notBefore > now + 60_000) {
      throw createAppError(400, 'Identity token is not yet valid')
    }

    const tenantId = normalizeText(settings.tenantId).toLowerCase()
    const tokenTenantId = normalizeText(payload.tid).toLowerCase()
    if (tenantId && !['common', 'organizations'].includes(tenantId)) {
      if (tokenTenantId && tokenTenantId !== tenantId) {
        throw createAppError(403, 'Identity token tenant mismatch')
      }
      const expectedIssuer = `https://login.microsoftonline.com/${tenantId}/v2.0`
      if (normalizeText(payload.iss).toLowerCase() !== expectedIssuer.toLowerCase()) {
        throw createAppError(403, 'Identity token issuer mismatch')
      }
    }

    if (normalizeText(payload.nonce) !== normalizeText(expectedNonce)) {
      throw createAppError(400, 'Identity token nonce mismatch')
    }

    return payload
  }

  async function resolveEntraMappedUser(claims: EntraIdTokenClaims): Promise<AuthUserListItem> {
    const identifiers = [...new Set(
      [claims.email, claims.preferred_username, claims.upn]
        .map((value) => normalizeClaimValue(value))
        .filter(Boolean)
    )]

    if (!identifiers.length) {
      throw createAppError(400, 'Microsoft identity did not include an email or UPN')
    }

    const matches = new Map<string, AuthUserListItem>()
    for (const identifier of identifiers) {
      const results = await authUserStore.findUsersByLoginIdentifier(identifier)
      for (const user of results) {
        if (user.inviteStatus !== 'active') {
          continue
        }
        matches.set(normalizeAuthUsernameKey(user.username), user)
      }
    }

    if (!matches.size) {
      throw createAppError(403, 'No local user matches the Microsoft identity')
    }

    if (matches.size > 1) {
      throw createAppError(409, 'Microsoft identity matched multiple local users')
    }

    return matches.values().next().value as AuthUserListItem
  }

  async function buildEntraLoginRedirectUrl(req: express.Request, returnTo: string): Promise<string> {
    const settings = await getEntraSettings()
    if (
      !settings.enabled ||
      !normalizeText(settings.tenantId) ||
      !normalizeText(settings.clientId) ||
      !normalizeText(settings.clientSecret)
    ) {
      throw createAppError(400, 'Microsoft Entra sign-in is not configured')
    }

    const redirectUri = buildEntraRedirectUri(req)
    const state = createEntraLoginState(returnTo)
    return buildEntraAuthorizeUrl(settings, redirectUri, state.state, state.nonce, state.codeVerifier)
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

  function createAuthPasswordChangeChallenge(username: string): AuthPasswordChangeChallengeRecord {
    const challenge: AuthPasswordChangeChallengeRecord = {
      token: randomBytes(24).toString('hex'),
      username,
      expiresAt: Date.now() + 10 * 60 * 1000
    }
    authPasswordChangeChallenges.set(challenge.token, challenge)
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

  function clearAuthPasswordChangeChallenge(req: express.Request): void {
    if (!authConfig.enabled) {
      return
    }

    const token = getCookieValue(req, getAuthPasswordChangeChallengeCookieName())
    if (token) {
      authPasswordChangeChallenges.delete(token)
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

  function revokeAuthPasswordChangeChallengesForUsername(username: string): number {
    if (!authConfig.enabled) {
      return 0
    }

    const normalizedUsername = normalizeAuthUsernameKey(username)
    if (!normalizedUsername) {
      return 0
    }

    let revokedCount = 0
    for (const [token, challenge] of authPasswordChangeChallenges.entries()) {
      if (normalizeAuthUsernameKey(challenge.username) !== normalizedUsername) {
        continue
      }
      authPasswordChangeChallenges.delete(token)
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

  function setAuthPasswordChangeChallengeCookie(
    res: express.Response,
    req: express.Request,
    challenge: AuthPasswordChangeChallengeRecord
  ): void {
    if (!authConfig.enabled) {
      return
    }

    res.cookie(getAuthPasswordChangeChallengeCookieName(), challenge.token, {
      httpOnly: true,
      sameSite: 'lax',
      secure: Boolean(req.secure || String(req.headers['x-forwarded-proto'] || '').split(',')[0] === 'https'),
      path: '/',
      maxAge: 10 * 60 * 1000
    })
  }

  function clearAuthPasswordChangeChallengeCookie(res: express.Response): void {
    if (!authConfig.enabled) {
      return
    }

    res.clearCookie(getAuthPasswordChangeChallengeCookieName(), {
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

  function buildPasswordResetPath(token: string): string {
    return `/reset/${encodeURIComponent(token)}`
  }

  function buildPasswordResetUrl(req: express.Request, token: string): string {
    const origin =
      normalizeOrigin(authConfig.publicBaseUrl) || getRequestBaseOrigin(req) || canonicalRequestOrigin(req)
    const path = buildPasswordResetPath(token)
    return origin ? `${origin}${path}` : path
  }

  function generateTemporaryPassword(policy: PasswordPolicyRecord): string {
    const normalizedPolicy = {
      minLength: Math.max(12, Math.floor(policy?.minLength || 12)),
      requireUppercase: Boolean(policy?.requireUppercase),
      requireLowercase: Boolean(policy?.requireLowercase),
      requireNumber: Boolean(policy?.requireNumber),
      requireSpecial: Boolean(policy?.requireSpecial)
    }
    const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
    const lower = 'abcdefghijkmnopqrstuvwxyz'
    const numbers = '23456789'
    const special = '!@#$%^&*()-_=+[]{}'
    const pool = `${upper}${lower}${numbers}${special}`
    const length = Math.max(16, normalizedPolicy.minLength)

    for (let attempt = 0; attempt < 20; attempt += 1) {
      const chars: string[] = []
      if (normalizedPolicy.requireUppercase) {
        chars.push(upper[randomBytes(1)[0] % upper.length])
      }
      if (normalizedPolicy.requireLowercase) {
        chars.push(lower[randomBytes(1)[0] % lower.length])
      }
      if (normalizedPolicy.requireNumber) {
        chars.push(numbers[randomBytes(1)[0] % numbers.length])
      }
      if (normalizedPolicy.requireSpecial) {
        chars.push(special[randomBytes(1)[0] % special.length])
      }

      while (chars.length < length) {
        chars.push(pool[randomBytes(1)[0] % pool.length])
      }

      const shuffleBytes = randomBytes(chars.length)
      for (let index = chars.length - 1; index > 0; index -= 1) {
        const swapIndex = shuffleBytes[index] % (index + 1)
        ;[chars[index], chars[swapIndex]] = [chars[swapIndex], chars[index]]
      }

      const candidate = chars.join('')
      if (!validatePasswordAgainstPolicy(candidate, policy).length) {
        return candidate
      }
    }

    throw createAppError(500, 'Unable to generate a temporary password')
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

  function buildPasswordResetEmailText(input: {
    username: string
    resetUrl: string
    resetExpiresAt: string
  }): string {
    return [
      `A password reset was requested for ${input.username}.`,
      '',
      'Click here to reset your DV PST Mail Explorer password:',
      input.resetUrl,
      '',
      `This reset link expires at ${input.resetExpiresAt}.`
    ].join('\n')
  }

  function buildPasswordResetEmailHtml(input: {
    username: string
    resetUrl: string
    resetExpiresAt: string
  }): string {
    return [
      '<!doctype html>',
      '<html>',
      '<body style="margin:0;padding:24px;font-family:Arial,Helvetica,sans-serif;background:#f3f4f6;color:#111827;">',
      '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #d1d5db;border-radius:16px;padding:24px;">',
      `<h1 style="margin:0 0 12px;font-size:20px;line-height:28px;">Password reset requested</h1>`,
      `<p style="margin:0 0 16px;font-size:16px;line-height:24px;">A password reset was requested for <strong>${escapeHtml(
        input.username
      )}</strong>.</p>`,
      `<p style="margin:0 0 16px;font-size:16px;line-height:24px;"><a href="${escapeHtml(
        input.resetUrl
      )}" style="color:#2f6feb;text-decoration:underline;">Click here to reset your DV PST Mail Explorer password</a></p>`,
      `<p style="margin:0;font-size:14px;line-height:22px;color:#4b5563;">This reset link expires at ${escapeHtml(
        input.resetExpiresAt
      )}.</p>`,
      '</div>',
      '</body>',
      '</html>'
    ].join('')
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

  async function sendPasswordResetEmail(input: {
    recipientEmail: string
    username: string
    resetUrl: string
    resetExpiresAt: string
    req: express.Request
  }): Promise<{ emailSent: boolean; error?: string }> {
    try {
      const settings = await appSettingsStore.getSmtpSettings()
      if (!settings.enabled || !normalizeText(settings.host) || !normalizeText(settings.fromAddress)) {
        return { emailSent: false, error: 'SMTP is not configured' }
      }

      const transporter = smtpTransportFactory(settings)
      try {
        await transporter.sendMail({
          from: buildSmtpFromAddress(settings),
          to: input.recipientEmail,
          subject: 'DV PST Mail Explorer password reset',
          text: buildPasswordResetEmailText({
            username: input.username,
            resetUrl: input.resetUrl,
            resetExpiresAt: input.resetExpiresAt
          }),
          html: buildPasswordResetEmailHtml({
            username: input.username,
            resetUrl: input.resetUrl,
            resetExpiresAt: input.resetExpiresAt
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
    clearReusableMailboxSessions(normalizedMailboxKey)
    for (const [sessionId, session] of sessions.entries()) {
      if (session.filePath === normalizedMailboxKey) {
        invalidateMailboxDetailCache(session)
        session.index.messageDetailSnapshots.clear()
        sessions.delete(sessionId)
        closedSessionIds.push(sessionId)
      }
    }
    return closedSessionIds
  }

  function clearOpenMailboxDetailCaches(): void {
    clearReusableMailboxSessions()
    for (const session of sessions.values()) {
      invalidateMailboxDetailCache(session)
      session.index.messageDetailSnapshots.clear()
    }
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

  app.get('/reset/:token', (_req, res) => {
    res.status(200).sendFile(path.join(publicDir, 'index.html'))
  })

  app.get(API_ROUTES.authMe, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const session = getAuthSessionFromRequest(req)
      const passwordChangeChallenge = session ? null : getAuthPasswordChangeChallengeFromRequest(req)
      const challenge = session || passwordChangeChallenge ? null : getAuthMfaChallengeFromRequest(req)
      const passwordPolicy = await getPasswordPolicy()
      const entraEnabled = await isEntraEnabled()
      if (!authConfig.enabled && !session && !challenge && !passwordChangeChallenge) {
        responseJson(res, 200, buildAuthStatus(null, null, null, false, false, null, 0, false, false, null, entraEnabled))
        return
      }

      if (!session && !challenge && !passwordChangeChallenge) {
        responseJson(res, 401, {
          ...buildAuthStatus(null, null, null, false, false, null, 0, false, false, null, entraEnabled),
          error: 'Authentication required'
        })
        return
      }

      const currentUser = session
        ? await authUserStore.getUser(session.username)
        : passwordChangeChallenge
          ? await authUserStore.getUser(passwordChangeChallenge.username)
          : challenge
            ? await authUserStore.getUser(challenge.username)
            : null
      if (passwordChangeChallenge && !currentUser) {
        clearAuthPasswordChangeChallenge(req)
        clearAuthPasswordChangeChallengeCookie(res)
        responseJson(res, 401, {
          ...buildAuthStatus(null, null, null, false, false, null, 0, false, false, null, entraEnabled),
          error: 'Authentication required'
        })
        return
      }
      responseJson(
        res,
        200,
        passwordChangeChallenge
          ? buildAuthStatus(
              null,
              null,
              currentUser,
              false,
              Boolean(currentUser?.mfaEnforced || passwordPolicy.enforceMfa),
              null,
              0,
              false,
              true,
              new Date(passwordChangeChallenge.expiresAt).toISOString(),
              entraEnabled
            )
          : buildAuthStatus(
              session,
              challenge,
              currentUser,
              session?.mfaEnabled ?? Boolean(currentUser?.mfaEnabled),
              Boolean(currentUser?.mfaEnforced || passwordPolicy.enforceMfa),
              null,
              0,
              false,
              false,
              null,
              entraEnabled
            )
      )
    } catch (error) {
      if (res.headersSent) {
        res.destroy()
        return
      }
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authLogin, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const entraEnabled = await isEntraEnabled()
      if (!authConfig.enabled) {
        responseJson(res, 200, buildAuthStatus(null, null, null, false, false, null, 0, false, false, null, entraEnabled))
        return
      }

      const body = (req.body || {}) as { username?: string; email?: string; password?: string }
      const username = normalizeAuthUsername(body.username || body.email)
      const password = String(body.password ?? '')
      const passwordPolicy = await getPasswordPolicy()
      const result = await authUserStore.authenticate(username, password, passwordPolicy)
      if (!result.user) {
        const statusCode = result.lockedUntil ? 423 : 401
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
            reason: result.lockedUntil ? 'Account temporarily locked' : 'Invalid username or password',
            lockedUntil: result.lockedUntil || undefined,
            loginFailedCount: result.loginFailedCount
          }
        })
        responseJson(res, statusCode, {
          ...buildAuthStatus(
            null,
            null,
            null,
            false,
            false,
            result.lockedUntil,
            result.loginFailedCount,
            result.passwordResetAvailable,
            false,
            null,
            entraEnabled
          ),
          error: result.lockedUntil ? 'Account temporarily locked. Try again later.' : 'Invalid username or password'
        })
        return
      }

      const user = result.user
      const mfaEnforced = Boolean(user.mfaEnforced || passwordPolicy.enforceMfa)
      clearAuthMfaChallenge(req)
      clearAuthMfaChallengeCookie(res)
      clearAuthPasswordChangeChallenge(req)
      clearAuthPasswordChangeChallengeCookie(res)

      if (result.passwordChangeRequired) {
        const challenge = createAuthPasswordChangeChallenge(user.username)
        setAuthPasswordChangeChallengeCookie(res, req, challenge)
        recordAuditEvent({
          req,
          actor: buildAuditActor(null, user.username),
          action: 'auth.login',
          target: user.username,
          outcome: 'success',
          metadata: {
            passwordChangeRequired: true,
            passwordChangeChallengeExpiresAt: new Date(challenge.expiresAt).toISOString(),
            mfaEnforced
          }
        })
        responseJson(
          res,
          200,
          buildAuthStatus(
            null,
            null,
            user,
            false,
            mfaEnforced,
            null,
            0,
            false,
            true,
            new Date(challenge.expiresAt).toISOString(),
            entraEnabled
          )
        )
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
            mfaEnforced
          }
        })
        responseJson(res, 200, buildAuthStatus(null, challenge, user, false, mfaEnforced, null, 0, false, false, null, entraEnabled))
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
          mfaEnforced
        }
      })
      responseJson(
        res,
        200,
        buildAuthStatus(session, null, user, Boolean(user.mfaEnabled), mfaEnforced, null, 0, false, false, null, entraEnabled)
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authLogout, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const entraEnabled = await isEntraEnabled()
      if (authConfig.enabled) {
        const session = getAuthSessionFromRequest(req)
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
        clearAuthPasswordChangeChallenge(req)
        clearAuthCookie(res)
        clearAuthMfaChallengeCookie(res)
        clearAuthPasswordChangeChallengeCookie(res)
      }
      responseJson(res, 200, buildAuthStatus(null, null, null, false, false, null, 0, false, false, null, entraEnabled))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.authEntraStart, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const returnTo = sanitizeReturnTo(
        (req.query.returnTo as string | undefined) ||
          (req.query.redirectTo as string | undefined) ||
          (req.headers.referer as string | undefined) ||
          '/'
      )
      const redirectUrl = await buildEntraLoginRedirectUrl(req, returnTo)
      res.redirect(302, redirectUrl)
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to start Microsoft sign in'
      const fallbackUrl = `/${message ? `?entraError=${encodeURIComponent(message)}` : ''}`
      if (!res.headersSent) {
        res.redirect(302, fallbackUrl)
      }
    }
  })

  app.get(API_ROUTES.authEntraCallback, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const error = normalizeText(req.query.error)
      if (error) {
        const description = normalizeText(req.query.error_description || req.query.errorDescription || error)
        res.redirect(302, `/?entraError=${encodeURIComponent(description || error)}`)
        return
      }

      const code = normalizeText(req.query.code)
      const state = normalizeText(req.query.state)
      if (!code || !state) {
        res.redirect(302, '/?entraError=Microsoft%20sign-in%20did%20not%20return%20a%20code')
        return
      }

      const loginState = consumeEntraLoginState(state)
      if (!loginState) {
        res.redirect(302, '/?entraError=Microsoft%20sign-in%20expired')
        return
      }

      const settings = await getEntraSettings()
      if (
        !settings.enabled ||
        !normalizeText(settings.tenantId) ||
        !normalizeText(settings.clientId) ||
        !normalizeText(settings.clientSecret)
      ) {
        res.redirect(302, '/?entraError=Microsoft%20sign-in%20is%20not%20configured')
        return
      }

      const redirectUri = buildEntraRedirectUri(req)
      const tokenParams = new URLSearchParams({
        client_id: normalizeText(settings.clientId),
        client_secret: normalizeText(settings.clientSecret),
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        code_verifier: loginState.codeVerifier
      })
      const tokenResponse = await fetch(buildEntraTokenUrl(settings), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Accept: 'application/json'
        },
        body: tokenParams.toString()
      })
      const tokenPayload = (await tokenResponse.json()) as EntraTokenResponse
      if (!tokenResponse.ok) {
        throw createAppError(
          502,
          normalizeText(tokenPayload.error_description || tokenPayload.error || 'Unable to complete Microsoft sign in')
        )
      }

      const idToken = normalizeText(tokenPayload.id_token)
      if (!idToken) {
        throw createAppError(502, 'Microsoft sign in did not return an identity token')
      }

      const claims = await verifyEntraIdToken(idToken, settings, loginState.nonce)
      const mappedUser = await resolveEntraMappedUser(claims)
      clearAuthMfaChallenge(req)
      clearAuthMfaChallengeCookie(res)
      clearAuthPasswordChangeChallenge(req)
      clearAuthPasswordChangeChallengeCookie(res)

      const session = createAuthSession(mappedUser.username, Boolean(mappedUser.mfaEnabled))
      setAuthCookie(res, req, session)
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, mappedUser.username),
        action: 'auth.entra.login',
        target: mappedUser.username,
        outcome: 'success',
        metadata: {
          tenantId: normalizeText(claims.tid),
          email: normalizeText(claims.email),
          preferredUsername: normalizeText(claims.preferred_username),
          upn: normalizeText(claims.upn)
        }
      })

      res.redirect(302, loginState.returnTo || '/')
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unable to complete Microsoft sign in'
      if (!res.headersSent) {
        res.redirect(302, `/?entraError=${encodeURIComponent(message)}`)
      }
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

  app.put(API_ROUTES.authUserAccess, async (req, res) => {
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
          action: 'auth.users.access.update',
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

      const body = (req.body || {}) as { assignedCasePaths?: unknown }
      const assignedCasePaths = normalizeAssignedCasePaths(body.assignedCasePaths)
      const user = await authUserStore.setAssignedCasePaths(targetUsername, assignedCasePaths)
      if (!user) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.access.update',
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
        action: 'auth.users.access.update',
        target: user.username,
        outcome: 'success',
        metadata: {
          assignedCasePaths: [...user.assignedCasePaths]
        }
      })
      responseJson(res, 200, {
        user
      } satisfies AuthUserDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authUserPasswordReset, async (req, res) => {
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
          action: 'auth.users.password.reset',
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

      const body = (req.body || {}) as { mode?: string }
      const mode = normalizeText(body.mode).toLowerCase() === 'temporary' ? 'temporary' : 'link'
      const existingUser = await authUserStore.getUser(targetUsername)
      if (!existingUser) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.password.reset',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User not found',
            mode
          }
        })
        responseJson(res, 404, {
          error: 'User not found'
        })
        return
      }
      if (existingUser.inviteStatus !== 'active') {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.password.reset',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'User is not active',
            mode
          }
        })
        responseJson(res, 409, {
          error: 'User is not active'
        })
        return
      }

      const passwordPolicy = await getPasswordPolicy()
      if (mode === 'temporary') {
        const temporaryPassword = generateTemporaryPassword(passwordPolicy)
        const user = await authUserStore.changePassword(
          targetUsername,
          temporaryPassword,
          passwordPolicy,
          true
        )
        const revokedSessions = revokeAuthSessionsForUsername(user.username)
        const revokedChallenges = revokeAuthMfaChallengesForUsername(user.username)
        const revokedPasswordChangeChallenges = revokeAuthPasswordChangeChallengesForUsername(user.username)
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.password.reset',
          target: user.username,
          outcome: 'success',
          metadata: {
            mode,
            revokedSessions,
            revokedChallenges,
            revokedPasswordChangeChallenges
          }
        })
        responseJson(res, 200, {
          user,
          mode,
          temporaryPassword
        } satisfies AuthUserPasswordResetResponse)
        return
      }

      const result = await authUserStore.requestPasswordReset(
        targetUsername,
        passwordPolicy.resetTokenTtlMinutes,
        passwordPolicy,
        {
          bypassGate: true,
          allowMissingRecipient: true
        }
      )
      if (!result) {
        recordAuditEvent({
          req,
          session,
          action: 'auth.users.password.reset',
          target: targetUsername,
          outcome: 'denied',
          metadata: {
            reason: 'Password reset unavailable',
            mode
          }
        })
        responseJson(res, 404, {
          error: 'Password reset unavailable'
        })
        return
      }

      const resetUrl = buildPasswordResetUrl(req, result.resetToken)
      let emailSent = false
      let emailError = ''
      if (result.user.recipientEmail) {
        const emailResult = await sendPasswordResetEmail({
          recipientEmail: result.user.recipientEmail,
          username: result.user.username,
          resetUrl,
          resetExpiresAt: result.resetExpiresAt,
          req
        })
        emailSent = emailResult.emailSent
        emailError = emailResult.error || ''
      } else {
        emailError = 'Recipient email is not configured'
      }

      recordAuditEvent({
        req,
        session,
        action: 'auth.users.password.reset',
        target: result.user.username,
        outcome: 'success',
        metadata: {
          mode,
          emailSent,
          emailError: emailError || undefined,
          resetExpiresAt: result.resetExpiresAt
        }
      })
      responseJson(res, 200, {
        user: result.user,
        mode,
        resetUrl,
        resetExpiresAt: result.resetExpiresAt,
        emailSent,
        emailError: emailError || undefined
      } satisfies AuthUserPasswordResetResponse)
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

      const passwordPolicy = await getPasswordPolicy()
      const user = await authUserStore.acceptInvite(token, password, passwordPolicy)
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

  app.post(API_ROUTES.authPasswordResetRequest, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        responseJson(res, 200, { sent: true })
        return
      }

      const body = (req.body || {}) as { usernameOrEmail?: string }
      const usernameOrEmail = normalizeText(body.usernameOrEmail)
      const passwordPolicy = await getPasswordPolicy()
      const result = await authUserStore.requestPasswordReset(
        usernameOrEmail,
        passwordPolicy.resetTokenTtlMinutes,
        passwordPolicy
      )
      if (result) {
        const resetUrl = buildPasswordResetUrl(req, result.resetToken)
        const emailResult = await sendPasswordResetEmail({
          recipientEmail: result.user.recipientEmail,
          username: result.user.username,
          resetUrl,
          resetExpiresAt: result.resetExpiresAt,
          req
        })
        recordAuditEvent({
          req,
          actor: buildAuditActor(null, result.user.username),
          action: 'auth.password.reset.request',
          target: result.user.username,
          outcome: 'success',
          metadata: {
            emailSent: emailResult.emailSent,
            resetExpiresAt: result.resetExpiresAt
          }
        })
      } else {
        recordAuditEvent({
          req,
          actor: buildAuditActor(null, usernameOrEmail || 'anonymous'),
          action: 'auth.password.reset.request',
          target: usernameOrEmail || 'anonymous',
          outcome: 'denied',
          metadata: {
            reason: 'Account not found or reset unavailable'
          }
        })
      }

      responseJson(res, 200, {
        sent: true
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.authPasswordResetLookup, async (req, res) => {
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
          error: 'Password reset token is required'
        })
        return
      }

      const user = await authUserStore.getPasswordResetByToken(token)
      if (!user) {
        responseJson(res, 404, {
          error: 'Password reset token not found'
        })
        return
      }

      responseJson(res, 200, {
        reset: {
          username: user.username,
          recipientEmail: user.recipientEmail
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authPasswordResetConfirm, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const token = normalizeText(req.params.token)
      if (!token) {
        responseJson(res, 400, {
          error: 'Password reset token is required'
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

      const passwordPolicy = await getPasswordPolicy()
      const user = await authUserStore.resetPassword(token, password, passwordPolicy)
      const revokedSessions = revokeAuthSessionsForUsername(user.username)
      const revokedMfaChallenges = revokeAuthMfaChallengesForUsername(user.username)
      const revokedPasswordChangeChallenges = revokeAuthPasswordChangeChallengesForUsername(user.username)
      clearAuthMfaChallenge(req)
      clearAuthMfaChallengeCookie(res)
      clearAuthPasswordChangeChallenge(req)
      clearAuthPasswordChangeChallengeCookie(res)
      recordAuditEvent({
        req,
        actor: buildAuditActor(null, user.username),
        action: 'auth.password.reset.complete',
        target: user.username,
        outcome: 'success',
        metadata: {
          revokedSessions,
          revokedMfaChallenges,
          revokedPasswordChangeChallenges
        }
      })
      responseJson(res, 200, {
        user,
        message: 'Password updated'
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.authPasswordChangeConfirm, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      if (!authConfig.enabled) {
        throw createAppError(400, 'Authentication is disabled')
      }

      const challenge = getAuthPasswordChangeChallengeFromRequest(req)
      if (!challenge) {
        responseJson(res, 401, {
          error: 'Password change challenge required'
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

      const passwordPolicy = await getPasswordPolicy()
      const user = await authUserStore.changePassword(challenge.username, password, passwordPolicy, false)
      const revokedSessions = revokeAuthSessionsForUsername(user.username)
      const revokedMfaChallenges = revokeAuthMfaChallengesForUsername(user.username)
      const revokedPasswordChangeChallenges = revokeAuthPasswordChangeChallengesForUsername(user.username)
      clearAuthPasswordChangeChallenge(req)
      clearAuthPasswordChangeChallengeCookie(res)
      clearAuthMfaChallenge(req)
      clearAuthMfaChallengeCookie(res)

      if (user.mfaEnabled) {
        const mfaChallenge = createAuthMfaChallenge(user.username)
        setAuthMfaChallengeCookie(res, req, mfaChallenge)
        recordAuditEvent({
          req,
          actor: buildAuditActor(null, user.username),
          action: 'auth.password.change.complete',
          target: user.username,
          outcome: 'success',
          metadata: {
            revokedSessions,
            revokedMfaChallenges,
            revokedPasswordChangeChallenges,
            mfaRequired: true
          }
        })
        responseJson(
          res,
          200,
          buildAuthStatus(
            null,
            mfaChallenge,
            user,
            false,
            Boolean(user.mfaEnforced || passwordPolicy.enforceMfa)
          )
        )
        return
      }

      const session = createAuthSession(user.username, Boolean(user.mfaEnabled))
      setAuthCookie(res, req, session)
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, user.username),
        action: 'auth.password.change.complete',
        target: user.username,
        outcome: 'success',
        metadata: {
          revokedSessions,
          revokedMfaChallenges,
          revokedPasswordChangeChallenges,
          mfaRequired: false
        }
      })
      responseJson(
        res,
        200,
        buildAuthStatus(
          session,
          null,
          user,
          Boolean(user.mfaEnabled),
          Boolean(user.mfaEnforced || passwordPolicy.enforceMfa)
        )
      )
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
      const passwordPolicy = await getPasswordPolicy()
      recordAuditEvent({
        req,
        actor: buildAuditActor(session, user.username),
        action: 'auth.mfa.challenge',
        target: user.username,
        outcome: 'success'
      })
      responseJson(
        res,
        200,
        buildAuthStatus(
          session,
          null,
          user,
          Boolean(user.mfaEnabled),
          Boolean(user.mfaEnforced || passwordPolicy.enforceMfa)
        )
      )
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

  app.get(API_ROUTES.authEntraSettings, async (req, res) => {
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
          action: 'settings.entra.read',
          target: 'entra settings',
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

      const settings = await appSettingsStore.getEntraSettings()
      responseJson(res, 200, {
        settings: buildEntraSettingsView(settings),
        redirectUri: buildEntraRedirectUri(req)
      } satisfies EntraSettingsResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.put(API_ROUTES.authEntraSettings, async (req, res) => {
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
          action: 'settings.entra.update',
          target: 'entra settings',
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

      const body = normalizeEntraSettingsInput((req.body || {}) as EntraSettingsInput)
      const updated = await appSettingsStore.updateEntraSettings(body)
      recordAuditEvent({
        req,
        session,
        action: 'settings.entra.update',
        target: normalizeText(body.tenantId || updated.tenantId || 'entra settings'),
        outcome: 'success',
        metadata: {
          enabled: updated.enabled,
          tenantId: updated.tenantId,
          clientId: updated.clientId,
          hasClientSecret: Boolean(updated.clientSecret)
        }
      })
      responseJson(res, 200, {
        settings: buildEntraSettingsView(updated),
        redirectUri: buildEntraRedirectUri(req)
      } satisfies EntraSettingsResponse)
    } catch (error) {
      const session = getAuthSessionFromRequest(req)
      if (session) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.entra.update',
          target: 'entra settings',
          outcome: 'failure',
          metadata: {
            reason: error instanceof Error ? error.message : String(error)
          }
        })
      }
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.passwordPolicySettings, async (req, res) => {
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
          action: 'settings.password_policy.read',
          target: 'password policy',
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
        settings: await appSettingsStore.getPasswordPolicy()
      } satisfies PasswordPolicyResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.put(API_ROUTES.passwordPolicySettings, async (req, res) => {
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
          action: 'settings.password_policy.update',
          target: 'password policy',
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

      const body = normalizePasswordPolicyInput((req.body || {}) as PasswordPolicyInput)
      const updated = await appSettingsStore.updatePasswordPolicy(body)
      recordAuditEvent({
        req,
        session,
        action: 'settings.password_policy.update',
        target: 'password policy',
        outcome: 'success',
        metadata: {
          minLength: updated.minLength,
          requireUppercase: updated.requireUppercase,
          requireLowercase: updated.requireLowercase,
          requireNumber: updated.requireNumber,
          requireSpecial: updated.requireSpecial,
          forgotPasswordAfterFailures: updated.forgotPasswordAfterFailures,
          lockoutThreshold: updated.lockoutThreshold,
          lockoutDurationSeconds: updated.lockoutDurationSeconds,
          resetTokenTtlMinutes: updated.resetTokenTtlMinutes,
          enforceMfa: updated.enforceMfa
        }
      })
      responseJson(res, 200, {
        settings: updated
      } satisfies PasswordPolicyResponse)
    } catch (error) {
      const session = getAuthSessionFromRequest(req)
      if (session) {
        recordAuditEvent({
          req,
          session,
          action: 'settings.password_policy.update',
          target: 'password policy',
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

  app.get(API_ROUTES.pstCatalog, async (req, res) => {
    try {
      const session = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = session ? await authUserStore.getUser(session.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(session, authConfig)
      const scopePath =
        typeof req.query.scopePath === 'string' ? normalizeText(req.query.scopePath) : ''
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const catalog = resolveAccessibleCatalogSelection(
        pstRootDir,
        scopePath,
        allowedCasePaths,
        listPstMailboxFiles,
        allowAllCases
      )
      responseJson(res, 200, catalog)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.pstRemovedCatalog, async (req, res) => {
    try {
      const session = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !session) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = session ? await authUserStore.getUser(session.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(session, authConfig)
      const scopePath =
        typeof req.query.scopePath === 'string' ? normalizeText(req.query.scopePath) : ''
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const catalog = resolveAccessibleCatalogSelection(
        pstRootDir,
        scopePath,
        allowedCasePaths,
        listRemovedPstMailboxFiles,
        allowAllCases
      )
      responseJson(res, 200, catalog)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.pstOpen, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const body = (req.body || {}) as OpenMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      if (!isScopePathAllowed(scopePath, allowAllCases ? [] : getAccessibleCasePaths(currentUser), allowAllCases)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'mailbox.open',
          target: `${scopePath ? `${scopePath}/` : ''}${fileName}`,
          outcome: 'denied',
          metadata: {
            scopePath,
            scopeLabel: getScopeLabel(scopePath),
            fileName,
            reason: 'Case access required'
          }
        })
        responseJson(res, 403, {
          error: 'Case access required'
        })
        return
      }

      const mailboxKey = resolvePstMailboxPath(pstRootDir, scopePath, fileName)
      const scopeLabel = scopePath ? scopePath.split('/').join(' / ') : 'PST root'
      const existingSessionId = reusableMailboxSessions.get(mailboxKey)
      if (existingSessionId) {
        const existingSession = sessions.get(existingSessionId)
        if (existingSession) {
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
              messageCount: existingSession.index.messages.size,
              reused: true
            }
          })
          responseJson(res, 200, buildSessionResponse(existingSession))
          return
        }
        reusableMailboxSessions.delete(mailboxKey)
      }

      const pendingSession = openingMailboxSessions.get(mailboxKey)
      if (pendingSession) {
        const response = await pendingSession
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
            messageCount: response.summary.stats.messageCount,
            reused: true
          }
        })
        responseJson(res, 200, response)
        return
      }

      const openPromise = (async () => {
        const index = openPstMailbox(pstRootDir, scopePath, fileName)
        const record = registerMailboxSession(index, mailboxKey, scopePath, fileName)
        return buildSessionResponse(record)
      })()

      openingMailboxSessions.set(mailboxKey, openPromise)
      try {
        const response = await openPromise
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
            messageCount: response.summary.stats.messageCount,
            reused: false
          }
        })
        responseJson(res, 200, response)
      } finally {
        if (openingMailboxSessions.get(mailboxKey) === openPromise) {
          openingMailboxSessions.delete(mailboxKey)
        }
      }
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.pstRemove, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const body = (req.body || {}) as MoveMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      if (!isScopePathAllowed(scopePath, allowAllCases ? [] : getAccessibleCasePaths(currentUser), allowAllCases)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'mailbox.remove',
          target: `${scopePath ? `${scopePath}/` : ''}${fileName}`,
          outcome: 'denied',
          metadata: {
            scopePath,
            scopeLabel: getScopeLabel(scopePath),
            fileName,
            reason: 'Case access required'
          }
        })
        responseJson(res, 403, {
          error: 'Case access required'
        })
        return
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
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const body = (req.body || {}) as MoveMailboxRequestBody
      const scopePath = normalizeText(body.scopePath)
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      if (!isScopePathAllowed(scopePath, allowAllCases ? [] : getAccessibleCasePaths(currentUser), allowAllCases)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'mailbox.restore',
          target: `${scopePath ? `${scopePath}/` : ''}${fileName}`,
          outcome: 'denied',
          metadata: {
            scopePath,
            scopeLabel: getScopeLabel(scopePath),
            fileName,
            reason: 'Case access required'
          }
        })
        responseJson(res, 403, {
          error: 'Case access required'
        })
        return
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
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const collapseDuplicates = parseBoolean(req.query.collapseDuplicates, false)
      const sourceType = parseSearchSourceType(req.query.sourceType)
      const requestedScope = parseSearchScope(req.query.scope) as SearchScope
      const requestedScopePath = normalizeScopePath(req.query.scopePath)
      const requestedCasePath = normalizeScopePath(req.query.casePath)
      const sessionId = normalizeText(req.query.sessionId)
      const scope = sourceType === 'mailbox' ? requestedScope : 'search'
      let scopePath = ''
      let scopeLabel = 'All cases/searches'
      let mailboxKey = ''
      let allowedMailboxKeys: string[] = []

      if (requestedCasePath && !isScopePathAllowed(requestedCasePath, allowedCasePaths, allowAllCases)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'search.execute',
          target: requestedCasePath,
          outcome: 'denied',
          metadata: {
            reason: 'Case access required',
            casePath: requestedCasePath
          }
        })
        responseJson(res, 403, {
          error: 'Case access required'
        })
        return
      }

      const mailboxSearchFingerprints =
        sourceType === 'mailbox' || sourceType === 'all'
          ? await searchIndexStore.listFileFingerprints('mailboxes')
          : []
      const archiveSearchFingerprints =
        sourceType === 'mailbox' ? [] : await searchIndexStore.listFileFingerprints('items')

      if (sourceType === 'mailbox' && scope === 'pst') {
        if (!sessionId) {
          throw createAppError(400, 'Session id is required for selected PST search')
        }
        const session = getSessionOrThrow(sessions, sessionId)
        if (!isScopePathAllowed(session.scopePath, allowedCasePaths, allowAllCases)) {
          recordAuditEvent({
            req,
            session: authSession,
            action: 'search.execute',
            target: session.scopeLabel || session.fileName,
            outcome: 'denied',
            metadata: {
              reason: 'Case access required',
              scope: 'pst',
              scopePath: session.scopePath
            }
          })
          responseJson(res, 403, {
            error: 'Case access required'
          })
          return
        }
        scopePath = session.scopePath
        scopeLabel = session.scopeLabel || getScopeLabel(scopePath)
        mailboxKey = session.filePath
        allowedMailboxKeys = [mailboxKey]
      } else if (sourceType === 'mailbox' && scope === 'search') {
        const fingerprintSelection = resolveAccessibleFingerprintCatalogSelection(
          requestedScopePath,
          allowedCasePaths,
          mailboxSearchFingerprints,
          allowAllCases
        )
        if (fingerprintSelection) {
          scopePath = fingerprintSelection.scopePath
          scopeLabel = fingerprintSelection.scopeLabel
          allowedMailboxKeys = fingerprintSelection.files.map((file) => file.mailboxKey)
        } else {
          const catalog = resolveAccessibleCatalogSelection(
            pstRootDir,
            requestedScopePath,
            allowedCasePaths,
            listPstMailboxFiles,
            allowAllCases
          )
          scopePath = catalog.scopePath
          scopeLabel = catalog.scopeLabel
          allowedMailboxKeys = catalog.files.map((file) =>
            resolvePstMailboxPath(pstRootDir, catalog.scopePath, file.fileName)
          )
        }
      } else if (sourceType === 'mailbox' && scope === 'all') {
        const fingerprintSelection = resolveAccessibleFingerprintCatalogSelection(
          '',
          allowedCasePaths,
          mailboxSearchFingerprints,
          allowAllCases
        )
        if (fingerprintSelection) {
          allowedMailboxKeys = collectFingerprintMailboxKeys(fingerprintSelection.scopes)
        } else {
          const activeCatalog = resolveAccessibleCatalogSelection(
            pstRootDir,
            '',
            allowedCasePaths,
            listPstMailboxFiles,
            allowAllCases
          )
          allowedMailboxKeys = activeCatalog.scopes.flatMap((entry) =>
            entry.files.map((file) => resolvePstMailboxPath(pstRootDir, entry.scopePath, file.fileName))
          )
        }
      } else if (sourceType === 'all') {
        const mailboxSelection = selectFingerprintMailboxKeys(
          requestedCasePath,
          requestedScopePath,
          allowedCasePaths,
          mailboxSearchFingerprints,
          allowAllCases
        )
        const archiveSelection = selectFingerprintMailboxKeys(
          requestedCasePath,
          requestedScopePath,
          allowedCasePaths,
          archiveSearchFingerprints,
          allowAllCases
        )
        const mergedSelection = mergeFingerprintMailboxSelections(mailboxSelection, archiveSelection)
        scopePath = mergedSelection.scopePath
        scopeLabel = mergedSelection.scopeLabel
        allowedMailboxKeys = mergedSelection.mailboxKeys
      } else {
        const archiveSelection = selectFingerprintMailboxKeys(
          requestedCasePath,
          requestedScopePath,
          allowedCasePaths,
          archiveSearchFingerprints,
          allowAllCases
        )
        scopePath = archiveSelection.scopePath
        scopeLabel = archiveSelection.scopeLabel
        allowedMailboxKeys = archiveSelection.mailboxKeys
      }

      const page = await searchIndexStore.search({
        scope,
        scopePath,
        mailboxKey,
        allowedMailboxKeys,
        reviewerUsername,
        sourceType,
        requirePreviewPayload: true,
        casePath: requestedCasePath || undefined,
        query: filters.query,
        mode: filters.mode,
        mailOnly: filters.mailOnly,
        sort: filters.sort,
        page: filters.page,
        pageSize: filters.pageSize,
        reviewFlaggedOnly: filters.reviewFlaggedOnly,
        reviewTaggedOnly: filters.reviewTaggedOnly,
        reviewTag: filters.reviewTag,
        collapseDuplicates
      })
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.execute',
        target: scopeLabel,
        outcome: 'success',
        metadata: {
          sourceType,
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
        sourceType,
        page: {
          ...page,
          items: page.items.map((item) => buildSearchResultSummary(item))
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemDetail, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: 'search.item.view',
          target: item.subject || item.archiveEntryName || item.messageId,
          outcome: 'denied',
          metadata: {
            reason: 'Case access required',
            scopePath: item.scopePath,
            sourceType: item.sourceType
          }
        })
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const reviewerUsername = getReviewOwnerUsername(authSession)
      const detail = await resolveSearchItemDetail(
        item,
        reviewStore,
        reviewerUsername
      )

      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.item.view',
        target: detail.subject || item.archiveEntryName || detail.id || 'item',
        outcome: 'success',
        metadata: {
          sourceType: item.sourceType,
          scopePath: item.scopePath,
          fileName: item.fileName,
          archivePath: item.archivePath || '',
          entryPath: item.archiveEntryPath || ''
        }
      })

      responseJson(res, 200, { detail })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemContent, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item || !item.archivePath || !item.archiveEntryChain?.length) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const { buffer, contentType, fileName } = await readArchiveBundleItemContent(item.archivePath, item.archiveEntryChain)
      const downloadName = sanitizeFileNameForDownload(item.downloadFilename || fileName, fileName)
      res.status(200)
      res.setHeader('Content-Type', contentType || item.contentType || 'application/octet-stream')
      res.setHeader('Content-Length', String(buffer.length))
      res.setHeader('Content-Disposition', `inline; filename="${downloadName}"`)
      res.send(buffer)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemAttachment, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item || item.sourceType !== 'mailbox') {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const attachmentIndex = parseZeroBasedInt(req.params.attachmentIndex)
      if (attachmentIndex < 0) {
        responseJson(res, 404, { error: 'Attachment not found' })
        return
      }

      const liveSession = getLiveSessionForMailboxKey(item.mailboxKey)
      const sessionIndex = liveSession?.index || openPstMailbox(pstRootDir, item.scopePath, item.fileName)
      const payload = getAttachmentDownloadBuffer(sessionIndex, item.messageId, attachmentIndex)
      const fileName = safeDownloadName(payload.filename, 'attachment')
      responseBinary(res, 200, payload.contentType, fileName, payload.buffer)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'item.attachment.download',
        target: item.subject || item.messageId,
        outcome: 'success',
        metadata: {
          scopePath: item.scopePath,
          fileName: item.fileName,
          attachmentIndex,
          downloadName: fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemExportEml, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item || item.sourceType !== 'mailbox') {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const liveSession = getLiveSessionForMailboxKey(item.mailboxKey)
      const sessionIndex = liveSession?.index || openPstMailbox(pstRootDir, item.scopePath, item.fileName)
      const summary = getMessageSummary(sessionIndex, item.messageId)
      const fileName = `${safeDownloadName(summary.subject || item.subject || 'message', 'message')}.eml`
      const eml = summary.parseError
        ? exportMessageAsEml(buildReviewedDetail(buildEmptyMessageDetail(summary), null))
        : exportMessageAsEmlFromSession(sessionIndex, item.messageId)
      responseBinary(res, 200, 'message/rfc822; charset=utf-8', fileName, Buffer.from(eml, 'utf8'))
      recordAuditEvent({
        req,
        session: authSession,
        action: 'item.export.eml',
        target: item.subject || item.messageId,
        outcome: 'success',
        metadata: {
          scopePath: item.scopePath,
          fileName: item.fileName,
          downloadName: fileName
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemPreview, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item || !item.archivePath || !item.archiveEntryChain?.length) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const { buffer, contentType, fileName } = await readArchiveBundleItemContent(item.archivePath, item.archiveEntryChain)
      const responseFileName = sanitizeFileNameForDownload(item.downloadFilename || fileName, fileName)
      const sourceContentType = contentType || item.contentType || 'application/octet-stream'
      let responseBuffer = buffer
      let responseContentType = sourceContentType
      let responseDownloadName = responseFileName

      if (isOfficePreviewable(sourceContentType, responseFileName)) {
        const preview = await buildOfficePreview({
          cacheKey: item.id || item.messageId || `${item.archivePath}:${item.archiveEntryPath || ''}`,
          fileName: responseFileName,
          contentType: sourceContentType,
          buffer,
          previewText: item.previewText || ''
        })
        responseBuffer = preview.buffer
        responseContentType = preview.contentType
        responseDownloadName = preview.fileName
      }

      res.status(200)
      res.setHeader('Content-Type', responseContentType)
      res.setHeader('Content-Length', String(responseBuffer.length))
      res.setHeader(
        'Content-Disposition',
        `inline; filename="${sanitizeFileNameForDownload(responseDownloadName, responseDownloadName)}"`
      )
      res.send(responseBuffer)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.itemReview, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const reviewerUsername = getReviewOwnerUsername(authSession)
      const review = normalizeReviewState(
        await reviewStore.getReview(item.mailboxKey, item.messageId, reviewerUsername)
      )
      responseJson(res, 200, {
        sessionId: '',
        messageId: item.messageId,
        review
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.patch(API_ROUTES.itemReview, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const body = (req.body || {}) as ReviewPatchBody
      if (body.flagged === undefined && body.tags === undefined) {
        throw createAppError(400, 'Provide flagged or tags to update review state')
      }

      const reviewerUsername = getReviewOwnerUsername(authSession)
      const review = await reviewStore.upsertReview({
        ...buildReviewContextFromSearchIndexDocument(item, reviewerUsername),
        flagged: body.flagged,
        tags: body.tags
      })
      await searchIndexStore.updateReviewState(item.mailboxKey, item.messageId, reviewerUsername, review)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.item.review.update',
        target: item.subject || item.archiveEntryName || item.messageId,
        outcome: 'success',
        metadata: {
          sourceType: item.sourceType,
          scopePath: item.scopePath,
          archivePath: item.archivePath || '',
          reviewerUsername,
          flagged: review ? review.flagged : false,
          tagCount: review ? review.tags.length : 0
        }
      })

      responseJson(res, 200, {
        sessionId: '',
        messageId: item.messageId,
        review: normalizeReviewState(review)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.itemReview, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const item = await searchIndexStore.findDocumentById(req.params.itemId)
      if (!item) {
        responseJson(res, 404, { error: 'Item not found' })
        return
      }
      if (!isScopePathAllowed(item.scopePath, allowedCasePaths, allowAllCases)) {
        responseJson(res, 403, { error: 'Case access required' })
        return
      }

      const reviewerUsername = getReviewOwnerUsername(authSession)
      await reviewStore.deleteReview(item.mailboxKey, item.messageId, reviewerUsername)
      await searchIndexStore.updateReviewState(item.mailboxKey, item.messageId, reviewerUsername, null)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'search.item.review.delete',
        target: item.subject || item.archiveEntryName || item.messageId,
        outcome: 'success',
        metadata: {
          sourceType: item.sourceType,
          scopePath: item.scopePath,
          archivePath: item.archivePath || '',
          reviewerUsername
        }
      })

      responseJson(res, 200, {
        sessionId: '',
        messageId: item.messageId,
        review: normalizeReviewState(null)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.loadedItemsCsv, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const itemCount = await streamLoadedItemsCsv(res, authSession)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'items.export.csv',
        target: 'Indexed items',
        outcome: 'success',
        metadata: {
          scopePath: '',
          sourceType: 'all',
          itemCount
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.allItemsCsv, async (req, res) => {
    try {
      res.set('Cache-Control', 'no-store')
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const loadedItems = await collectLoadedReviewableItems(
        req.query as WorkspaceItemsRequestQuery,
        authSession
      )
      const csv = buildLoadedItemsCsv(loadedItems.items)
      res
        .status(200)
        .type('text/csv; charset=utf-8')
        .set('Content-Disposition', `attachment; filename="${buildAllItemsCsvFileName(loadedItems.sourceType)}"`)
        .send(csv)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'items.export.csv',
        target: loadedItems.scopeLabel,
        outcome: 'success',
        metadata: {
          scopePath: loadedItems.scopePath,
          sourceType: loadedItems.sourceType,
          itemCount: loadedItems.items.length
        }
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.reviewClearFlags, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const reviewerUsername = getReviewOwnerUsername(authSession)
      const loadedItems = await collectLoadedReviewableItems(
        req.query as WorkspaceItemsRequestQuery,
        authSession
      )
      let clearedCount = 0

      for (const item of loadedItems.items) {
        if (!item.review.flagged) {
          continue
        }

        clearedCount += 1
        const reviewContext = buildReviewPatchInputFromLoadedItem(item, reviewerUsername)
        if (item.review.tags.length) {
          const review = await reviewStore.upsertReview({
            ...reviewContext,
            flagged: false,
            tags: [...item.review.tags]
          })
          await searchIndexStore.updateReviewState(item.mailboxKey, item.messageId, reviewerUsername, review)
        } else {
          await reviewStore.deleteReview(item.mailboxKey, item.messageId, reviewerUsername)
          await searchIndexStore.updateReviewState(item.mailboxKey, item.messageId, reviewerUsername, null)
        }
      }

      recordAuditEvent({
        req,
        session: authSession,
        action: 'review.clear.flags',
        target: loadedItems.scopeLabel,
        outcome: 'success',
        metadata: {
          workspaceMode:
            normalizeText((req.query as WorkspaceItemsRequestQuery).workspaceMode).toLowerCase() === 'search'
              ? 'search'
              : 'folder',
          scopePath: loadedItems.scopePath,
          sourceType: loadedItems.sourceType,
          itemCount: loadedItems.items.length,
          clearedCount
        }
      })

      responseJson(res, 200, {
        clearedCount,
        itemCount: loadedItems.items.length,
        scopePath: loadedItems.scopePath,
        scopeLabel: loadedItems.scopeLabel
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.post(API_ROUTES.searchIndexRefresh, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const source = parseRefreshSource(req.query.source)
      if (authConfig.enabled && !isAdminAuthSession(authSession, authConfig)) {
        recordAuditEvent({
          req,
          session: authSession,
          action: `search.index.refresh.${source}`,
          target: source === 'items' ? 'Items index' : 'Mailbox index',
          outcome: 'denied',
          metadata: {
            reason: 'Admin access required',
            source
          }
        })
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      const status = await searchIndexRefreshCoordinator.start(source, 'manual')
      recordAuditEvent({
        req,
        session: authSession,
        action: `search.index.refresh.${source}`,
        target: source === 'items' ? 'Items index' : 'Mailbox index',
        outcome: 'success',
        metadata: {
          jobId: status.jobId,
          trigger: status.trigger,
          status: status.status,
          source
        }
      })
      responseJson(res, 202, {
        status
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.searchIndexRefreshStatus, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      const source = parseRefreshSource(req.query.source)
      if (authConfig.enabled && !isAdminAuthSession(authSession, authConfig)) {
        responseJson(res, 403, {
          error: 'Admin access required'
        })
        return
      }

      responseJson(res, 200, {
        status: searchIndexRefreshCoordinator.getStatus(source)
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
      const detailByMessageId = new Map<string, MessageDetail>()
      if (!isSummaryOnlyExtraction(fields)) {
        await Promise.all(
          page.items.map(async (item) => {
            detailByMessageId.set(item.id, await resolveMailboxMessageDetail(session, item.id))
          })
        )
      }
      const extraction = buildFolderExtractionPage(
        page,
        new Map(page.items.map((item) => [item.id, item.review])),
        fields,
        (summary, review, fieldList) => {
          if (isSummaryOnlyExtraction(fieldList)) {
            return buildSummaryExtractionRecord(summary, review, fieldList)
          }
          const detail = detailByMessageId.get(summary.id) || buildMessageDetailFromSession(session.index, summary.id, 1)
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

  class FlaggedBundlePartBuilder {
    private readonly entries: FlaggedBundleZipEntry[] = []
    private readonly artifacts: FlaggedBundleExportArtifactRecord[] = []
    private currentEstimatedSize = 22
    private currentOverLimit = false
    private partNumber = 0
    private successfulCount = 0
    private failedCount = 0

    constructor(
      private readonly groupType: 'mailbox' | 'archive',
      private readonly exportRootDir: string,
      private readonly maxBytes: number,
      private readonly onArtifact?: (
        artifact: FlaggedBundleExportArtifactRecord,
        filePath: string
      ) => Promise<void>
    ) {}

    get itemCount(): number {
      return this.successfulCount
    }

    get failureCount(): number {
      return this.failedCount
    }

    recordFailure(count = 1): void {
      this.failedCount += Math.max(0, Math.floor(count))
    }

    async add(entryName: string, content: string | Buffer, mtime: Date): Promise<void> {
      const buffer = toFlaggedBundleBuffer(content)
      const entrySize = estimateZipEntrySize(entryName, buffer.length)
      if (this.entries.length && this.currentEstimatedSize + entrySize > this.maxBytes) {
        await this.flush()
      }

      const exceedsMaxSize = this.entries.length === 0 && this.currentEstimatedSize + entrySize > this.maxBytes
      this.currentOverLimit = this.currentOverLimit || exceedsMaxSize
      this.entries.push({ entryName, content: buffer, mtime })
      this.currentEstimatedSize += entrySize
      this.successfulCount += 1
    }

    async flush(): Promise<void> {
      if (!this.entries.length) {
        return
      }

      this.partNumber += 1
      const artifactId = buildFlaggedBundleArtifactId(this.groupType, this.partNumber)
      const fileName = buildFlaggedBundleArtifactFileName(this.groupType, this.partNumber)
      const filePath = path.join(this.exportRootDir, fileName)
      const sizeBytes = await writeFlaggedBundleZipFile(filePath, this.entries)
      const artifact = {
        artifactId,
        fileName,
        filePath,
        partNumber: this.partNumber,
        partCount: 0,
        itemCount: this.entries.length,
        sizeBytes,
        exceedsMaxSize: this.currentOverLimit
      }
      this.artifacts.push(artifact)
      if (this.onArtifact) {
        await this.onArtifact(artifact, filePath)
      }
      this.entries.length = 0
      this.currentEstimatedSize = 22
      this.currentOverLimit = false
    }

    async finish(): Promise<FlaggedBundleExportArtifactRecord[]> {
      await this.flush()
      const totalParts = this.artifacts.length
      return this.artifacts.map((artifact) => ({
        ...artifact,
        partCount: totalParts
      }))
    }
  }

  function buildFlaggedBundleExportJobResponse(
    job: FlaggedBundleJobRecord
  ): FlaggedBundleExportJobResponse {
    return {
      exportId: job.exportId,
      ownerUsername: job.ownerUsername,
      workspaceKey: job.workspaceKey,
      generatedAt: job.generatedAt,
      startedAt: job.startedAt,
      completedAt: job.completedAt,
      updatedAt: job.updatedAt,
      status: job.status,
      scope: { ...job.scope },
      maxSizeBytes: job.maxSizeBytes,
      progress: { ...job.progress },
      error: job.error,
      groups: job.groups.map((group) => ({
        groupType: group.groupType,
        label: group.label,
        itemCount: group.itemCount,
        failedCount: group.failedCount,
        artifactCount: group.artifacts.length,
        artifacts: group.artifacts.map((artifact) => ({
          artifactId: artifact.artifactId,
          fileName: artifact.fileName,
          downloadUrl: buildFlaggedBundleExportDownloadUrl(job.exportId, artifact.artifactId),
          partNumber: artifact.partNumber,
          partCount: artifact.partCount,
          itemCount: artifact.itemCount,
          sizeBytes: artifact.sizeBytes,
          exceedsMaxSize: artifact.exceedsMaxSize
        }))
      }))
    }
  }

  function buildFlaggedBundleExportHistoryResponse(
    scope: FlaggedBundleExportScope,
    workspaceKey: string,
    jobs: FlaggedBundleJobRecord[]
  ): FlaggedBundleExportHistoryResponse {
    return {
      scope,
      workspaceKey,
      jobs: jobs.map(buildFlaggedBundleExportJobResponse)
    }
  }

  function createFlaggedBundleJobRecord(input: {
    ownerUsername: string
    workspaceKey: string
    generatedAt: string
    scope: FlaggedBundleExportScope
    maxSizeBytes: number
  }): FlaggedBundleJobRecord {
    const now = input.generatedAt
    return {
      exportId: randomBytes(12).toString('hex'),
      ownerUsername: normalizeText(input.ownerUsername) || 'anonymous',
      workspaceKey: normalizeText(input.workspaceKey),
      workspaceLockKey: buildFlaggedBundleWorkspaceKey(
        input.scope.scope,
        input.scope.scopePath,
        input.scope.sessionId
      ),
      generatedAt: now,
      startedAt: now,
      completedAt: null,
      updatedAt: now,
      status: 'running',
      scope: { ...input.scope },
      maxSizeBytes: Math.max(1, Math.floor(input.maxSizeBytes)),
      progress: {
        stage: 'collecting',
        totalItems: 0,
        processedItems: 0,
        failedItems: 0,
        percent: 0,
        currentGroup: null,
        currentLabel: 'Gathering flagged items'
      },
      error: null,
      groups: [
        {
          groupType: 'mailbox',
          label: buildFlaggedBundleGroupLabel('mailbox'),
          itemCount: 0,
          failedCount: 0,
          artifacts: []
        },
        {
          groupType: 'archive',
          label: buildFlaggedBundleGroupLabel('archive'),
          itemCount: 0,
          failedCount: 0,
          artifacts: []
        }
      ]
    }
  }

  async function resolveFlaggedBundleScopeDetails(
    authSession: AuthSessionRecord | null,
    scope: FlaggedBundleScope,
    requestedScopePath: string,
    sessionId: string
  ): Promise<{
    currentUser: AuthUserListItem | null
    allowAllCases: boolean
    allowedCasePaths: string[]
    reviewerUsername: string
    scopeDetails: FlaggedBundleExportScope
    workspaceKey: string
  }> {
    const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
    const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
    const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
    const reviewerUsername = getReviewOwnerUsername(authSession)

    const scopeDetails =
      scope === 'pst'
        ? (() => {
            if (!sessionId) {
              throw createAppError(400, 'Session id is required for selected PST exports')
            }
            const session = getSessionOrThrow(sessions, sessionId)
            if (!isScopePathAllowed(session.scopePath, allowedCasePaths, allowAllCases)) {
              throw createAppError(403, 'Case access required')
            }
            return {
              scopePath: session.scopePath,
              scopeLabel: session.scopeLabel || getScopeLabel(session.scopePath),
              sessionId: session.id,
              sessionFileName: session.fileName,
              scope
            } satisfies FlaggedBundleExportScope
          })()
        : scope === 'search'
          ? (() => {
              const archiveScope = resolveAccessibleArchiveCatalogSelection(
                pstRootDir,
                requestedScopePath,
                allowedCasePaths,
                allowAllCases
              )
              return {
                scopePath: normalizeText(archiveScope.scopePath || requestedScopePath),
                scopeLabel: archiveScope.scopeLabel || getScopeLabel(requestedScopePath),
                sessionId: '',
                sessionFileName: '',
                scope
              } satisfies FlaggedBundleExportScope
            })()
          : {
              scopePath: '',
              scopeLabel: 'All cases/searches',
              sessionId: '',
              sessionFileName: '',
              scope
            }

    return {
      currentUser,
      allowAllCases,
      allowedCasePaths,
      reviewerUsername,
      scopeDetails,
      workspaceKey: buildFlaggedBundleWorkspaceKey(
        scopeDetails.scope,
        scopeDetails.scopePath,
        scopeDetails.sessionId
      )
    }
  }

  async function createFlaggedBundleExportJob(
    authSession: AuthSessionRecord | null,
    body: FlaggedBundlePrepareBody,
    notificationOrigin = ''
  ): Promise<FlaggedBundleExportJobResponse> {
    const scope = parseFlaggedBundleScope(body.scope)
    const requestedScopePath = normalizeScopePath(body.scopePath)
    const sessionId = normalizeText(body.sessionId)
    const maxSizeBytes = Math.max(1, Math.floor(Number(body.maxSizeBytes) || 250 * 1024 * 1024))
    const generatedAt = new Date().toISOString()
    const {
      reviewerUsername,
      allowAllCases,
      allowedCasePaths,
      scopeDetails,
      workspaceKey
    } = await resolveFlaggedBundleScopeDetails(authSession, scope, requestedScopePath, sessionId)
    const job = await flaggedBundleStore.createJob(
      createFlaggedBundleJobRecord({
        ownerUsername: reviewerUsername,
        workspaceKey,
        generatedAt,
        scope: scopeDetails,
        maxSizeBytes
      })
    )
    const exportId = job.exportId
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), `pst-flagged-bundle-${exportId}-`))
    const requestOrigin = normalizeOrigin(notificationOrigin)

    void (async () => {
      let currentJob: FlaggedBundleJobRecord = job
      let mailboxBuilder: FlaggedBundlePartBuilder | null = null
      let archiveBuilder: FlaggedBundlePartBuilder | null = null
      let processedItems = 0

      async function persistCurrentJob(): Promise<void> {
        currentJob.updatedAt = new Date().toISOString()
        const saved = await flaggedBundleStore.saveJob(currentJob)
        if (saved) {
          currentJob = saved
        }
      }

      async function updateJobProgress(patch: Partial<FlaggedBundleProgress>): Promise<void> {
        const totalItems = Math.max(0, Math.floor(patch.totalItems ?? currentJob.progress.totalItems))
        const nextProcessedItems = Math.max(
          0,
          Math.floor(patch.processedItems ?? currentJob.progress.processedItems)
        )
        const nextFailedItems = Math.max(
          0,
          Math.floor(patch.failedItems ?? currentJob.progress.failedItems)
        )
        const nextStage = patch.stage ?? currentJob.progress.stage
        const nextCurrentGroup =
          patch.currentGroup === undefined ? currentJob.progress.currentGroup : patch.currentGroup
        const nextCurrentLabel =
          patch.currentLabel === undefined ? currentJob.progress.currentLabel : normalizeText(patch.currentLabel)
        currentJob.progress = {
          ...currentJob.progress,
          stage: nextStage,
          totalItems,
          processedItems: nextProcessedItems,
          failedItems: nextFailedItems,
          percent:
            totalItems > 0
              ? Math.min(100, Math.floor((nextProcessedItems / totalItems) * 100))
              : nextStage === 'succeeded'
                ? 100
                : 0,
          currentGroup: nextCurrentGroup,
          currentLabel: nextCurrentLabel
        }
        await persistCurrentJob()
      }

      async function persistArtifact(
        groupType: FlaggedBundleGroupType,
        artifact: FlaggedBundleExportArtifactRecord,
        filePath: string
      ): Promise<void> {
        const buffer = await fs.promises.readFile(filePath)
        const saved = await flaggedBundleStore.addArtifact(exportId, groupType, { ...artifact, fileId: '' }, buffer)
        if (!saved) {
          throw createAppError(500, 'Unable to persist flagged bundle artifact')
        }
        currentJob = saved
        try {
          await fs.promises.rm(filePath, { force: true })
        } catch {
          // Ignore temp file cleanup failures.
        }
      }

      async function markGroupComplete(
        groupType: FlaggedBundleGroupType,
        finishedArtifacts: FlaggedBundleExportArtifactRecord[],
        itemCount: number,
        failedCount: number
      ): Promise<void> {
        const group = currentJob.groups.find((entry) => entry.groupType === groupType)
        if (!group) {
          return
        }
        group.itemCount = itemCount
        group.failedCount = failedCount
        const partCountById = new Map(finishedArtifacts.map((artifact) => [artifact.artifactId, artifact.partCount]))
        for (const artifact of group.artifacts) {
          const partCount = partCountById.get(artifact.artifactId)
          if (typeof partCount === 'number') {
            artifact.partCount = partCount
          }
        }
        await persistCurrentJob()
      }

      async function sendReadyEmail(): Promise<void> {
        const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
        const recipientEmail = normalizeText(currentUser?.recipientEmail || '')
        if (!recipientEmail) {
          return
        }

        const settings = await appSettingsStore.getSmtpSettings()
        if (!settings.enabled || !normalizeText(settings.host) || !normalizeText(settings.fromAddress)) {
          return
        }

        const transport = smtpTransportFactory(settings)
        try {
          const response = buildFlaggedBundleExportJobResponse(currentJob)
          const downloadPrefix = requestOrigin || normalizeOrigin(authConfig.publicBaseUrl) || ''
          const downloadLines: string[] = []
          for (const group of response.groups) {
            for (const artifact of group.artifacts) {
              const url = downloadPrefix ? `${downloadPrefix}${artifact.downloadUrl}` : artifact.downloadUrl
              downloadLines.push(`${group.label} - ${artifact.fileName}: ${url}`)
            }
          }

          const scopeLabel = response.scope.scopeLabel || 'Flagged bundle'
          const scopeDescription =
            response.scope.scope === 'pst'
              ? response.scope.sessionFileName || response.scope.scopeLabel
              : response.scope.scopePath || response.scope.scopeLabel

          const text = [
            `Your flagged bundle for ${scopeLabel} is ready.`,
            `Scope: ${scopeDescription}`,
            '',
            ...downloadLines
          ].join('\n')

          const html = [
            '<!doctype html>',
            '<html>',
            '<body style="margin:0;padding:0;background:#f4f7fb;color:#1f2937;font-family:Arial,Helvetica,sans-serif;">',
            '<div style="max-width:720px;margin:0 auto;padding:32px;">',
            '<div style="background:#ffffff;border:1px solid #d7e0ee;border-radius:20px;padding:28px;">',
            `<p style="margin:0 0 12px;font-size:16px;line-height:24px;">Your flagged bundle for <strong>${escapeHtml(
              scopeLabel
            )}</strong> is ready.</p>`,
            `<p style="margin:0 0 18px;font-size:14px;line-height:22px;color:#4b5563;">Scope: ${escapeHtml(
              scopeDescription
            )}</p>`,
            '<ul style="margin:0;padding-left:20px;font-size:14px;line-height:22px;color:#1f2937;">',
            ...downloadLines.map((line) => {
              const [label, url] = line.split(': ', 2)
              return `<li style="margin:0 0 8px;"><strong>${escapeHtml(label)}</strong><br><a href="${escapeHtml(
                url || ''
              )}" style="color:#2f6feb;text-decoration:underline;">${escapeHtml(url || '')}</a></li>`
            }),
            '</ul>',
            '</div>',
            '</div>',
            '</body>',
            '</html>'
          ].join('')

          await transport.sendMail({
            from: buildSmtpFromAddress(settings),
            to: recipientEmail,
            subject: `Flagged bundle ready for ${scopeLabel}`,
            text,
            html
          })
        } catch {
          // Best-effort notification only.
        } finally {
          try {
            await transport.close?.()
          } catch {
            // Ignore transport shutdown failures.
          }
        }
      }

      try {
        const scopeSession = scope === 'pst' ? getSessionOrThrow(sessions, sessionId) : null
        const mailboxes = buildBundleMailboxes(
          pstRootDir,
          scope,
          currentJob.scope.scopePath,
          scopeSession,
          allowedCasePaths,
          allowAllCases
        )
        const archiveBundles = buildBundleArchiveDescriptors(
          pstRootDir,
          scope,
          currentJob.scope.scopePath,
          allowedCasePaths,
          allowAllCases
        )

        const mailboxWork: Array<{ mailbox: BundleMailboxDescriptor; reviews: ReviewRecord[] }> = []
        const archiveWork: Array<{ bundle: ArchiveBundleDescriptor; reviews: ReviewRecord[] }> = []

        for (const mailbox of mailboxes) {
          const flaggedReviews = await reviewStore.listReviews(mailbox.mailboxKey, {
            flaggedOnly: true,
            reviewerUsername
          })
          if (flaggedReviews.length) {
            mailboxWork.push({ mailbox, reviews: flaggedReviews })
          }
        }

        for (const bundle of archiveBundles) {
          const flaggedReviews = await reviewStore.listReviews(bundle.bundlePath, {
            flaggedOnly: true,
            reviewerUsername
          })
          if (flaggedReviews.length) {
            archiveWork.push({ bundle, reviews: flaggedReviews })
          }
        }

        const totalItems = mailboxWork.reduce((count, entry) => count + entry.reviews.length, 0) +
          archiveWork.reduce((count, entry) => count + entry.reviews.length, 0)

        mailboxBuilder = new FlaggedBundlePartBuilder(
          'mailbox',
          rootDir,
          maxSizeBytes,
          async (artifact, filePath) => persistArtifact('mailbox', artifact, filePath)
        )
        archiveBuilder = new FlaggedBundlePartBuilder(
          'archive',
          rootDir,
          maxSizeBytes,
          async (artifact, filePath) => persistArtifact('archive', artifact, filePath)
        )

        await updateJobProgress({
          stage: 'collecting',
          totalItems,
          processedItems: 0,
          failedItems: 0,
          currentGroup: null,
          currentLabel: totalItems > 0 ? 'Gathering flagged items' : 'No flagged items found'
        })

        if (totalItems > 0) {
          await updateJobProgress({
            stage: 'mailbox',
            currentGroup: 'mailbox',
            currentLabel: buildFlaggedBundleGroupLabel('mailbox')
          })

          for (const entry of mailboxWork) {
            let mailboxSession = entry.mailbox.session || null
            if (!mailboxSession) {
              try {
                mailboxSession = openPstMailbox(pstRootDir, entry.mailbox.scopePath, entry.mailbox.fileName)
              } catch {
                mailboxBuilder.recordFailure(entry.reviews.length)
                processedItems += entry.reviews.length
                await updateJobProgress({
                  processedItems,
                  failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                  currentLabel: entry.mailbox.scopeLabel
                })
                continue
              }
            }

            if (!mailboxSession) {
              mailboxBuilder.recordFailure(entry.reviews.length)
              processedItems += entry.reviews.length
              await updateJobProgress({
                processedItems,
                failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                currentLabel: entry.mailbox.scopeLabel
              })
              continue
            }

            for (const review of entry.reviews) {
              const summary = mailboxSession.messages.get(review.messageId) || null
              if (!summary) {
                mailboxBuilder.recordFailure()
                processedItems += 1
                await updateJobProgress({
                  processedItems,
                  failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                  currentLabel: review.subject || review.messageId
                })
                continue
              }

              const payload =
                review.kind === 'appointment'
                  ? exportAppointmentAsIcsFromSession(mailboxSession, review.messageId)
                  : exportMessageAsEmlFromSession(mailboxSession, review.messageId)
              const mtime = summary.modificationTime
                ? new Date(summary.modificationTime)
                : summary.creationTime
                  ? new Date(summary.creationTime)
                  : new Date()
              await mailboxBuilder.add(
                buildBundleEntryPath(
                  review.kind,
                  entry.mailbox.scopePath,
                  entry.mailbox.fileName,
                  review.folderPath,
                  review.subject,
                  review.messageId
                ),
                payload,
                mtime
              )
              processedItems += 1
              await updateJobProgress({
                processedItems,
                failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                currentLabel: review.subject || review.messageId
              })
            }
          }

          const mailboxArtifacts = await mailboxBuilder.finish()
          await markGroupComplete('mailbox', mailboxArtifacts, mailboxBuilder.itemCount, mailboxBuilder.failureCount)

          await updateJobProgress({
            stage: 'archive',
            currentGroup: 'archive',
            currentLabel: buildFlaggedBundleGroupLabel('archive')
          })

          for (const entry of archiveWork) {
            for (const review of entry.reviews) {
              try {
                const item = await searchIndexStore.findDocumentById(review.messageId)
                if (!item || item.archivePath !== entry.bundle.bundlePath || !item.archiveEntryChain?.length) {
                  archiveBuilder.recordFailure()
                  processedItems += 1
                  await updateJobProgress({
                    processedItems,
                    failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                    currentLabel: review.subject || review.messageId
                  })
                  continue
                }

                const { buffer } = await readArchiveBundleItemContent(item.archivePath, item.archiveEntryChain)
                const mtime = item.sortDate
                  ? new Date(item.sortDate)
                  : item.modificationTime
                    ? new Date(item.modificationTime)
                    : item.creationTime
                      ? new Date(item.creationTime)
                      : new Date()
                await archiveBuilder.add(
                  buildArchiveBundleEntryPath(
                    entry.bundle.scopePath,
                    entry.bundle.fileName,
                    item.archiveEntryPath || review.folderPath || review.messageId,
                    item.downloadFilename || item.archiveEntryName || review.subject || review.messageId,
                    review.messageId
                  ),
                  buffer,
                  mtime
                )
                processedItems += 1
                await updateJobProgress({
                  processedItems,
                  failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                  currentLabel: review.subject || review.messageId
                })
              } catch (error) {
                archiveBuilder.recordFailure()
                processedItems += 1
                await updateJobProgress({
                  processedItems,
                  failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
                  currentLabel: review.subject || review.messageId
                })
                console.warn(
                  `Unable to add flagged archive item ${review.messageId} from ${entry.bundle.fileName}:`,
                  error instanceof Error ? error.message : error
                )
                continue
              }
            }
          }

          const archiveArtifacts = await archiveBuilder.finish()
          await markGroupComplete('archive', archiveArtifacts, archiveBuilder.itemCount, archiveBuilder.failureCount)
        }

        await updateJobProgress({
          stage: 'finalizing',
          currentGroup: null,
          currentLabel: 'Finalizing flagged bundle',
          processedItems,
          failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount
        })

        currentJob.status = 'succeeded'
        currentJob.completedAt = new Date().toISOString()
        currentJob.workspaceLockKey = null
        await updateJobProgress({
          stage: 'succeeded',
          currentGroup: null,
          currentLabel: 'Flagged bundle ready',
          processedItems,
          failedItems: mailboxBuilder.failureCount + archiveBuilder.failureCount,
          percent: 100
        })
        await persistCurrentJob()
        await sendReadyEmail()
      } catch (error) {
        currentJob.status = 'failed'
        currentJob.completedAt = new Date().toISOString()
        currentJob.workspaceLockKey = null
        currentJob.error = error instanceof Error ? error.message : String(error)
        const failedItems =
          (mailboxBuilder?.failureCount || 0) + (archiveBuilder?.failureCount || 0) ||
          currentJob.progress.failedItems
        await updateJobProgress({
          stage: 'failed',
          currentGroup: null,
          currentLabel: 'Flagged bundle generation failed',
          processedItems,
          failedItems
        })
        await persistCurrentJob()
      } finally {
        try {
          await fs.promises.rm(rootDir, { recursive: true, force: true })
        } catch {
          // ignore cleanup failures
        }
      }
    })().catch((error) => {
      console.warn('Flagged bundle export background task failed:', error)
    })

    return buildFlaggedBundleExportJobResponse(job)
  }

  app.post(API_ROUTES.flaggedBundlePrepare, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const notificationOrigin = normalizeOrigin(authConfig.publicBaseUrl) || getRequestBaseOrigin(req) || canonicalRequestOrigin(req)
      const job = await createFlaggedBundleExportJob(
        authSession,
        req.body as FlaggedBundlePrepareBody,
        notificationOrigin
      )
      recordAuditEvent({
        req,
        session: authSession,
        action: 'bundle.export.prepare',
        target: job.scope.scopeLabel,
        outcome: 'success',
        metadata: {
          scope: job.scope.scope,
          scopePath: job.scope.scopePath,
          scopeLabel: job.scope.scopeLabel,
          maxSizeBytes: job.maxSizeBytes
        }
      })
      responseJson(res, 202, job)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.flaggedBundlePrepare, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const query = req.query as FlaggedBundleQuery
      const scope = parseFlaggedBundleScope(query.scope)
      const requestedScopePath = normalizeScopePath(query.scopePath)
      const sessionId = normalizeText(query.sessionId)
      const { reviewerUsername, workspaceKey, scopeDetails } = await resolveFlaggedBundleScopeDetails(
        authSession,
        scope,
        requestedScopePath,
        sessionId
      )
      const jobs = await flaggedBundleStore.listJobsForWorkspace(reviewerUsername, workspaceKey)
      recordAuditEvent({
        req,
        session: authSession,
        action: 'bundle.export.list',
        target: scopeDetails.scopeLabel,
        outcome: 'success',
        metadata: {
          scope: scopeDetails.scope,
          scopePath: scopeDetails.scopePath,
          scopeLabel: scopeDetails.scopeLabel,
          jobCount: jobs.length
        }
      })
      responseJson(res, 200, buildFlaggedBundleExportHistoryResponse(scopeDetails, workspaceKey, jobs))
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.flaggedBundleJob, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const exportId = normalizeText(req.params.exportId)
      const job = await flaggedBundleStore.getJob(exportId)
      if (!job || job.ownerUsername !== getReviewOwnerUsername(authSession)) {
        responseJson(res, 404, { error: 'Flagged bundle export not found' })
        return
      }
      if (job.status === 'running') {
        responseJson(res, 409, { error: 'Flagged bundle export is still running' })
        return
      }

      const deleted = await flaggedBundleStore.deleteJob(exportId)
      if (!deleted) {
        responseJson(res, 404, { error: 'Flagged bundle export not found' })
        return
      }

      recordAuditEvent({
        req,
        session: authSession,
        action: 'bundle.export.delete',
        target: job.scope.scopeLabel,
        outcome: 'success',
        metadata: {
          scope: job.scope.scope,
          scopePath: job.scope.scopePath,
          scopeLabel: job.scope.scopeLabel,
          exportId
        }
      })
      responseJson(res, 200, { deleted: true, exportId } satisfies FlaggedBundleExportDeleteResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.flaggedBundleArtifact, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }

      const exportId = normalizeText(req.params.exportId)
      const artifactId = normalizeText(req.params.artifactId)
      const job = await flaggedBundleStore.getJob(exportId)
      if (!job || job.ownerUsername !== getReviewOwnerUsername(authSession)) {
        responseJson(res, 404, {
          error: 'Flagged bundle export not found'
        })
        return
      }

      const download = await flaggedBundleStore.openArtifactDownload(exportId, artifactId)
      if (!download) {
        responseJson(res, 404, {
          error: 'Export artifact not found'
        })
        return
      }

      res.status(200)
        .type('application/zip')
        .set('Cache-Control', 'private, no-store, max-age=0')
        .set('Content-Disposition', `attachment; filename="${download.artifact.fileName}"`)

      const stream = download.stream
      stream.on('error', (error) => {
        if (!res.headersSent) {
          createRouteErrorHandler(res, error)
          return
        }
        res.destroy(error instanceof Error ? error : undefined)
      })
      stream.pipe(res)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.flaggedBundleExport, async (req, res) => {
    try {
      const authSession = getAuthSessionFromRequest(req)
      if (authConfig.enabled && !authSession) {
        responseJson(res, 401, {
          error: 'Authentication required'
        })
        return
      }
      const currentUser = authSession ? await authUserStore.getUser(authSession.username) : null
      const allowAllCases = !authConfig.enabled || isAdminAuthSession(authSession, authConfig)
      const allowedCasePaths = allowAllCases ? [] : getAccessibleCasePaths(currentUser)
      const reviewerUsername = getReviewOwnerUsername(authSession)
      const query = req.query as FlaggedBundleQuery
      const scope = parseFlaggedBundleScope(query.scope)
      const scopePath = normalizeScopePath(query.scopePath)
      const sessionId = normalizeText(query.sessionId)
      const session = scope === 'pst' ? getSessionOrThrow(sessions, sessionId) : null
      const bundleScope =
        scope === 'search'
          ? resolveAccessibleArchiveCatalogSelection(
              pstRootDir,
              scopePath,
              allowedCasePaths,
              allowAllCases
            )
          : scope === 'all'
            ? resolveAccessibleCatalogSelection(
                pstRootDir,
                '',
                allowedCasePaths,
                listPstMailboxFiles,
                allowAllCases
              )
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
        session,
        allowedCasePaths,
        allowAllCases
      )
      const archiveBundles = buildBundleArchiveDescriptors(
        pstRootDir,
        scope,
        scope === 'search' ? normalizeText(bundleScope?.scopePath || scopePath) : scopePath,
        allowedCasePaths,
        allowAllCases
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
                sourceType: 'mailbox',
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
              sourceType: 'mailbox',
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
              sourceType: 'mailbox',
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
              sourceType: 'mailbox',
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

      for (const bundle of archiveBundles) {
        const flaggedReviews = await reviewStore.listReviews(bundle.bundlePath, {
          flaggedOnly: true,
          reviewerUsername
        })
        if (!flaggedReviews.length) {
          continue
        }

        for (const review of flaggedReviews) {
          const item = await searchIndexStore.findDocumentById(review.messageId)
          const outputFile = buildArchiveBundleEntryPath(
            bundle.scopePath,
            bundle.fileName,
            item?.archiveEntryPath || review.folderPath || review.messageId,
            item?.downloadFilename || item?.archiveEntryName || review.subject || review.messageId,
            review.messageId
          )

          if (!item || item.archivePath !== bundle.bundlePath || !item.archiveEntryChain?.length) {
            addBundleManifestItem(manifest, {
              sourcePstPath: bundle.bundlePath,
              mailboxName: bundle.fileName,
              mailboxKey: bundle.bundlePath,
              sourceType: 'archive',
              scopePath: bundle.scopePath,
              scopeLabel: bundle.scopeLabel,
              fileName: bundle.fileName,
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
              outputType: 'raw',
              status: 'error',
              error: 'Archive item not found in search index'
            })
            continue
          }

          try {
            const { buffer, contentType, fileName } = await readArchiveBundleItemContent(
              item.archivePath,
              item.archiveEntryChain
            )
            const mtime = item.sortDate
              ? new Date(item.sortDate)
              : item.modificationTime
                ? new Date(item.modificationTime)
                : item.creationTime
                  ? new Date(item.creationTime)
                  : new Date()
            await writer.addFile(outputFile, buffer, { mtime })
            addBundleManifestItem(manifest, {
              sourcePstPath: bundle.bundlePath,
              mailboxName: bundle.fileName,
              mailboxKey: bundle.bundlePath,
              sourceType: 'archive',
              scopePath: bundle.scopePath,
              scopeLabel: bundle.scopeLabel,
              fileName: bundle.fileName,
              folderId: item.folderId,
              folderPath: item.folderPath,
              messageId: item.messageId,
              descriptorId: item.descriptorId,
              kind: item.kind,
              subject: item.subject,
              review: {
                flagged: review.flagged,
                tags: [...review.tags],
                createdAt: review.createdAt,
                updatedAt: review.updatedAt
              },
              outputFile,
              outputType: 'raw',
              status: 'exported',
              archivePath: item.archivePath,
              archiveEntryPath: item.archiveEntryPath,
              archiveEntryName: item.archiveEntryName,
              contentType: item.contentType || contentType,
              downloadFilename: item.downloadFilename || fileName
            })
          } catch (error) {
            addBundleManifestItem(manifest, {
              sourcePstPath: bundle.bundlePath,
              mailboxName: bundle.fileName,
              mailboxKey: bundle.bundlePath,
              sourceType: 'archive',
              scopePath: bundle.scopePath,
              scopeLabel: bundle.scopeLabel,
              fileName: bundle.fileName,
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
              outputType: 'raw',
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
      invalidateMailboxDetailCache(session, messageId, reviewerUsername)
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
      invalidateMailboxDetailCache(session, messageId, reviewerUsername)
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
