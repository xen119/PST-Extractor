export interface AuthUser {
  username: string
  assignedCasePaths: string[]
}

export type SearchSourceType = 'mailbox' | 'teams' | 'sharepoint'

export interface AuthStatus {
  authenticated: boolean
  enabled: boolean
  canManageUsers: boolean
  mfaEnabled: boolean
  mfaEnforced: boolean
  user: AuthUser | null
  expiresAt: string | null
  mfaRequired?: boolean
}

export interface ReviewState {
  flagged: boolean
  tags: string[]
  createdAt: string
  updatedAt: string
}

export interface CatalogEntry {
  fileName: string
  size: number
  modifiedAt: string | null
  scopePath?: string
  displayPath?: string
}

export interface CatalogScope {
  scopePath: string
  scopeLabel: string
  fileCount: number
  files: CatalogEntry[]
}

export interface PstCatalogResponse {
  rootPath: string
  rootExists: boolean
  message: string
  scopePath: string
  scopeLabel: string
  scopes: CatalogScope[]
  files: CatalogEntry[]
}

export interface SessionSummary {
  rootFolderId?: string
  messageCount?: number
  folderCount?: number
  mailCount?: number
  warningCount?: number
  warnings?: string[]
  mailboxName?: string
  fileName?: string
  createdAt?: string
  stats?: Record<string, number>
}

export interface FolderNode {
  id: string
  descriptorId?: string
  displayName: string
  path: string
  indexedMessageCount?: number
  mailMessageCount?: number
  children?: FolderNode[]
}

export interface SessionOpenResponse {
  sessionId: string
  scopePath: string
  scopeLabel: string
  fileName: string
  summary: SessionSummary
  tree: FolderNode
}

export interface MessageSummary {
  id: string
  sourceType?: SearchSourceType
  messageId?: string
  descriptorId?: string
  folderId?: string
  folderPath?: string
  order?: number
  messageClass?: string
  kind?: string
  subject?: string
  senderName?: string
  senderEmailAddress?: string
  recipientText?: string
  displayTo?: string
  displayCC?: string
  displayBCC?: string
  resolvedDisplayTo?: string
  resolvedDisplayCC?: string
  resolvedDisplayBCC?: string
  originalSubject?: string
  clientSubmitTime?: string | null
  creationTime?: string | null
  modificationTime?: string | null
  messageDeliveryTime?: string | null
  sortDate?: string | null
  sortDateMs?: number | null
  importance?: number
  hasAttachments?: boolean
  isRead?: boolean
  isMailLike?: boolean
  review?: ReviewState
  scopePath?: string
  scopeLabel?: string
  fileName?: string
  mailboxName?: string
  parseError?: string
  archivePath?: string
  archiveEntryPath?: string
  archiveEntryChain?: string[]
  archiveEntryName?: string
  contentType?: string
  downloadFilename?: string
  previewKind?: 'text' | 'html' | 'binary'
  previewText?: string
  previewHtml?: string
  previewUrl?: string
  downloadUrl?: string
}

export interface ReviewFilters {
  flaggedOnly: boolean
  taggedOnly: boolean
  tag: string
}

export interface PageResponse<TItem> {
  folder?: {
    id: string
    descriptorId?: string
    displayName: string
    path: string
  }
  items: TItem[]
  total: number
  page: number
  pageSize: number
  totalPages: number
  query: string
  mailOnly: boolean
  sort: string
  sourceType?: SearchSourceType | 'all'
  sourceCounts?: Record<SearchSourceType, number>
  reviewFilters?: ReviewFilters
}

export interface FolderMessagesResponse {
  sessionId: string
  page: PageResponse<MessageSummary>
}

export interface SearchResponse {
  scope: 'all' | 'search' | 'pst'
  scopePath: string
  scopeLabel: string
  sourceType?: SearchSourceType | 'all'
  page: PageResponse<MessageSummary>
}

export interface AttachmentDetail {
  attachmentId: string
  index: number
  filename?: string
  longFilename?: string
  downloadFilename?: string
  mimeTag?: string
  size?: number
  attachMethod?: number
  contentId?: string
  pathname?: string
  longPathname?: string
  isEmbeddedMessage?: boolean
  isDownloadable?: boolean
  downloadUrl?: string
  parseError?: string
}

export interface MessageDetail {
  id?: string
  sourceType?: SearchSourceType
  subject?: string
  senderName?: string
  senderEmailAddress?: string
  displayTo?: string
  displayCC?: string
  displayBCC?: string
  resolvedDisplayTo?: string
  resolvedDisplayCC?: string
  resolvedDisplayBCC?: string
  clientSubmitTime?: string | null
  creationTime?: string | null
  modificationTime?: string | null
  messageDeliveryTime?: string | null
  sortDate?: string | null
  bodyHtml?: string
  bodyText?: string
  bodyPrefix?: string
  parseError?: string
  attachments?: AttachmentDetail[]
  review?: ReviewState
  folderId?: string
  folderPath?: string
  mailboxName?: string
  archivePath?: string
  archiveEntryPath?: string
  archiveEntryChain?: string[]
  archiveEntryName?: string
  contentType?: string
  downloadFilename?: string
  previewKind?: 'text' | 'html' | 'binary'
  previewText?: string
  previewHtml?: string
  downloadUrl?: string
}

export interface MessageDetailResponse {
  sessionId: string
  detail: MessageDetail
}

export interface ReviewUpdateResponse {
  sessionId: string
  messageId: string
  review: ReviewState
}

export interface ReviewQueueRecord {
  mailboxKey?: string
  fileName?: string
  messageId?: string
  descriptorId?: string
  folderId?: string
  folderPath?: string
  messageClass?: string
  kind?: string
  isMailLike?: boolean
  subject?: string
  senderName?: string
  senderEmailAddress?: string
  displayTo?: string
  displayCC?: string
  displayBCC?: string
  resolvedDisplayTo?: string
  resolvedDisplayCC?: string
  resolvedDisplayBCC?: string
  flagged: boolean
  tags: string[]
  createdAt: string
  updatedAt: string
}

export interface HiddenRule {
  filterId: string
  kind: 'address' | 'subject'
  value: string
  label: string
  createdAt: string
  updatedAt: string
}

export interface HiddenRulesResponse {
  items: HiddenRule[]
}

export interface UserInvite {
  username: string
  createdAt: string
  recipientEmail: string
  inviteStatus: 'pending' | 'active' | 'revoked' | 'expired'
  inviteSentAt: string
  inviteExpiresAt: string
  inviteAcceptedAt: string
  inviteRevokedAt: string
  mfaEnabled: boolean
  mfaEnforced: boolean
  mfaEnrolledAt: string
  assignedCasePaths: string[]
  inviteUrl?: string
}

export interface UsersResponse {
  users: UserInvite[]
}

export interface UserInviteResponse {
  user: UserInvite
  inviteUrl: string
  emailSent: boolean
  inviteExpiresAt: string
}

export interface InviteLookupResponse {
  invite: UserInvite
}

export interface InviteAcceptResponse {
  user: UserInvite
  mfaAvailable: boolean
}

export interface MfaEnrollmentStartResponse {
  user: UserInvite
  secret: string
  otpauthUri: string
  qrCodeDataUrl: string
}

export interface MfaEnrollmentCompleteResponse {
  user: UserInvite
  recoveryCodes: string[]
}

export interface SmtpSettings {
  enabled: boolean
  host: string
  port: number
  secure: boolean
  username: string
  hasPassword: boolean
  fromName: string
  fromAddress: string
  replyTo: string
}

export interface SmtpSettingsResponse {
  settings: SmtpSettings
}

export interface SmtpTestResponse {
  success: boolean
  recipient: string
  messageId: string
  accepted: string[]
  rejected: string[]
}

export interface SearchIndexRefreshSummary {
  mailboxCount: number
  messageCount: number
}

export interface SearchIndexRefreshStatus {
  jobId: string | null
  status: 'idle' | 'running' | 'succeeded' | 'failed'
  trigger: 'startup' | 'manual' | null
  startedAt: string | null
  completedAt: string | null
  updatedAt: string
  summary: SearchIndexRefreshSummary | null
  error: string | null
}

export interface SearchIndexRefreshResponse {
  status: SearchIndexRefreshStatus
}

export interface ActivityLogActor {
  username: string
  authenticated: boolean
  admin: boolean
}

export interface ActivityLogRequest {
  method: string
  path: string
  origin: string
  ip: string
}

export interface ActivityLogEntry {
  timestamp: string
  actor: ActivityLogActor
  action: string
  target: string
  outcome: 'success' | 'failure' | 'denied'
  request: ActivityLogRequest
  metadata: Record<string, unknown>
}

export interface ActivityLogResponse {
  entries: ActivityLogEntry[]
}
