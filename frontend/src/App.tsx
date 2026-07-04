import * as React from 'react'
import {
  Activity,
  ChevronLeft,
  ChevronRight,
  Download,
  Flag,
  FolderCog,
  KeyRound,
  Link2,
  LogOut,
  Mail,
  MoonStar,
  Plus,
  RefreshCw,
  RotateCcw,
  ScanLine,
  Search,
  Send,
  Settings,
  ShieldAlert,
  ShieldCheck,
  SunMedium,
  Trash2,
  UserPlus,
  Users,
  X
} from 'lucide-react'
import { api, ApiError } from '@/api'
import {
  AuthScreen,
  MfaReminderDialog,
  MfaSetupDialog,
  type AuthView,
  type InviteStep
} from '@/components/auth'
import { AppShell, EmailPreview, EmptyState, MessageList, Sidebar, TagManagerDialog } from '@/components/layout'
import { Badge, Button, Dialog, DialogContent, DialogDescription, DialogTitle, IconButton, Input, PopoverClose, ScrollArea, Separator } from '@/components/ui'
import { deriveSearchMode, normalizeSearchResultsPage } from '@/lib/search'
import { cn, downloadTextFile, formatDate, normalizeText } from '@/lib/utils'
import {
  getWorkspaceStorageNamespace,
  readWorkspaceStorageBool,
  readWorkspaceStorageItem,
  removeWorkspaceStorageItem,
  writeWorkspaceStorageItem
} from '@/lib/workspace'
import type {
  ActivityLogEntry,
  ActivityLogResponse,
  AuthStatus,
  CatalogEntry,
  CatalogScope,
  FolderNode,
  HiddenRule,
  HiddenRulesResponse,
  InviteAcceptResponse,
  InviteLookupResponse,
  MessageDetail,
  MessageSummary,
  MfaEnrollmentCompleteResponse,
  MfaEnrollmentStartResponse,
  PasswordResetLookupResponse,
  PageResponse,
  PstCatalogResponse,
  SearchSourceType,
  SessionOpenResponse,
  SmtpSettings,
  SmtpSettingsResponse,
  SearchIndexRefreshSource,
  SearchIndexRefreshStatus,
  UserInvite,
  UsersResponse
} from '@/types'
import { useUiStore } from '@/store/ui'

type WorkspaceMode = 'folder' | 'search'
type SearchScope = 'pst' | 'search' | 'all'
type CorpusSourceType = SearchSourceType
type OpenMailboxOptions = {
  preserveWorkspaceMode?: boolean
  selectedMessageId?: string
  preservePreview?: boolean
}
type SmtpFormState = {
  enabled: boolean
  host: string
  port: number
  secure: boolean
  username: string
  fromName: string
  fromAddress: string
  replyTo: string
}

const DEFAULT_SMTP_FORM: SmtpFormState = {
  enabled: false,
  host: '',
  port: 587,
  secure: false,
  username: '',
  fromName: '',
  fromAddress: '',
  replyTo: ''
}

const SEARCH_INDEX_REFRESH_POLL_INTERVAL_MS = 2000
const SEARCH_INDEX_REFRESH_SOURCES: SearchIndexRefreshSource[] = ['mailboxes', 'items']

function getInviteToken(pathname = window.location.pathname): string | null {
  const match = pathname.match(/^\/invite\/([^/?#]+)/i)
  return match ? decodeURIComponent(match[1]) : null
}

function getPasswordResetToken(pathname = window.location.pathname): string | null {
  const match = pathname.match(/^\/reset\/([^/?#]+)/i)
  return match ? decodeURIComponent(match[1]) : null
}

function getReminderStorageKey(username: string): string {
  return `pst-mail-explorer.mfaReminder::${getWorkspaceStorageNamespace(true, username)}`
}

function readReminderDismissed(username: string): boolean {
  return sessionStorage.getItem(getReminderStorageKey(username)) === '1'
}

function writeReminderDismissed(username: string, dismissed: boolean): void {
  const key = getReminderStorageKey(username)
  if (dismissed) {
    sessionStorage.setItem(key, '1')
    return
  }
  sessionStorage.removeItem(key)
}

function triggerDownload(url: string, fileName?: string): void {
  const anchor = document.createElement('a')
  anchor.href = url
  if (fileName) {
    anchor.download = fileName
  }
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
}

export function getCasePathFromScopePath(scopePath: string): string {
  const normalized = normalizeText(scopePath)
  if (!normalized) {
    return ''
  }
  return normalized.split('/').filter(Boolean)[0] || ''
}

function getTopLevelCaseOptions(scopes: CatalogScope[]): Array<{ label: string; value: string; count: number }> {
  const options = new Map<string, { label: string; value: string; count: number }>()
  for (const scope of scopes) {
    const normalizedScopePath = normalizeText(scope.scopePath)
    if (!normalizedScopePath) {
      continue
    }
    const casePath = getCasePathFromScopePath(normalizedScopePath)
    const current = options.get(casePath)
    options.set(casePath, {
      label: casePath,
      value: casePath,
      count: (current?.count || 0) + scope.fileCount
    })
  }
  if (!options.size) {
    const rootScope = scopes.find((scope) => !normalizeText(scope.scopePath))
    if (rootScope) {
      options.set('', {
        label: 'PST root',
        value: '',
        count: rootScope.fileCount
      })
    }
  }
  return Array.from(options.values()).sort((left, right) => {
    if (left.value === right.value) {
      return 0
    }
    return left.label.localeCompare(right.label, undefined, { sensitivity: 'base' })
  })
}

function getSearchOptionsForCase(
  scopes: CatalogScope[],
  casePath: string
): Array<{ label: string; value: string; count: number }> {
  const normalizedCasePath = normalizeText(casePath)
  const options = new Map<string, { label: string; value: string; count: number }>()
  const caseDepth = normalizedCasePath ? normalizedCasePath.split('/').filter(Boolean).length : 0
  const prefix = normalizedCasePath ? `${normalizedCasePath}/` : ''

  for (const scope of scopes) {
    const normalizedScopePath = normalizeText(scope.scopePath)
    if (!normalizedScopePath) {
      if (!normalizedCasePath) {
        const current = options.get('')
        options.set('', {
          label: 'PST root',
          value: '',
          count: (current?.count || 0) + scope.fileCount
        })
      }
      continue
    }

    const scopeDepth = normalizedScopePath.split('/').filter(Boolean).length
    const isImmediateChild = normalizedCasePath
      ? normalizedScopePath.startsWith(prefix) && scopeDepth === caseDepth + 1
      : scopeDepth === 1
    if (!isImmediateChild) {
      continue
    }

    const current = options.get(normalizedScopePath)
    options.set(normalizedScopePath, {
      label: normalizedScopePath.split('/').pop() || scope.scopeLabel || normalizedScopePath,
      value: normalizedScopePath,
      count: (current?.count || 0) + scope.fileCount
    })
  }

  if (!options.size && normalizedCasePath) {
    const selected = scopes.find((scope) => normalizeText(scope.scopePath) === normalizedCasePath)
    if (selected) {
      options.set(selected.scopePath, {
        label: selected.scopePath.split('/').pop() || selected.scopeLabel || selected.scopePath || 'PST root',
        value: selected.scopePath,
        count: selected.fileCount
      })
    }
  }

  return Array.from(options.values()).sort((left, right) => {
    if (left.value === right.value) {
      return 0
    }
    return left.label.localeCompare(right.label, undefined, { sensitivity: 'base' })
  })
}

function getDefaultSearchPathForCase(scopes: CatalogScope[], casePath: string): string {
  return getSearchOptionsForCase(scopes, casePath)[0]?.value || getCasePathFromScopePath(casePath)
}

function getScopeEntryForPath(scopes: CatalogScope[], scopePath: string): CatalogScope | null {
  const normalizedScopePath = normalizeText(scopePath)
  if (!normalizedScopePath) {
    return scopes.find((scope) => !normalizeText(scope.scopePath)) || null
  }

  return scopes.find((scope) => normalizeText(scope.scopePath) === normalizedScopePath) || null
}

function getFolderNode(root: FolderNode | null, folderId: string): FolderNode | null {
  if (!root || !folderId) {
    return null
  }
  if (root.id === folderId) {
    return root
  }
  const children = Array.isArray(root.children) ? root.children : []
  for (const child of children) {
    const match = getFolderNode(child, folderId)
    if (match) {
      return match
    }
  }
  return null
}

function getMessagePreviewTitle(detail: MessageDetail | null): string {
  return detail?.subject || '(no subject)'
}

function isSameMailboxContext(
  message: MessageSummary,
  currentFileName: string,
  currentScopePath: string
): boolean {
  if (message.sourceType && message.sourceType !== 'mailbox') {
    return false
  }
  const nextFileName = normalizeText(message.fileName || '')
  const nextScopePath = normalizeText(message.scopePath || '')
  const normalizedCurrentFileName = normalizeText(currentFileName)
  const normalizedCurrentScopePath = normalizeText(currentScopePath)

  if (!nextFileName || !nextScopePath || !normalizedCurrentFileName || !normalizedCurrentScopePath) {
    return true
  }

  return nextFileName === normalizedCurrentFileName && nextScopePath === normalizedCurrentScopePath
}

export function App() {
  const theme = useUiStore((state) => state.theme)
  const sidebarCollapsed = useUiStore((state) => state.sidebarCollapsed)
  const previewCollapsed = useUiStore((state) => state.previewCollapsed)
  const hiddenFiltersOpen = useUiStore((state) => state.hiddenFiltersOpen)
  const setSidebarCollapsed = useUiStore((state) => state.setSidebarCollapsed)
  const setPreviewCollapsed = useUiStore((state) => state.setPreviewCollapsed)
  const toggleTheme = useUiStore((state) => state.toggleTheme)
  const setHiddenFiltersOpen = useUiStore((state) => state.setHiddenFiltersOpen)

  const [bootLoading, setBootLoading] = React.useState(true)
  const [authReady, setAuthReady] = React.useState(false)
  const [authStatus, setAuthStatus] = React.useState<AuthStatus | null>(null)
  const [authView, setAuthView] = React.useState<AuthView>('login')
  const [authBusy, setAuthBusy] = React.useState(false)
  const [authMessage, setAuthMessage] = React.useState('')
  const [authError, setAuthError] = React.useState('')
  const [mfaChallengeUsername, setMfaChallengeUsername] = React.useState('')
  const [inviteToken, setInviteToken] = React.useState<string | null>(() => getInviteToken())
  const [resetToken, setResetToken] = React.useState<string | null>(() => getPasswordResetToken())
  const [invite, setInvite] = React.useState<UserInvite | null>(null)
  const [inviteLoading, setInviteLoading] = React.useState(false)
  const [inviteStep, setInviteStep] = React.useState<InviteStep>('password')
  const [inviteMfaAvailable, setInviteMfaAvailable] = React.useState(false)
  const [inviteMfaEnforced, setInviteMfaEnforced] = React.useState(false)
  const [inviteSetup, setInviteSetup] = React.useState<MfaEnrollmentStartResponse | null>(null)
  const [inviteRecoveryCodes, setInviteRecoveryCodes] = React.useState<string[]>([])
  const [passwordResetAvailable, setPasswordResetAvailable] = React.useState(false)
  const [resetLookup, setResetLookup] = React.useState<PasswordResetLookupResponse | null>(null)
  const [selfMfaOpen, setSelfMfaOpen] = React.useState(false)
  const [selfMfaLoading, setSelfMfaLoading] = React.useState(false)
  const [selfMfaMessage, setSelfMfaMessage] = React.useState('')
  const [selfMfaError, setSelfMfaError] = React.useState('')
  const [selfMfaSetup, setSelfMfaSetup] = React.useState<MfaEnrollmentStartResponse | null>(null)
  const [selfMfaRecoveryCodes, setSelfMfaRecoveryCodes] = React.useState<string[]>([])

  const [catalog, setCatalog] = React.useState<PstCatalogResponse | null>(null)
  const [catalogMessage, setCatalogMessage] = React.useState('')
  const [catalogFiles, setCatalogFiles] = React.useState<CatalogEntry[]>([])
  const [caseOptions, setCaseOptions] = React.useState<CatalogScope[]>([])
  const [selectedCasePath, setSelectedCasePath] = React.useState('')
  const [selectedScopePath, setSelectedScopePath] = React.useState('')
  const [mailboxSearchScopePath, setMailboxSearchScopePath] = React.useState('')
  const [selectedPstFileName, setSelectedPstFileName] = React.useState('')
  const [sessionId, setSessionId] = React.useState<string | null>(null)
  const [sessionSummary, setSessionSummary] = React.useState<SessionOpenResponse['summary'] | null>(null)
  const [folderTree, setFolderTree] = React.useState<FolderNode | null>(null)
  const [currentFolderId, setCurrentFolderId] = React.useState('')
  const [currentPage, setCurrentPage] = React.useState<PageResponse<MessageSummary> | null>(null)
  const [pageIndex, setPageIndex] = React.useState(1)
  const [selectedMessageId, setSelectedMessageId] = React.useState('')
  const [selectedMessage, setSelectedMessage] = React.useState<MessageDetail | null>(null)
  const [messagesLoading, setMessagesLoading] = React.useState(false)
  const [messageLoading, setMessageLoading] = React.useState(false)
  const [workspaceMode, setWorkspaceMode] = React.useState<WorkspaceMode>('folder')
  const [searchQuery, setSearchQuery] = React.useState('')
  const [searchInputQuery, setSearchInputQuery] = React.useState('')
  const [searchScope, setSearchScope] = React.useState<SearchScope>('all')
  const [sourceType, setSourceType] = React.useState<CorpusSourceType>('mailbox')
  const [mailOnly, setMailOnly] = React.useState(false)
  const [sort, setSort] = React.useState('date-desc')
  const [reviewFlaggedOnly, setReviewFlaggedOnly] = React.useState(false)
  const [reviewTaggedOnly, setReviewTaggedOnly] = React.useState(false)
  const [hiddenRules, setHiddenRules] = React.useState<HiddenRule[]>([])

  const [usersDialogOpen, setUsersDialogOpen] = React.useState(false)
  const [usersLoading, setUsersLoading] = React.useState(false)
  const [usersError, setUsersError] = React.useState('')
  const [usersMessage, setUsersMessage] = React.useState('')
  const [users, setUsers] = React.useState<UserInvite[]>([])
  const [inviteUsername, setInviteUsername] = React.useState('')
  const [inviteEmail, setInviteEmail] = React.useState('')
  const [selectedAdminUser, setSelectedAdminUser] = React.useState('')
  const [caseAccessDialogUser, setCaseAccessDialogUser] = React.useState('')
  const [caseAccessDraftPaths, setCaseAccessDraftPaths] = React.useState<string[]>([])
  const [userActivity, setUserActivity] = React.useState<ActivityLogEntry[]>([])
  const [userActivityLoading, setUserActivityLoading] = React.useState(false)
  const [userActivityError, setUserActivityError] = React.useState('')

  const [smtpDialogOpen, setSmtpDialogOpen] = React.useState(false)
  const [smtpLoading, setSmtpLoading] = React.useState(false)
  const [smtpError, setSmtpError] = React.useState('')
  const [smtpMessage, setSmtpMessage] = React.useState('')
  const [smtpForm, setSmtpForm] = React.useState<SmtpFormState>(DEFAULT_SMTP_FORM)
  const [smtpPassword, setSmtpPassword] = React.useState('')
  const [smtpTestRecipient, setSmtpTestRecipient] = React.useState('')

  const [activityDialogOpen, setActivityDialogOpen] = React.useState(false)
  const [activityLoading, setActivityLoading] = React.useState(false)
  const [activityError, setActivityError] = React.useState('')
  const [activityMessage, setActivityMessage] = React.useState('')
  const [activityEntries, setActivityEntries] = React.useState<ActivityLogEntry[]>([])
  const [activityFilterUser, setActivityFilterUser] = React.useState('')
  const [clearFlagsDialogOpen, setClearFlagsDialogOpen] = React.useState(false)
  const [clearFlagsLoading, setClearFlagsLoading] = React.useState(false)
  const [clearFlagsError, setClearFlagsError] = React.useState('')
  const [searchIndexRefreshStatuses, setSearchIndexRefreshStatuses] = React.useState<
    Record<SearchIndexRefreshSource, SearchIndexRefreshStatus | null>
  >({
    mailboxes: null,
    items: null
  })
  const [searchIndexRefreshActionBusyBySource, setSearchIndexRefreshActionBusyBySource] = React.useState<
    Record<SearchIndexRefreshSource, boolean>
  >({
    mailboxes: false,
    items: false
  })
  const [mfaReminderDismissed, setMfaReminderDismissed] = React.useState(false)
  const [tagsDialogOpen, setTagsDialogOpen] = React.useState(false)
  const [fullViewOpen, setFullViewOpen] = React.useState(false)
  const refreshCurrentPageRef = React.useRef<() => Promise<void>>(async () => undefined)
  const searchIndexRefreshPollTimeoutRef = React.useRef<Record<SearchIndexRefreshSource, number | null>>({
    mailboxes: null,
    items: null
  })
  const searchIndexRefreshPollInFlightRef = React.useRef<Record<SearchIndexRefreshSource, boolean>>({
    mailboxes: false,
    items: false
  })
  const searchIndexRefreshPollJobIdRef = React.useRef<Record<SearchIndexRefreshSource, string | null>>({
    mailboxes: null,
    items: null
  })
  const messagePreviewRequestRef = React.useRef(0)
  const skipNextMessageReloadRef = React.useRef(false)
  const skipNextMailboxOpenRef = React.useRef('')
  const previewRequestKeyRef = React.useRef('')
  const mailboxDetailCacheRef = React.useRef(new Map<string, MessageDetail>())
  const mailboxPreviewCacheRef = React.useRef(new Map<string, MessageDetail>())
  const mailboxDetailInFlightRef = React.useRef(new Map<string, Promise<MessageDetail>>())

  const catalogLoadKeyRef = React.useRef('')

  const username = authStatus?.user?.username || ''
  const authenticated = Boolean(authStatus?.authenticated)
  const canManageUsers = Boolean(authStatus?.canManageUsers)
  const mfaEnabled = Boolean(authStatus?.mfaEnabled)
  const inviteFlowActive = Boolean(inviteToken)
  const resetFlowActive = Boolean(resetToken)
  const authFlowActive = inviteFlowActive || resetFlowActive
  const mfaEnforced = Boolean(authStatus?.mfaEnforced)
  const assignedCasePathsKey = (authStatus?.user?.assignedCasePaths || []).join('|')
  const selectedPageItem = React.useMemo(
    () => currentPage?.items?.find((item) => item.id === selectedMessageId) || null,
    [currentPage, selectedMessageId]
  )
  const showReminder =
    authenticated &&
    !mfaEnabled &&
    !selfMfaOpen &&
    !authFlowActive &&
    (mfaEnforced || !mfaReminderDismissed)
  const workspaceReady =
    authenticated && !authFlowActive && (mfaEnabled || (!mfaEnforced && mfaReminderDismissed) || !authStatus?.enabled)
  const breadcrumbs = React.useMemo(() => {
    const folderNode = getFolderNode(folderTree, currentFolderId)
    const parts = [
      { label: catalog?.scopeLabel || 'Mailbox review' },
      { label: selectedPstFileName || sessionSummary?.fileName || 'PST' }
    ]
    if (folderNode?.displayName) {
      parts.push({ label: folderNode.displayName })
    }
    return parts
  }, [catalog?.scopeLabel, currentFolderId, folderTree, selectedPstFileName, sessionSummary?.fileName])
  const caseSelectorOptions = React.useMemo(() => getTopLevelCaseOptions(caseOptions), [caseOptions])
  const searchSelectorOptions = React.useMemo(
    () => getSearchOptionsForCase(caseOptions, selectedCasePath),
    [caseOptions, selectedCasePath]
  )
  const activeSearchScopePath = React.useMemo(() => {
    if (sourceType === 'mailbox') {
      return mailboxSearchScopePath || selectedScopePath || selectedCasePath
    }
    return selectedScopePath || selectedCasePath
  }, [mailboxSearchScopePath, selectedCasePath, selectedScopePath, sourceType])
  const searchSessionKey = React.useMemo(() => {
    if (workspaceMode !== 'search') {
      return sessionId || ''
    }
    if (sourceType === 'mailbox' && searchScope === 'pst') {
      return sessionId || ''
    }
    return 'search'
  }, [searchScope, sessionId, sourceType, workspaceMode])
  const activeCatalogScope = React.useMemo(() => {
    if (!catalog?.scopes?.length) {
      return null
    }
    return (
      getScopeEntryForPath(catalog.scopes, selectedScopePath || selectedCasePath) ||
      getScopeEntryForPath(catalog.scopes, selectedScopePath) ||
      getScopeEntryForPath(catalog.scopes, selectedCasePath)
    )
  }, [catalog?.scopes, selectedCasePath, selectedScopePath])

  function invalidateMessagePreview(loading = true): number {
    messagePreviewRequestRef.current += 1
    previewRequestKeyRef.current = ''
    setSelectedMessage(null)
    setMessageLoading(loading)
    return messagePreviewRequestRef.current
  }

  function getMailboxPreviewKey(
    messageId: string,
    fileName = selectedPstFileName || sessionSummary?.fileName || '',
    scopePath = selectedScopePath || selectedCasePath || ''
  ): string {
    return `${normalizeText(fileName)}::${normalizeText(scopePath)}::${normalizeText(messageId)}`
  }

  function getMailboxDetailCacheKey(messageId: string): string {
    return `${sessionId || ''}::${messageId}`
  }

  function clearMailboxDetailCache(messageId?: string): void {
    if (!sessionId) {
      mailboxDetailCacheRef.current.clear()
      mailboxDetailInFlightRef.current.clear()
      mailboxPreviewCacheRef.current.clear()
      return
    }

    if (!messageId) {
      const prefix = `${sessionId}::`
      for (const key of mailboxDetailCacheRef.current.keys()) {
        if (key.startsWith(prefix)) {
          mailboxDetailCacheRef.current.delete(key)
        }
      }
      for (const key of mailboxDetailInFlightRef.current.keys()) {
        if (key.startsWith(prefix)) {
          mailboxDetailInFlightRef.current.delete(key)
        }
      }
      return
    }

    const cacheKey = getMailboxDetailCacheKey(messageId)
    mailboxDetailCacheRef.current.delete(cacheKey)
    mailboxDetailInFlightRef.current.delete(cacheKey)
    clearMailboxPreviewCache(messageId)
  }

  function clearMailboxPreviewCache(messageId?: string): void {
    if (!messageId) {
      mailboxPreviewCacheRef.current.clear()
      return
    }

    const suffix = `::${normalizeText(messageId)}`
    for (const key of mailboxPreviewCacheRef.current.keys()) {
      if (key.endsWith(suffix)) {
        mailboxPreviewCacheRef.current.delete(key)
      }
    }
  }

  function clearMailboxPreviewCacheForRefresh(source: SearchIndexRefreshSource): void {
    if (source === 'mailboxes') {
      clearMailboxDetailCache()
      clearMailboxPreviewCache()
    }
  }

  function loadMailboxMessageDetail(messageId: string): Promise<MessageDetail> {
    if (!sessionId) {
      return Promise.reject(new Error('Mailbox session not available'))
    }

    const previewKey = getMailboxPreviewKey(messageId)
    const cacheKey = getMailboxDetailCacheKey(messageId)
    const cachedPreviewDetail = mailboxPreviewCacheRef.current.get(previewKey)
    if (cachedPreviewDetail) {
      return Promise.resolve(cachedPreviewDetail)
    }

    const cachedDetail = mailboxDetailCacheRef.current.get(cacheKey)
    if (cachedDetail) {
      mailboxPreviewCacheRef.current.set(previewKey, cachedDetail)
      return Promise.resolve(cachedDetail)
    }

    const inFlightRequest = mailboxDetailInFlightRef.current.get(cacheKey)
    if (inFlightRequest) {
      return inFlightRequest
    }

    const request = api.session.messageDetail(sessionId, messageId).then((response) => {
      if (mailboxDetailInFlightRef.current.get(cacheKey) === request) {
        mailboxDetailInFlightRef.current.delete(cacheKey)
        mailboxDetailCacheRef.current.set(cacheKey, response.detail)
        mailboxPreviewCacheRef.current.set(previewKey, response.detail)
      }
      return response.detail
    })

    mailboxDetailInFlightRef.current.set(cacheKey, request)
    return request.catch((error) => {
      if (mailboxDetailInFlightRef.current.get(cacheKey) === request) {
        mailboxDetailInFlightRef.current.delete(cacheKey)
      }
      throw error
    })
  }

  function prefetchMailboxMessageDetail(messageId: string): void {
    void loadMailboxMessageDetail(messageId).catch(() => undefined)
  }

  function applyLoadedPage(page: PageResponse<MessageSummary>): void {
    setCurrentPage(page)
    if (page.page !== pageIndex) {
      setPageIndex(page.page)
    }
    const storedMessageId = readWorkspaceStorageItem('messageId', true, username, '')
    const nextMessageId = page.items.some((item) => item.id === selectedMessageId)
      ? selectedMessageId
      : page.items.find((item) => item.id === storedMessageId)?.id || page.items[0]?.id || ''
    if (nextMessageId && nextMessageId !== selectedMessageId) {
      const nextMessage = page.items.find((item) => item.id === nextMessageId)
      if (nextMessage) {
        void openMessageSummary(nextMessage)
      } else {
        setSelectedMessageId(nextMessageId)
      }
    }
  }

  React.useEffect(() => {
    document.documentElement.dataset.theme = theme
    document.documentElement.style.colorScheme = theme
  }, [theme])

  React.useEffect(() => {
    let cancelled = false

    async function bootstrap(): Promise<void> {
      try {
        const token = getInviteToken()
        const passwordResetToken = getPasswordResetToken()
        setInviteToken(token)
        setResetToken(passwordResetToken)
        if (passwordResetToken) {
          try {
            const response = await api.auth.passwordResetLookup(passwordResetToken)
            if (cancelled) {
              return
            }
            setResetLookup(response)
            setAuthError('')
            setAuthMessage('Create a new password to finish resetting your account.')
          } catch (error) {
            if (cancelled) {
              return
            }
            setResetLookup(null)
            setAuthError(error instanceof Error ? error.message : 'Unable to validate password reset link')
            setAuthMessage('')
          }
          if (cancelled) {
            return
          }
          setAuthStatus(null)
          setPasswordResetAvailable(false)
          setAuthReady(true)
          setAuthView('reset')
          return
        }
        if (token) {
          await loadInvite(token)
          if (cancelled) {
            return
          }
        }

        try {
          const status = await api.auth.me()
          if (cancelled) {
            return
          }
          if (status.authenticated) {
            setAuthStatus(status)
            setPasswordResetAvailable(Boolean(status.passwordResetAvailable))
            setAuthReady(true)
            setAuthView('login')
            setMfaChallengeUsername('')
            setAuthError('')
            setAuthMessage('')
            if (status.user?.username && !status.mfaEnabled) {
              const dismissed = status.mfaEnforced ? false : readReminderDismissed(status.user.username)
              setMfaReminderDismissed(dismissed)
              if (status.mfaEnforced) {
                writeReminderDismissed(status.user.username, false)
              }
            }
          } else if (status.mfaRequired && status.user?.username) {
            setPasswordResetAvailable(Boolean(status.passwordResetAvailable))
            setAuthView('mfa')
            setMfaChallengeUsername(status.user.username)
            setAuthMessage('')
          } else {
            setAuthStatus(null)
            setPasswordResetAvailable(Boolean(status.passwordResetAvailable))
            setAuthReady(true)
            setAuthView(token ? 'invite' : 'login')
          }
        } catch (error) {
          if (error instanceof ApiError && error.statusCode === 401) {
            const payload = error.payload as AuthStatus | undefined
            if (payload?.mfaRequired && payload?.user?.username) {
              setPasswordResetAvailable(Boolean(payload.passwordResetAvailable))
              setAuthView('mfa')
              setMfaChallengeUsername(payload.user.username)
              setAuthMessage('')
            } else {
              setAuthStatus(null)
              setPasswordResetAvailable(Boolean(payload?.passwordResetAvailable))
              setAuthReady(true)
              setAuthView(token ? 'invite' : 'login')
            }
          } else {
            setAuthError(error instanceof Error ? error.message : 'Unable to load authentication status')
            setAuthReady(true)
          }
        }
      } catch (error) {
        setAuthError(error instanceof Error ? error.message : 'Unable to load the application')
        setAuthReady(true)
      } finally {
        if (!cancelled) {
          setBootLoading(false)
        }
      }
    }

    void bootstrap()

    return () => {
      cancelled = true
    }
  }, [])

  React.useEffect(() => {
    if (!authenticated || !username || authFlowActive) {
      return
    }

    setMfaReminderDismissed(mfaEnforced ? false : readReminderDismissed(username))
    const rememberedCasePath = readWorkspaceStorageItem('casePath', true, username, '')
    const rememberedScopePath = readWorkspaceStorageItem('scopePath', true, username, '')
    const rememberedMailboxSearchScopePath = readWorkspaceStorageItem('searchScopePath', true, username, '')
    const normalizedCasePath = getCasePathFromScopePath(rememberedCasePath || rememberedScopePath)
    setSelectedCasePath(normalizedCasePath)
    setSelectedScopePath(rememberedScopePath || normalizedCasePath)
    setMailboxSearchScopePath(rememberedMailboxSearchScopePath || rememberedScopePath || normalizedCasePath)
    setSelectedPstFileName(readWorkspaceStorageItem('pstFileName', true, username, ''))
    setCurrentFolderId(readWorkspaceStorageItem('folderId', true, username, ''))
    setSelectedMessageId(readWorkspaceStorageItem('messageId', true, username, ''))
    const rememberedQuery = readWorkspaceStorageItem('query', true, username, '')
    setSearchQuery(rememberedQuery)
    setSearchInputQuery(rememberedQuery)
    const rememberedSearchScope = (readWorkspaceStorageItem('searchScope', true, username, 'all') as SearchScope) || 'all'
    setSearchScope(rememberedSearchScope === 'pst' ? 'all' : rememberedSearchScope)
    const rememberedSourceType = readWorkspaceStorageItem('sourceType', true, username, 'mailbox') as CorpusSourceType
    setSourceType(
      rememberedSourceType === 'teams' || rememberedSourceType === 'sharepoint' ? rememberedSourceType : 'mailbox'
    )
    setMailOnly(readWorkspaceStorageBool('mailOnly', true, username, false))
    setSort(readWorkspaceStorageItem('sort', true, username, 'date-desc'))
    setReviewFlaggedOnly(readWorkspaceStorageBool('reviewFlaggedOnly', true, username, false))
    setReviewTaggedOnly(readWorkspaceStorageBool('reviewTaggedOnly', true, username, false))
    setActivityFilterUser(readWorkspaceStorageItem('activityFilterUser', true, username, ''))
    setHiddenFiltersOpen(readWorkspaceStorageBool('hiddenFiltersOpen', true, username, hiddenFiltersOpen))
  }, [authenticated, authFlowActive, hiddenFiltersOpen, mfaEnforced, setHiddenFiltersOpen, username])

  React.useEffect(() => {
    if (!authenticated || !canManageUsers || authFlowActive) {
      setSearchIndexRefreshStatuses({ mailboxes: null, items: null })
      stopSearchIndexRefreshPolling('mailboxes')
      stopSearchIndexRefreshPolling('items')
      return
    }

    let cancelled = false

    async function loadRefreshStatus(): Promise<void> {
      try {
        const responses = await Promise.all(
          SEARCH_INDEX_REFRESH_SOURCES.map(async (source) => ({
            source,
            response: await api.pst.refreshSearchIndexStatus(source)
          }))
        )
        if (cancelled) {
          return
        }

        const nextStatuses = { mailboxes: null, items: null } as Record<
          SearchIndexRefreshSource,
          SearchIndexRefreshStatus | null
        >

        for (const { source, response } of responses) {
          const nextStatus = response.status
          if (nextStatus.status === 'running') {
            nextStatuses[source] = nextStatus
            startSearchIndexRefreshPolling(source, nextStatus.jobId)
            continue
          }
          if (nextStatus.status === 'failed') {
            nextStatuses[source] = nextStatus
          } else {
            stopSearchIndexRefreshPolling(source)
          }
        }

        setSearchIndexRefreshStatuses(nextStatuses)
      } catch {
        if (!cancelled) {
          setSearchIndexRefreshStatuses({ mailboxes: null, items: null })
        }
      }
    }

    void loadRefreshStatus()

    return () => {
      cancelled = true
    }
  }, [authenticated, authFlowActive, canManageUsers])

  React.useEffect(() => {
    if (!authenticated || !username || authFlowActive) {
      return
    }

    writeWorkspaceStorageItem('casePath', true, username, selectedCasePath)
    writeWorkspaceStorageItem('scopePath', true, username, selectedScopePath)
    writeWorkspaceStorageItem('searchScopePath', true, username, mailboxSearchScopePath)
    writeWorkspaceStorageItem('pstFileName', true, username, selectedPstFileName)
    writeWorkspaceStorageItem('folderId', true, username, currentFolderId)
    writeWorkspaceStorageItem('messageId', true, username, selectedMessageId)
    writeWorkspaceStorageItem('query', true, username, searchQuery)
    writeWorkspaceStorageItem('searchScope', true, username, searchScope)
    writeWorkspaceStorageItem('mailOnly', true, username, mailOnly)
    writeWorkspaceStorageItem('sort', true, username, sort)
    writeWorkspaceStorageItem('reviewFlaggedOnly', true, username, reviewFlaggedOnly)
    writeWorkspaceStorageItem('reviewTaggedOnly', true, username, reviewTaggedOnly)
    writeWorkspaceStorageItem('sourceType', true, username, sourceType)
    writeWorkspaceStorageItem('hiddenFiltersOpen', true, username, hiddenFiltersOpen)
    writeWorkspaceStorageItem('activityFilterUser', true, username, activityFilterUser)
  }, [
    activityFilterUser,
    authenticated,
    authFlowActive,
    currentFolderId,
    hiddenFiltersOpen,
    mailOnly,
    reviewFlaggedOnly,
    reviewTaggedOnly,
    searchQuery,
    searchScope,
    sourceType,
    selectedCasePath,
    selectedMessageId,
    selectedPstFileName,
    selectedScopePath,
    mailboxSearchScopePath,
    sort,
    username
  ])

  React.useEffect(() => {
    if (!authenticated || !username || authFlowActive) {
      return
    }

    let cancelled = false

    async function loadCatalog(): Promise<void> {
      const loadKey = `${username}|${assignedCasePathsKey}`
      if (catalogLoadKeyRef.current === loadKey) {
        return
      }

      try {
        setMessagesLoading(true)
        const response = await api.pst.catalog('')
        if (cancelled) {
          return
        }
        setCatalogMessage(response.message || '')
        const nextCasePath = (() => {
          const caseOptions = getTopLevelCaseOptions(response.scopes || [])
          if (selectedCasePath && caseOptions.some((option) => option.value === selectedCasePath)) {
            return selectedCasePath
          }
          return caseOptions[0]?.value || getCasePathFromScopePath(response.scopePath || '')
        })()
        const availableSearchOptions = getSearchOptionsForCase(response.scopes || [], nextCasePath)
        const nextScopePath = (() => {
          if (selectedScopePath && availableSearchOptions.some((option) => option.value === selectedScopePath)) {
            return selectedScopePath
          }
          return getDefaultSearchPathForCase(response.scopes || [], nextCasePath)
        })()
        catalogLoadKeyRef.current = loadKey
        setCatalog(response)
        setCaseOptions(response.scopes || [])
        if (nextCasePath !== selectedCasePath) {
          setSelectedCasePath(nextCasePath)
        }
        if (nextScopePath !== selectedScopePath) {
          setSelectedScopePath(nextScopePath)
          if (sourceType === 'mailbox') {
            setMailboxSearchScopePath(nextScopePath)
          }
        }
      } catch (error) {
        if (!cancelled) {
          setCatalogMessage(error instanceof Error ? error.message : 'Unable to load PST catalog')
          setCatalog(null)
          setCatalogFiles([])
          setCaseOptions([])
        }
      } finally {
        if (!cancelled) {
          setMessagesLoading(false)
        }
      }
    }

    void loadCatalog()

    return () => {
      cancelled = true
    }
  }, [assignedCasePathsKey, authenticated, authFlowActive, selectedCasePath, selectedScopePath, username])

  React.useEffect(() => {
    if (!workspaceReady || !authenticated || sourceType !== 'mailbox' || !catalog?.scopes?.length || !activeCatalogScope) {
      return
    }

    const nextFiles = activeCatalogScope.files || []
    const nextFile =
      nextFiles.find((file) => file.fileName === selectedPstFileName)?.fileName || nextFiles[0]?.fileName || ''
    const nextScopePath = activeCatalogScope.scopePath || selectedScopePath || selectedCasePath || ''
    setCatalogFiles(nextFiles)
    setCatalogMessage(
      `Found ${activeCatalogScope.fileCount || 0} mailbox file${
        activeCatalogScope.fileCount === 1 ? '' : 's'
      } in ${activeCatalogScope.scopeLabel || 'PST root'}.`
    )

    const nextMailboxOpenKey = `${normalizeText(nextFile)}::${normalizeText(nextScopePath)}`
    if (skipNextMailboxOpenRef.current && skipNextMailboxOpenRef.current === nextMailboxOpenKey) {
      skipNextMailboxOpenRef.current = ''
      return
    }

    if (nextFile && nextFile !== selectedPstFileName) {
      setSelectedPstFileName(nextFile)
      void openMailbox(nextFile, nextScopePath, catalog)
    }
  }, [
    activeCatalogScope,
    authenticated,
    catalog,
    sourceType,
    selectedCasePath,
    selectedPstFileName,
    selectedScopePath,
    username,
    workspaceReady
  ])

  React.useEffect(() => {
    if (!workspaceReady || !authenticated) {
      return
    }
    if (sourceType !== 'mailbox') {
      setWorkspaceMode('search')
      setSearchScope((current) => (current === 'pst' ? 'search' : current))
      setSelectedMessage(null)
      setSelectedMessageId('')
      setPageIndex(1)
    }
  }, [authenticated, sourceType, workspaceReady])

  React.useEffect(() => {
    const activeSessionId = sessionId
    const activeFolderId = currentFolderId
    if (!workspaceReady || workspaceMode === 'search' || !activeSessionId || !activeFolderId) {
      return
    }
    if (skipNextMessageReloadRef.current) {
      skipNextMessageReloadRef.current = false
      return
    }

    let cancelled = false

    async function loadFolderMessages(): Promise<void> {
      try {
        setMessagesLoading(true)
        const response = await api.session.folderMessages(activeSessionId, activeFolderId, {
          q: '',
          page: pageIndex,
          pageSize: 50,
          mailOnly,
          sort,
          reviewFlagged: reviewFlaggedOnly,
          reviewTagged: reviewTaggedOnly
        })

        if (cancelled) {
          return
        }

        const page = normalizeSearchResultsPage(response.page)
        applyLoadedPage(page)
      } catch (error) {
        if (!cancelled) {
          setCurrentPage(null)
          setAuthError(error instanceof Error ? error.message : 'Unable to load messages')
        }
      } finally {
        if (!cancelled) {
          setMessagesLoading(false)
        }
      }
    }

    void loadFolderMessages()

    return () => {
      cancelled = true
    }
  }, [
    authenticated,
    currentFolderId,
    mailOnly,
    pageIndex,
    reviewFlaggedOnly,
    reviewTaggedOnly,
    sort,
    sessionId,
    username,
    workspaceMode,
    workspaceReady
  ])

  React.useEffect(() => {
    const activeSessionId = sessionId
    if (!workspaceReady || workspaceMode !== 'search') {
      return
    }
    if (skipNextMessageReloadRef.current) {
      skipNextMessageReloadRef.current = false
      return
    }

    const sessionToken = activeSessionId
    let cancelled = false

    async function loadSearchMessages(): Promise<void> {
      try {
        setMessagesLoading(true)
        const response = await api.search({
          scope:
            sourceType === 'mailbox'
              ? searchScope === 'pst' && !activeSessionId
                ? 'all'
                : searchScope
              : searchScope === 'pst'
                ? 'search'
                : searchScope,
          sourceType,
          query: searchQuery,
          mode: deriveSearchMode(searchQuery, 'and'),
          page: pageIndex,
          pageSize: 50,
          mailOnly,
          sort,
          reviewFlagged: reviewFlaggedOnly,
          reviewTagged: reviewTaggedOnly,
          scopePath: activeSearchScopePath,
          sessionId: sourceType === 'mailbox' && searchScope === 'pst' && activeSessionId ? sessionToken : undefined
        })

        if (cancelled) {
          return
        }

        if ('page' in response && response.page) {
          applyLoadedPage(normalizeSearchResultsPage(response.page))
        }
      } catch (error) {
        if (!cancelled) {
          setCurrentPage(null)
          setAuthError(error instanceof Error ? error.message : 'Unable to load messages')
        }
      } finally {
        if (!cancelled) {
          setMessagesLoading(false)
        }
      }
    }

    void loadSearchMessages()

    return () => {
      cancelled = true
    }
  }, [
    activeSearchScopePath,
    authenticated,
    mailOnly,
    pageIndex,
    reviewFlaggedOnly,
    reviewTaggedOnly,
    searchQuery,
    searchScope,
    sort,
    sourceType,
    sessionId,
    workspaceMode,
    workspaceReady
  ])

  React.useEffect(() => {
    const mailboxSessionMatchesSelection =
      !selectedPstFileName || !sessionSummary?.fileName || sessionSummary.fileName === selectedPstFileName

    if (
      !workspaceReady ||
      !sessionId ||
      !mailboxSessionMatchesSelection ||
      !currentPage?.items?.length ||
      sourceType !== 'mailbox'
    ) {
      return
    }

    const currentFileName = selectedPstFileName || sessionSummary?.fileName || ''
    const currentScopePath = selectedScopePath || selectedCasePath
    const mailboxItems = currentPage.items.filter(
      (item) => item.sourceType === 'mailbox' && isSameMailboxContext(item, currentFileName, currentScopePath)
    )
    if (!mailboxItems.length) {
      return
    }

    const selectedIndex = mailboxItems.findIndex((item) => item.id === selectedMessageId)
    const nextItems = selectedIndex >= 0 ? mailboxItems.slice(selectedIndex, selectedIndex + 3) : mailboxItems.slice(0, 3)
    for (const item of nextItems) {
      prefetchMailboxMessageDetail(item.id)
    }
  }, [
    currentPage,
    selectedCasePath,
    selectedMessageId,
    selectedPstFileName,
    selectedScopePath,
    sessionId,
    sessionSummary?.fileName,
    sourceType,
    workspaceReady
  ])

  React.useEffect(() => {
    const activeMessageId = selectedMessageId
    if (!workspaceReady || !activeMessageId) {
      setSelectedMessage(null)
      setMessageLoading(false)
      return
    }
    if (!selectedPageItem) {
      return
    }
    if (selectedPageItem.sourceType && selectedPageItem.sourceType !== 'mailbox') {
      return
    }
    const activeSessionId = sessionId
    const mailboxSessionMatchesSelection =
      !selectedPstFileName || !sessionSummary?.fileName || sessionSummary.fileName === selectedPstFileName
    const previewKey = getMailboxPreviewKey(activeMessageId)

    if (previewRequestKeyRef.current === previewKey) {
      return
    }

    if (selectedMessage?.id === activeMessageId) {
      if (!activeSessionId || !mailboxSessionMatchesSelection) {
        return
      }

      const cachedPreviewDetail = mailboxPreviewCacheRef.current.get(previewKey)
      if (cachedPreviewDetail) {
        previewRequestKeyRef.current = previewKey
        setSelectedMessage(cachedPreviewDetail)
        return
      }

      const cacheKey = getMailboxDetailCacheKey(activeMessageId)
      const cachedDetail = mailboxDetailCacheRef.current.get(cacheKey)
      if (cachedDetail) {
        previewRequestKeyRef.current = previewKey
        mailboxPreviewCacheRef.current.set(previewKey, cachedDetail)
        setSelectedMessage(cachedDetail)
        return
      }

      const hydrationRequestId = messagePreviewRequestRef.current
      previewRequestKeyRef.current = previewKey
      void loadMailboxMessageDetail(activeMessageId)
        .then((detail) => {
          if (
            messagePreviewRequestRef.current !== hydrationRequestId ||
            selectedMessageId !== activeMessageId
          ) {
            return
          }
          setSelectedMessage(detail)
        })
        .catch(() => undefined)
      return
    }

    if (!activeSessionId || !mailboxSessionMatchesSelection) {
      if (!messageLoading) {
        setSelectedMessage(null)
        setMessageLoading(false)
      }
      return
    }

    const cachedPreviewDetail = mailboxPreviewCacheRef.current.get(previewKey)
    if (cachedPreviewDetail) {
      previewRequestKeyRef.current = previewKey
      setSelectedMessage(cachedPreviewDetail)
      setMessageLoading(false)
      return
    }

    const cacheKey = getMailboxDetailCacheKey(activeMessageId)
    const cachedDetail = mailboxDetailCacheRef.current.get(cacheKey)
    if (cachedDetail) {
      previewRequestKeyRef.current = previewKey
      mailboxPreviewCacheRef.current.set(previewKey, cachedDetail)
      setSelectedMessage(cachedDetail)
      setMessageLoading(false)
      return
    }

    const requestId = invalidateMessagePreview(true)
    previewRequestKeyRef.current = previewKey
    let cancelled = false

    async function loadMessage(): Promise<void> {
      try {
        const detail = await loadMailboxMessageDetail(activeMessageId)
        if (cancelled || messagePreviewRequestRef.current !== requestId) {
          return
        }
        setSelectedMessage(detail)
      } catch (error) {
        if (!cancelled && messagePreviewRequestRef.current === requestId) {
          setSelectedMessage(null)
          setAuthError(error instanceof Error ? error.message : 'Unable to load message detail')
        }
      } finally {
        if (!cancelled && messagePreviewRequestRef.current === requestId) {
          setMessageLoading(false)
        }
      }
    }

    void loadMessage()

    return () => {
      cancelled = true
    }
  }, [
    currentPage,
    messageLoading,
    selectedMessage?.id,
    selectedMessageId,
    selectedPageItem?.sourceType,
    selectedPstFileName,
    sessionId,
    sessionSummary?.fileName,
    workspaceReady
  ])

  React.useEffect(() => {
    if (!workspaceReady || (!selectedMessage && !messageLoading)) {
      setFullViewOpen(false)
    }
  }, [messageLoading, selectedMessage, workspaceReady])

  React.useEffect(() => {
    if (!workspaceReady || !sessionId || !currentPage?.items?.length) {
      return
    }

    const first = currentPage.items.find((item) => item.id === selectedMessageId) || currentPage.items[0]
    if (first && first.id !== selectedMessageId) {
      void openMessageSummary(first)
    }
  }, [currentPage, selectedMessageId, sessionId, workspaceReady])

  React.useEffect(() => {
    if (!workspaceReady) {
      return
    }

    let cancelled = false

    async function loadHidden(): Promise<void> {
      try {
        const response: HiddenRulesResponse = await api.hiddenFilters.list()
        if (!cancelled) {
          setHiddenRules(response.items || [])
        }
      } catch {
        if (!cancelled) {
          setHiddenRules([])
        }
      }
    }

    void loadHidden()

    return () => {
      cancelled = true
    }
  }, [workspaceReady])

  React.useEffect(() => {
    if (!usersDialogOpen || !canManageUsers) {
      return
    }

    void loadUsers()
  }, [usersDialogOpen, canManageUsers])

  React.useEffect(() => {
    if (!smtpDialogOpen || !canManageUsers) {
      return
    }

    void loadSmtpSettings()
  }, [canManageUsers, smtpDialogOpen])

  React.useEffect(() => {
    if (!activityDialogOpen || !canManageUsers) {
      return
    }

    void loadActivityLog(activityFilterUser)
  }, [activityDialogOpen, activityFilterUser, canManageUsers])

  React.useEffect(() => {
    if (!selectedAdminUser || !usersDialogOpen) {
      return
    }

    void loadUserActivity(selectedAdminUser)
  }, [selectedAdminUser, usersDialogOpen])

  React.useEffect(() => {
    if (!caseAccessDialogUser) {
      return
    }
    const record = users.find((user) => user.username === caseAccessDialogUser) || null
    setCaseAccessDraftPaths([...(record?.assignedCasePaths || [])])
  }, [caseAccessDialogUser, users])

  async function handleLogin(usernameInput: string, password: string): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    setAuthMessage('')
    try {
      const response = await api.auth.login(usernameInput, password)
      setPasswordResetAvailable(Boolean(response.passwordResetAvailable))
      if (response.mfaRequired && response.user?.username) {
        setAuthView('mfa')
        setMfaChallengeUsername(response.user.username)
        setAuthMessage('')
        setAuthStatus(null)
        return
      }
      setAuthStatus(response)
      setAuthView('login')
      setMfaChallengeUsername('')
      if (response.user?.username && !response.mfaEnabled) {
        const dismissed = response.mfaEnforced ? false : readReminderDismissed(response.user.username)
        setMfaReminderDismissed(dismissed)
        if (response.mfaEnforced) {
          writeReminderDismissed(response.user.username, false)
        }
      }
    } catch (error) {
      if (error instanceof ApiError && error.payload && typeof error.payload === 'object') {
        const payload = error.payload as Partial<AuthStatus>
        setPasswordResetAvailable(Boolean(payload.passwordResetAvailable))
      }
      setAuthError(error instanceof Error ? error.message : 'Login failed')
    } finally {
      setAuthBusy(false)
    }
  }

  async function handlePasswordResetRequest(usernameOrEmail: string): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    setAuthMessage('')
    try {
      await api.auth.passwordResetRequest(usernameOrEmail)
      setAuthMessage('If the account is eligible, a password reset link has been sent.')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to request a password reset')
    } finally {
      setAuthBusy(false)
    }
  }

  async function handlePasswordResetConfirm(password: string, confirmPassword: string): Promise<void> {
    if (!resetToken) {
      setAuthError('Password reset link is missing.')
      return
    }

    setAuthBusy(true)
    setAuthError('')
    setAuthMessage('')
    try {
      const response = await api.auth.passwordResetConfirm(resetToken, password, confirmPassword)
      setAuthStatus(null)
      setPasswordResetAvailable(false)
      setResetLookup(null)
      setResetToken(null)
      setAuthView('login')
      setMfaChallengeUsername('')
      setAuthMessage(`Password updated for ${response.user.username}. Sign in with your new password.`)
      window.history.replaceState({}, '', '/')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to update password')
    } finally {
      setAuthBusy(false)
    }
  }

  async function handleMfaChallenge(code: string): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    try {
      const response = await api.auth.mfaChallenge(code)
      setAuthStatus(response)
      setPasswordResetAvailable(Boolean(response.passwordResetAvailable))
      setAuthView('login')
      setMfaChallengeUsername('')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to verify MFA code')
    } finally {
      setAuthBusy(false)
    }
  }

  async function handleLogout(): Promise<void> {
    try {
      await api.auth.logout()
    } finally {
      if (username) {
        writeReminderDismissed(username, false)
      }
      setMfaReminderDismissed(false)
      setAuthStatus(null)
      catalogLoadKeyRef.current = ''
      setInvite(null)
      setInviteStep('password')
      setInviteMfaAvailable(false)
      setInviteMfaEnforced(false)
      setInviteSetup(null)
      setInviteRecoveryCodes([])
      setPasswordResetAvailable(false)
      setResetLookup(null)
      setResetToken(null)
      setSelfMfaOpen(false)
      setSelfMfaSetup(null)
      setSelfMfaRecoveryCodes([])
      setCatalog(null)
      setCatalogFiles([])
      setCaseOptions([])
      setSelectedCasePath('')
      setSelectedScopePath('')
      setSelectedPstFileName('')
      setSessionId(null)
      setSessionSummary(null)
      setFolderTree(null)
      setCurrentFolderId('')
      setCurrentPage(null)
      setPageIndex(1)
      setSelectedMessageId('')
      setSelectedMessage(null)
      setWorkspaceMode('folder')
      setSearchQuery('')
      setSearchInputQuery('')
      setSearchScope('all')
      setSourceType('mailbox')
      setMailOnly(false)
      setSort('date-desc')
      setReviewFlaggedOnly(false)
      setReviewTaggedOnly(false)
      setHiddenRules([])
      setUsers([])
      setUsersError('')
      setUsersMessage('')
      setCaseAccessDialogUser('')
      setCaseAccessDraftPaths([])
      setUserActivity([])
      setUserActivityError('')
      setSmtpError('')
      setSmtpMessage('')
      setActivityEntries([])
      setActivityError('')
      setAuthView('login')
      setAuthMessage('')
      setAuthError('')
      setSearchIndexRefreshStatuses({ mailboxes: null, items: null })
      setSearchIndexRefreshActionBusyBySource({ mailboxes: false, items: false })
      stopSearchIndexRefreshPolling('mailboxes')
      stopSearchIndexRefreshPolling('items')
    }
  }

  async function loadInvite(token: string): Promise<void> {
    setInviteLoading(true)
    setAuthError('')
    try {
      const response: InviteLookupResponse = await api.auth.inviteLookup(token)
      setInvite(response.invite)
      setInviteMfaAvailable(!response.invite.mfaEnabled)
      setInviteMfaEnforced(Boolean(response.invite.mfaEnforced))
      setAuthView('invite')
      setAuthMessage('Invite validated. Choose a password to continue.')
    } catch (error) {
      setInvite(null)
      setInviteStep('password')
      setInviteMfaAvailable(false)
      setInviteMfaEnforced(false)
      setAuthError(error instanceof Error ? error.message : 'Unable to validate invite')
    } finally {
      setInviteLoading(false)
    }
  }

  async function finishInviteOnboarding(): Promise<void> {
    const status = await api.auth.me().catch(() => null)
    if (status?.authenticated) {
      setAuthStatus(status)
      setAuthView('login')
      setInvite(null)
      setInviteToken(null)
      window.history.replaceState({}, '', '/')
    }
  }

  async function handleInviteAccept(password: string): Promise<void> {
    if (!inviteToken) {
      return
    }
    setAuthBusy(true)
    setAuthError('')
    try {
      const response: InviteAcceptResponse = await api.auth.inviteAccept(inviteToken, password)
      setInviteMfaAvailable(Boolean(response.mfaAvailable))
      setInviteMfaEnforced(Boolean(response.user?.mfaEnforced))
      setAuthMessage(
        response.user?.mfaEnforced
          ? 'Password saved. MFA setup is required before you can continue.'
          : 'Password saved. You can now continue with optional MFA setup.'
      )
      if (response.mfaAvailable) {
        setInviteStep('prompt')
      } else {
        await finishInviteOnboarding()
      }
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to save password')
    } finally {
      setAuthBusy(false)
    }
  }

  async function startInviteMfa(): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    try {
      const response = await api.auth.mfaEnrollmentStart()
      setInviteSetup(response)
      setInviteStep('setup')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to start MFA setup')
    } finally {
      setAuthBusy(false)
    }
  }

  async function submitInviteMfa(code: string): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    try {
      const response: MfaEnrollmentCompleteResponse = await api.auth.mfaEnrollmentComplete(code)
      setInviteRecoveryCodes(response.recoveryCodes)
      setInviteStep('complete')
      setAuthMessage('Recovery codes generated. Download them before continuing.')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to verify MFA code')
    } finally {
      setAuthBusy(false)
    }
  }

  async function continueInviteToPlatform(): Promise<void> {
    if (invite?.username && !invite.mfaEnforced) {
      writeReminderDismissed(invite.username, true)
    }
    setMfaReminderDismissed(true)
    await finishInviteOnboarding()
  }

  async function openSelfServiceMfa(): Promise<void> {
    setSelfMfaOpen(true)
    setSelfMfaError('')
    setSelfMfaMessage('Loading MFA setup...')
    setSelfMfaRecoveryCodes([])
    try {
      const response: MfaEnrollmentStartResponse = await api.auth.mfaEnrollmentStart()
      setSelfMfaSetup(response)
      setSelfMfaMessage('Scan the QR code or enter the setup key manually.')
    } catch (error) {
      setSelfMfaError(error instanceof Error ? error.message : 'Unable to start MFA setup')
      setSelfMfaMessage('')
    }
  }

  async function submitSelfServiceMfa(code: string): Promise<void> {
    setSelfMfaLoading(true)
    setSelfMfaError('')
    try {
      const response: MfaEnrollmentCompleteResponse = await api.auth.mfaEnrollmentComplete(code)
      setSelfMfaRecoveryCodes(response.recoveryCodes)
      setSelfMfaMessage('Recovery codes generated. Download them before continuing.')
      setAuthStatus((current) => (current ? { ...current, mfaEnabled: true } : current))
    } catch (error) {
      setSelfMfaError(error instanceof Error ? error.message : 'Unable to verify MFA code')
    } finally {
      setSelfMfaLoading(false)
    }
  }

  async function finishSelfServiceMfa(): Promise<void> {
    setSelfMfaOpen(false)
    setSelfMfaSetup(null)
    setSelfMfaRecoveryCodes([])
    setSelfMfaMessage('')
    setSelfMfaError('')
    const status = await api.auth.me().catch(() => null)
    if (status?.authenticated) {
      setAuthStatus(status)
      setPasswordResetAvailable(Boolean(status.passwordResetAvailable))
    }
  }

  async function loadUsers(): Promise<void> {
    setUsersLoading(true)
    setUsersError('')
    try {
      const response: UsersResponse = await api.auth.users()
      setUsers(response.users || [])
      if (!selectedAdminUser && response.users.length) {
        setSelectedAdminUser(response.users[0].username)
      }
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to load users')
    } finally {
      setUsersLoading(false)
    }
  }

  async function inviteUser(): Promise<void> {
    if (!inviteUsername.trim() || !inviteEmail.trim()) {
      setUsersError('Username and email are required.')
      return
    }
    setUsersLoading(true)
    setUsersError('')
    setUsersMessage('')
    try {
      await api.auth.inviteUser(inviteUsername.trim(), inviteEmail.trim())
      setUsersMessage(`Invite sent to ${inviteEmail.trim()}.`)
      setInviteUsername('')
      setInviteEmail('')
      await loadUsers()
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to send invite')
    } finally {
      setUsersLoading(false)
    }
  }

  async function deleteUser(usernameToDelete: string): Promise<void> {
    if (!window.confirm(`Delete ${usernameToDelete}?`)) {
      return
    }
    setUsersLoading(true)
    try {
      await api.auth.deleteUser(usernameToDelete)
      await loadUsers()
      if (selectedAdminUser === usernameToDelete) {
        setSelectedAdminUser('')
        setUserActivity([])
      }
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to delete user')
    } finally {
      setUsersLoading(false)
    }
  }

  async function resendInvite(usernameToResend: string): Promise<void> {
    setUsersLoading(true)
    try {
      await api.auth.resendInvite(usernameToResend)
      await loadUsers()
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to resend invite')
    } finally {
      setUsersLoading(false)
    }
  }

  async function revokeInvite(usernameToRevoke: string): Promise<void> {
    setUsersLoading(true)
    try {
      await api.auth.revokeInvite(usernameToRevoke)
      await loadUsers()
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to revoke invite')
    } finally {
      setUsersLoading(false)
    }
  }

  async function resetMfa(usernameToReset: string): Promise<void> {
    setUsersLoading(true)
    try {
      await api.auth.resetMfa(usernameToReset)
      await loadUsers()
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to reset MFA')
    } finally {
      setUsersLoading(false)
    }
  }

  async function toggleMfaEnforcement(usernameToToggle: string, enforced: boolean): Promise<void> {
    setUsersLoading(true)
    try {
      const response = await api.auth.setMfaEnforced(usernameToToggle, enforced)
      setAuthStatus((current) =>
        current && usernameToToggle === username ? { ...current, mfaEnforced: Boolean(response.user.mfaEnforced) } : current
      )
      await loadUsers()
      if (selectedAdminUser === usernameToToggle) {
        await loadUserActivity(usernameToToggle)
      }
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to update MFA enforcement')
    } finally {
      setUsersLoading(false)
    }
  }

  function openCaseAccessDialog(usernameToOpen: string): void {
    const current = users.find((user) => user.username === usernameToOpen) || null
    setCaseAccessDialogUser(usernameToOpen)
    setCaseAccessDraftPaths([...(current?.assignedCasePaths || [])])
  }

  async function saveUserCaseAccess(usernameToUpdate: string, assignedCasePaths: string[]): Promise<void> {
    setUsersLoading(true)
    setUsersError('')
    setUsersMessage('')
    try {
      const response = await api.auth.setUserAccess(usernameToUpdate, assignedCasePaths)
      setUsers((currentUsers) =>
        currentUsers.map((user) => (user.username === usernameToUpdate ? response.user : user))
      )
      if (usernameToUpdate === username && authStatus?.user) {
        setAuthStatus((currentStatus) =>
          currentStatus && currentStatus.user
            ? {
                ...currentStatus,
                user: {
                  ...currentStatus.user,
                  assignedCasePaths: [...response.user.assignedCasePaths]
                }
              }
            : currentStatus
        )
      }
      setUsersMessage(`Case access updated for ${usernameToUpdate}.`)
      setCaseAccessDialogUser('')
      setCaseAccessDraftPaths([])
      if (selectedAdminUser === usernameToUpdate) {
        await loadUserActivity(usernameToUpdate)
      }
    } catch (error) {
      setUsersError(error instanceof Error ? error.message : 'Unable to update case access')
    } finally {
      setUsersLoading(false)
    }
  }

  async function loadUserActivity(usernameToLoad: string): Promise<void> {
    setUserActivityLoading(true)
    setUserActivityError('')
    try {
      const response: ActivityLogResponse = await api.activityLog.list(100, usernameToLoad)
      setUserActivity(response.entries || [])
    } catch (error) {
      setUserActivityError(error instanceof Error ? error.message : 'Unable to load activity log')
    } finally {
      setUserActivityLoading(false)
    }
  }

  async function loadSmtpSettings(): Promise<void> {
    setSmtpLoading(true)
    setSmtpError('')
    try {
      const response: SmtpSettingsResponse = await api.settings.smtpGet()
      setSmtpForm({
        enabled: response.settings.enabled,
        host: response.settings.host,
        port: response.settings.port,
        secure: response.settings.secure,
        username: response.settings.username,
        fromName: response.settings.fromName,
        fromAddress: response.settings.fromAddress,
        replyTo: response.settings.replyTo
      })
      setSmtpPassword('')
    } catch (error) {
      setSmtpError(error instanceof Error ? error.message : 'Unable to load SMTP settings')
    } finally {
      setSmtpLoading(false)
    }
  }

  async function saveSmtpSettings(): Promise<void> {
    setSmtpLoading(true)
    setSmtpError('')
    setSmtpMessage('')
    try {
      const response: SmtpSettingsResponse = await api.settings.smtpPut({
        ...smtpForm,
        password: smtpPassword || undefined
      })
      setSmtpForm({
        enabled: response.settings.enabled,
        host: response.settings.host,
        port: response.settings.port,
        secure: response.settings.secure,
        username: response.settings.username,
        fromName: response.settings.fromName,
        fromAddress: response.settings.fromAddress,
        replyTo: response.settings.replyTo
      })
      setSmtpPassword('')
      setSmtpMessage('SMTP settings saved.')
    } catch (error) {
      setSmtpError(error instanceof Error ? error.message : 'Unable to save SMTP settings')
    } finally {
      setSmtpLoading(false)
    }
  }

  async function sendSmtpTest(): Promise<void> {
    if (!smtpTestRecipient.trim()) {
      setSmtpError('Test recipient is required.')
      return
    }
    setSmtpLoading(true)
    setSmtpError('')
    setSmtpMessage('')
    try {
      const response = await api.settings.smtpTest({
        ...smtpForm,
        password: smtpPassword || undefined,
        recipient: smtpTestRecipient.trim()
      })
      setSmtpMessage(`Test email sent to ${response.recipient}.`)
    } catch (error) {
      setSmtpError(error instanceof Error ? error.message : 'Unable to send test email')
    } finally {
      setSmtpLoading(false)
    }
  }

  async function loadActivityLog(usernameFilter = activityFilterUser): Promise<void> {
    setActivityLoading(true)
    setActivityError('')
    try {
      const response: ActivityLogResponse = await api.activityLog.list(100, usernameFilter)
      setActivityEntries(response.entries || [])
    } catch (error) {
      setActivityError(error instanceof Error ? error.message : 'Unable to load activity log')
    } finally {
      setActivityLoading(false)
    }
  }

  async function refreshCurrentPage(): Promise<void> {
    if (workspaceMode === 'search') {
      await loadSearchPage(pageIndex)
      return
    }
    if (!sessionId || !currentFolderId) {
      return
    } else {
      await loadFolderPage(currentFolderId, pageIndex)
    }
  }

  React.useEffect(() => {
    refreshCurrentPageRef.current = refreshCurrentPage
  }, [refreshCurrentPage])

  async function loadFolderPage(folderId: string, nextPage = 1): Promise<void> {
    if (!sessionId) {
      return
    }
    skipNextMessageReloadRef.current = true
    setWorkspaceMode('folder')
    setMessagesLoading(true)
    try {
      setPageIndex(nextPage)
      const response = await api.session.folderMessages(sessionId, folderId, {
        q: '',
        page: nextPage,
        pageSize: 50,
        mailOnly,
        sort,
        reviewFlagged: reviewFlaggedOnly,
        reviewTagged: reviewTaggedOnly
      })
      setCurrentPage(normalizeSearchResultsPage(response.page))
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to load folder messages')
    } finally {
      setMessagesLoading(false)
    }
  }

  async function loadSearchPage(nextPage = 1, query = searchQuery): Promise<void> {
    skipNextMessageReloadRef.current = true
    setWorkspaceMode('search')
    setMessagesLoading(true)
    try {
      setPageIndex(nextPage)
      const effectiveSearchScope =
        sourceType === 'mailbox'
          ? searchScope === 'pst' && !sessionId
            ? 'all'
            : searchScope
          : 'search'
      if (effectiveSearchScope !== searchScope) {
        setSearchScope(effectiveSearchScope)
      }
      const response = await api.search({
        scope: effectiveSearchScope,
        sourceType,
        query,
        mode: deriveSearchMode(query, 'and'),
        page: nextPage,
        pageSize: 50,
        mailOnly,
        sort,
        reviewFlagged: reviewFlaggedOnly,
        reviewTagged: reviewTaggedOnly,
        scopePath: activeSearchScopePath,
        sessionId: sourceType === 'mailbox' && effectiveSearchScope === 'pst' && sessionId ? sessionId : undefined
      })
      setCurrentPage(normalizeSearchResultsPage(response.page))
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to run search')
    } finally {
      setMessagesLoading(false)
    }
  }

  async function openMailbox(
    fileName: string,
    scopePath: string,
    catalogResponse?: PstCatalogResponse,
    options: OpenMailboxOptions = {}
  ): Promise<void> {
    if (!options.preservePreview) {
      invalidateMessagePreview(Boolean(options.selectedMessageId))
    }
    const effectiveScope = scopePath || selectedScopePath || selectedCasePath || catalogResponse?.scopePath || catalog?.scopePath || ''
    try {
      const response: SessionOpenResponse = await api.pst.open(effectiveScope, fileName)
      if (!options.preservePreview) {
        clearMailboxDetailCache()
      }
      const nextCasePath = getCasePathFromScopePath(effectiveScope || response.scopePath)
      setSessionId(response.sessionId)
      setSessionSummary(response.summary)
      setFolderTree(response.tree)
      setCurrentFolderId(response.tree?.id || '')
      setSelectedPstFileName(response.fileName)
      setSelectedCasePath(nextCasePath)
      setSelectedScopePath(effectiveScope || response.scopePath)
      if (options.selectedMessageId && !options.preservePreview) {
        setSelectedMessageId(options.selectedMessageId)
        writeWorkspaceStorageItem('messageId', true, username, options.selectedMessageId)
      }
      if (!options.preserveWorkspaceMode) {
        setWorkspaceMode('folder')
        setPageIndex(1)
        const storedFolderId = readWorkspaceStorageItem('folderId', true, username, '')
        const folderToOpen = getFolderNode(response.tree, storedFolderId)?.id || response.tree?.id || ''
        if (folderToOpen) {
          setCurrentFolderId(folderToOpen)
        }
      }
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to open PST')
    }
  }

  async function selectFolder(folderId: string): Promise<void> {
    setCurrentFolderId(folderId)
    writeWorkspaceStorageItem('folderId', true, username, folderId)
    if (workspaceMode === 'search') {
      setWorkspaceMode('folder')
      setSearchQuery('')
      setSearchInputQuery('')
    }
    await loadFolderPage(folderId, 1)
  }

  async function openMessage(messageId: string): Promise<void> {
    invalidateMessagePreview(true)
    setSelectedMessageId(messageId)
    writeWorkspaceStorageItem('messageId', true, username, messageId)
  }

  async function openMessageSummary(message: MessageSummary): Promise<void> {
    if (message.sourceType && message.sourceType !== 'mailbox') {
      const requestId = invalidateMessagePreview(true)
      setSelectedMessageId(message.id)
      writeWorkspaceStorageItem('messageId', true, username, message.id)
      setWorkspaceMode('search')
      try {
        const response = await api.item.detail(message.id)
        if (messagePreviewRequestRef.current !== requestId) {
          return
        }
        setSelectedMessage(response.detail)
      } catch (error) {
        if (messagePreviewRequestRef.current === requestId) {
          setAuthError(error instanceof Error ? error.message : 'Unable to load item detail')
        }
      } finally {
        if (messagePreviewRequestRef.current === requestId) {
          setMessageLoading(false)
        }
      }
      return
    }

    const currentFileName = selectedPstFileName || sessionSummary?.fileName || ''
    const currentScopePath = selectedScopePath || selectedCasePath

    if (!isSameMailboxContext(message, currentFileName, currentScopePath)) {
      const nextFileName = message.fileName || currentFileName
      const nextScopePath = message.scopePath || currentScopePath
      if (nextFileName && nextScopePath) {
        const requestId = invalidateMessagePreview(true)
        const nextCasePath = getCasePathFromScopePath(nextScopePath)
        skipNextMailboxOpenRef.current = `${normalizeText(nextFileName)}::${normalizeText(nextScopePath)}`
        if (workspaceMode === 'search' && sourceType === 'mailbox' && searchScope !== 'pst') {
          skipNextMessageReloadRef.current = true
        }
        setSelectedMessageId(message.id)
        writeWorkspaceStorageItem('messageId', true, username, message.id)
        setWorkspaceMode('search')
        setSelectedPstFileName(nextFileName)
        setSelectedCasePath(nextCasePath)
        setSelectedScopePath(nextScopePath)
        writeWorkspaceStorageItem('casePath', true, username, nextCasePath)
        writeWorkspaceStorageItem('scopePath', true, username, nextScopePath)
        writeWorkspaceStorageItem('pstFileName', true, username, nextFileName)
        try {
          const response = await api.item.detail(message.id)
          if (messagePreviewRequestRef.current !== requestId) {
            return
          }
          setSelectedMessage(response.detail)
          const nextPreviewKey = getMailboxPreviewKey(message.id, nextFileName, nextScopePath)
          previewRequestKeyRef.current = nextPreviewKey
          mailboxPreviewCacheRef.current.set(nextPreviewKey, response.detail)
          void openMailbox(nextFileName, nextScopePath, catalog || undefined, {
            preserveWorkspaceMode: true,
            preservePreview: true
          })
        } catch (error) {
          if (messagePreviewRequestRef.current === requestId) {
            setAuthError(error instanceof Error ? error.message : 'Unable to load item detail')
          }
        } finally {
          if (messagePreviewRequestRef.current === requestId) {
            setMessageLoading(false)
          }
        }
        return
      }
    }

    await openMessage(message.id)
  }

  async function openPrevMessage(): Promise<void> {
    if (!currentPage?.items?.length) {
      return
    }
    const index = currentPage.items.findIndex((item) => item.id === selectedMessageId)
    const next = index > 0 ? currentPage.items[index - 1] : null
    if (next) {
      await openMessageSummary(next)
    }
  }

  async function openNextMessage(): Promise<void> {
    if (!currentPage?.items?.length) {
      return
    }
    const index = currentPage.items.findIndex((item) => item.id === selectedMessageId)
    const next = index >= 0 && index < currentPage.items.length - 1 ? currentPage.items[index + 1] : null
    if (next) {
      await openMessageSummary(next)
    }
  }

  async function downloadJson(): Promise<void> {
    if (sessionId && selectedMessageId) {
      triggerDownload(api.session.messageJsonUrl(sessionId, selectedMessageId), `${selectedMessageId}.json`)
    }
  }

  async function downloadEml(): Promise<void> {
    if (sessionId && selectedMessageId) {
      triggerDownload(api.session.messageEmlUrl(sessionId, selectedMessageId), `${selectedMessageId}.eml`)
    }
  }

  async function toggleFlag(): Promise<void> {
    if (!selectedMessageId) {
      return
    }
    if (selectedMessage?.archivePath || !sessionId) {
      await api.item.updateReview(selectedMessageId, {
        flagged: !(selectedMessage?.review?.flagged ?? false),
        tags: selectedMessage?.review?.tags || []
      })
      await refreshCurrentPage()
      const response = await api.item.detail(selectedMessageId)
      setSelectedMessage(response.detail)
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      flagged: !(selectedMessage?.review?.flagged ?? false),
      tags: selectedMessage?.review?.tags || []
    })
    clearMailboxDetailCache(selectedMessageId)
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function clearFlag(): Promise<void> {
    if (!selectedMessageId) {
      return
    }
    if (selectedMessage?.archivePath || !sessionId) {
      await api.item.clearReview(selectedMessageId)
      await refreshCurrentPage()
      const response = await api.item.detail(selectedMessageId)
      setSelectedMessage(response.detail)
      return
    }
    await api.session.clearReview(sessionId, selectedMessageId)
    clearMailboxDetailCache(selectedMessageId)
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function addTag(tag: string): Promise<void> {
    if (!selectedMessageId) {
      return
    }
    if (selectedMessage?.archivePath || !sessionId) {
      await api.item.updateReview(selectedMessageId, {
        tags: Array.from(new Set([...(selectedMessage?.review?.tags || []), tag]))
      })
      await refreshCurrentPage()
      const response = await api.item.detail(selectedMessageId)
      setSelectedMessage(response.detail)
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      tags: Array.from(new Set([...(selectedMessage?.review?.tags || []), tag]))
    })
    clearMailboxDetailCache(selectedMessageId)
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function removeTag(tag: string): Promise<void> {
    if (!selectedMessageId) {
      return
    }
    if (selectedMessage?.archivePath || !sessionId) {
      await api.item.updateReview(selectedMessageId, {
        tags: (selectedMessage?.review?.tags || []).filter((item) => item !== tag)
      })
      await refreshCurrentPage()
      const response = await api.item.detail(selectedMessageId)
      setSelectedMessage(response.detail)
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      tags: (selectedMessage?.review?.tags || []).filter((item) => item !== tag)
    })
    clearMailboxDetailCache(selectedMessageId)
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  function stopSearchIndexRefreshPolling(source: SearchIndexRefreshSource): void {
    const timeout = searchIndexRefreshPollTimeoutRef.current[source]
    if (timeout !== null) {
      window.clearTimeout(timeout)
    }
    searchIndexRefreshPollTimeoutRef.current[source] = null
    searchIndexRefreshPollInFlightRef.current[source] = false
    searchIndexRefreshPollJobIdRef.current[source] = null
  }

  function scheduleSearchIndexRefreshPolling(source: SearchIndexRefreshSource, jobId: string): void {
    if (searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
      return
    }
    const timeout = searchIndexRefreshPollTimeoutRef.current[source]
    if (timeout !== null) {
      window.clearTimeout(timeout)
    }
    searchIndexRefreshPollTimeoutRef.current[source] = window.setTimeout(() => {
      searchIndexRefreshPollTimeoutRef.current[source] = null
      void pollSearchIndexRefreshStatus(source, jobId)
    }, SEARCH_INDEX_REFRESH_POLL_INTERVAL_MS)
  }

  function startSearchIndexRefreshPolling(source: SearchIndexRefreshSource, jobId: string): void {
    searchIndexRefreshPollJobIdRef.current[source] = jobId
    void pollSearchIndexRefreshStatus(source, jobId)
  }

  async function pollSearchIndexRefreshStatus(
    source: SearchIndexRefreshSource,
    jobId: string
  ): Promise<void> {
    if (searchIndexRefreshPollInFlightRef.current[source] || searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
      return
    }

    searchIndexRefreshPollInFlightRef.current[source] = true
    try {
      const response = await api.pst.refreshSearchIndexStatus(source)
      if (searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
        return
      }

      const nextStatus = response.status
      setSearchIndexRefreshStatuses((current) => ({ ...current, [source]: nextStatus }))
      if (nextStatus.status === 'running') {
        scheduleSearchIndexRefreshPolling(source, jobId)
        return
      }

      if (nextStatus.status === 'succeeded') {
        const hiddenRulesResponse = await api.hiddenFilters.list()
        if (searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
          return
        }
        setHiddenRules(hiddenRulesResponse.items || [])
        clearMailboxPreviewCacheForRefresh(source)
        await refreshCurrentPageRef.current()
        if (searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
          return
        }
        stopSearchIndexRefreshPolling(source)
        setSearchIndexRefreshStatuses((current) => ({ ...current, [source]: null }))
        return
      }

      stopSearchIndexRefreshPolling(source)
      setSearchIndexRefreshStatuses((current) => ({
        ...current,
        [source]: nextStatus.status === 'failed' ? nextStatus : null
      }))
    } catch (error) {
      if (searchIndexRefreshPollJobIdRef.current[source] !== jobId) {
        return
      }
      stopSearchIndexRefreshPolling(source)
      setSearchIndexRefreshStatuses((current) => ({
        ...current,
        [source]: {
          source,
          jobId,
          status: 'failed',
          trigger: current[source]?.trigger || 'manual',
          startedAt: current[source]?.startedAt || null,
          completedAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          summary: null,
          error: error instanceof Error ? error.message : 'Unable to refresh search index'
        }
      }))
    } finally {
      searchIndexRefreshPollInFlightRef.current[source] = false
    }
  }

  async function refreshSearchIndex(source: SearchIndexRefreshSource): Promise<SearchIndexRefreshStatus> {
    const response = await api.pst.refreshSearchIndex(source)
    return response.status
  }

  async function handleRefreshSearchIndex(source: SearchIndexRefreshSource): Promise<void> {
    if (searchIndexRefreshActionBusyBySource[source] || searchIndexRefreshStatuses[source]?.status === 'running') {
      return
    }

    setSearchIndexRefreshActionBusyBySource((current) => ({ ...current, [source]: true }))
    try {
      const status = await refreshSearchIndex(source)
      setSearchIndexRefreshStatuses((current) => ({ ...current, [source]: status }))
      if (status.status === 'running') {
        startSearchIndexRefreshPolling(source, status.jobId)
        return
      }
      if (status.status === 'succeeded') {
        await api.hiddenFilters.list().then((result) => setHiddenRules(result.items || []))
        clearMailboxPreviewCacheForRefresh(source)
        await refreshCurrentPage()
        setSearchIndexRefreshStatuses((current) => ({ ...current, [source]: null }))
      }
    } catch (error) {
      setSearchIndexRefreshStatuses((current) => ({
        ...current,
        [source]: {
          source,
          jobId: null,
          status: 'failed',
          trigger: 'manual',
          startedAt: null,
          completedAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          summary: null,
          error: error instanceof Error ? error.message : 'Unable to refresh search index'
        }
      }))
    } finally {
      setSearchIndexRefreshActionBusyBySource((current) => ({ ...current, [source]: false }))
    }
  }

  async function createHiddenFilter(kind: 'address' | 'subject', value: string, label?: string): Promise<void> {
    await api.hiddenFilters.create(kind, value, label)
    const response = await api.hiddenFilters.list()
    setHiddenRules(response.items || [])
    await refreshCurrentPage()
  }

  async function deleteHiddenFilter(filterId: string): Promise<void> {
    await api.hiddenFilters.delete(filterId)
    const response = await api.hiddenFilters.list()
    setHiddenRules(response.items || [])
    await refreshCurrentPage()
  }

  async function runSearch(): Promise<void> {
    const effectiveScope =
      sourceType === 'mailbox'
        ? searchScope
        : 'search'
    if (!searchInputQuery.trim() && effectiveScope === 'pst') {
      setSearchQuery('')
      await refreshCurrentPage()
      return
    }
    setSearchQuery(searchInputQuery)
    await loadSearchPage(1, searchInputQuery)
  }

  function getWorkspaceItemsRequestParams(): Record<string, string | number | boolean | undefined> {
    const effectiveSearchScope =
      sourceType === 'mailbox'
        ? searchScope === 'pst' && !sessionId
          ? 'all'
          : searchScope
        : 'search'
    const scopePath = workspaceMode === 'search' ? activeSearchScopePath : selectedScopePath || selectedCasePath

    return {
      workspaceMode,
      scope: effectiveSearchScope,
      sourceType,
      query: searchQuery,
      mode: deriveSearchMode(searchQuery, 'and'),
      mailOnly,
      sort,
      reviewFlagged: reviewFlaggedOnly,
      reviewTagged: reviewTaggedOnly,
      scopePath,
      sessionId:
        workspaceMode === 'folder'
          ? sessionId || undefined
          : sourceType === 'mailbox' && effectiveSearchScope === 'pst' && sessionId
            ? sessionId
            : undefined,
      folderId: workspaceMode === 'folder' ? currentFolderId : undefined
    }
  }

  async function downloadWorkspaceItemsCsv(): Promise<void> {
    triggerDownload(api.workspace.itemsCsvUrl(getWorkspaceItemsRequestParams()), 'loaded-items.csv')
  }

  async function clearAllFlags(): Promise<void> {
    if (clearFlagsLoading) {
      return
    }
    setClearFlagsLoading(true)
    setClearFlagsError('')
    try {
      await api.workspace.clearAllFlags(getWorkspaceItemsRequestParams())
      setClearFlagsDialogOpen(false)
      await refreshCurrentPage()
    } catch (error) {
      setClearFlagsError(error instanceof Error ? error.message : 'Unable to clear flags')
    } finally {
      setClearFlagsLoading(false)
    }
  }

  async function setMfaReminderSkipped(): Promise<void> {
    if (username && !mfaEnforced) {
      writeReminderDismissed(username, true)
    }
    setMfaReminderDismissed(true)
  }

  const settingsMenu = (
    <div className="space-y-2">
      {!mfaEnabled ? (
        <PopoverClose asChild>
          <Button
            variant="ghost"
            className="w-full justify-start gap-3"
            onClick={() => {
              void openSelfServiceMfa()
            }}
          >
            <KeyRound className="h-4 w-4" />
            Set up MFA
          </Button>
        </PopoverClose>
      ) : null}
      <PopoverClose asChild>
        <Button
          variant="ghost"
          className="w-full justify-start gap-3"
          onClick={() => {
            void downloadWorkspaceItemsCsv()
          }}
        >
          <Download className="h-4 w-4" />
          Download items CSV
        </Button>
      </PopoverClose>
      <PopoverClose asChild>
        <Button
          variant="danger"
          className="w-full justify-start gap-3"
          onClick={() => setClearFlagsDialogOpen(true)}
        >
          <RotateCcw className="h-4 w-4" />
          Clear all flags
        </Button>
      </PopoverClose>
      {canManageUsers ? (
        <>
          <PopoverClose asChild>
            <Button
              variant="ghost"
              className="w-full justify-start gap-3"
              onClick={() => setUsersDialogOpen(true)}
            >
              <Users className="h-4 w-4" />
              Manage users
            </Button>
          </PopoverClose>
          <PopoverClose asChild>
            <Button
              variant="ghost"
              className="w-full justify-start gap-3"
              onClick={() => setSmtpDialogOpen(true)}
            >
              <Mail className="h-4 w-4" />
              SMTP settings
            </Button>
          </PopoverClose>
          <PopoverClose asChild>
            <Button
              variant="ghost"
              className="w-full justify-start gap-3"
              onClick={() => setActivityDialogOpen(true)}
            >
              <Activity className="h-4 w-4" />
              Activity log
            </Button>
          </PopoverClose>
        </>
      ) : null}
      <Separator />
      <PopoverClose asChild>
        <Button
          variant="ghost"
          className="w-full justify-start gap-3"
          onClick={toggleTheme}
        >
          {theme === 'dark' ? <SunMedium className="h-4 w-4" /> : <MoonStar className="h-4 w-4" />}
          {theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
        </Button>
      </PopoverClose>
    </div>
  )

  const sidebarNode = workspaceReady ? (
    <Sidebar
      catalogMessage={catalogMessage}
      caseOptions={caseSelectorOptions}
      selectedCasePath={selectedCasePath}
      selectedScopePath={selectedScopePath}
      sourceType={sourceType}
      sourceCounts={currentPage?.sourceCounts}
      searchOptions={searchSelectorOptions}
      catalogFiles={catalogFiles}
      selectedPstFileName={selectedPstFileName}
      onCaseChange={(value) => {
        const nextCasePath = normalizeText(value)
        const nextScopePath = getDefaultSearchPathForCase(caseOptions, nextCasePath)
        setSelectedCasePath(nextCasePath)
        setSelectedScopePath(nextScopePath)
        if (sourceType === 'mailbox') {
          setMailboxSearchScopePath(nextScopePath)
        }
        writeWorkspaceStorageItem('casePath', true, username, nextCasePath)
        writeWorkspaceStorageItem('scopePath', true, username, nextScopePath)
      }}
      onScopeChange={(value) => {
        const nextScopePath = value
        const nextCasePath = getCasePathFromScopePath(nextScopePath)
        setSelectedCasePath(nextCasePath)
        setSelectedScopePath(nextScopePath)
        if (sourceType === 'mailbox') {
          setMailboxSearchScopePath(nextScopePath)
        }
        writeWorkspaceStorageItem('casePath', true, username, nextCasePath)
        writeWorkspaceStorageItem('scopePath', true, username, nextScopePath)
      }}
      onSourceTypeChange={(value) => {
        setSourceType(value)
        if (value === 'mailbox') {
          setMailboxSearchScopePath(selectedScopePath || selectedCasePath)
        }
        writeWorkspaceStorageItem('sourceType', true, username, value)
      }}
      canRefreshSearchIndex={canManageUsers}
      searchIndexRefreshStatuses={searchIndexRefreshStatuses}
      searchIndexRefreshBusyBySource={searchIndexRefreshActionBusyBySource}
      onRefreshSearchIndex={(source) => {
        void handleRefreshSearchIndex(source)
      }}
      onOpenMailbox={(fileName, scopePath) => {
        void openMailbox(fileName, scopePath)
      }}
      folderTree={folderTree}
      currentFolderId={currentFolderId}
      onSelectFolder={(folderId) => {
        void selectFolder(folderId)
      }}
    />
  ) : (
    <div className="panel-surface flex h-full items-center justify-center p-6">
      <EmptyState title="Workspace locked" description="Sign in to start exploring PST files." />
    </div>
  )

  const messageNode = workspaceReady ? (
    <MessageList
      page={currentPage}
      loading={messagesLoading}
      query={searchInputQuery}
      activeQuery={searchQuery}
      sourceType={sourceType}
      searchScope={searchScope}
      mailOnly={mailOnly}
      sort={sort}
      reviewFlaggedOnly={reviewFlaggedOnly}
      reviewTaggedOnly={reviewTaggedOnly}
      onQueryChange={setSearchInputQuery}
      onSearch={() => {
        void runSearch()
      }}
      onSearchScopeChange={(value) => setSearchScope(value)}
      onMailOnlyChange={(value) => setMailOnly(value)}
      onSortChange={(value) => setSort(value)}
      onReviewFlaggedChange={(value) => setReviewFlaggedOnly(value)}
      onReviewTaggedChange={(value) => setReviewTaggedOnly(value)}
      onSelectMessage={(message) => {
        void openMessageSummary(message)
      }}
      onPrevPage={() => {
        const nextPage = Math.max(1, pageIndex - 1)
        setPageIndex(nextPage)
        workspaceMode === 'search' ? void loadSearchPage(nextPage) : void loadFolderPage(currentFolderId, nextPage)
      }}
      onNextPage={() => {
        const nextPage = pageIndex + 1
        setPageIndex(nextPage)
        workspaceMode === 'search' ? void loadSearchPage(nextPage) : void loadFolderPage(currentFolderId, nextPage)
      }}
      onOpenBundle={() => {
        if (!sessionId) {
          return
        }
        const scope = workspaceMode === 'search' && searchScope !== 'pst' ? searchScope : 'pst'
        const url = api.session.flaggedBundleUrl({
          scope,
          scopePath: activeSearchScopePath,
          sessionId,
          query: searchQuery,
          mailOnly,
          sort,
          reviewFlagged: reviewFlaggedOnly,
          reviewTagged: reviewTaggedOnly
        })
        triggerDownload(url, 'flagged-bundle.zip')
      }}
      selectedMessageId={selectedMessageId}
      sessionId={sessionId}
    />
  ) : (
    <div className="panel-surface flex h-full items-center justify-center p-6">
      <EmptyState title="No mailbox open" description="Open a PST from the left sidebar to start browsing." />
    </div>
  )

  const previewNode = workspaceReady ? (
    <EmailPreview
      detail={selectedMessage}
      loading={messageLoading}
      theme={theme}
      onDownloadJson={() => {
        void downloadJson()
      }}
      onDownloadEml={() => {
        void downloadEml()
      }}
      onDownloadItem={() => {
        if (selectedMessage?.downloadUrl) {
          triggerDownload(
            selectedMessage.downloadUrl,
            selectedMessage.downloadFilename || selectedMessage.archiveEntryName || selectedMessage.subject || undefined
          )
        }
      }}
      onToggleFlag={() => {
        void toggleFlag()
      }}
      onClearReview={() => {
        void clearFlag()
      }}
      onOpenTags={() => {
        setTagsDialogOpen(true)
      }}
      onOpenFullView={() => {
        setFullViewOpen(true)
      }}
      tagCount={selectedMessage?.review?.tags?.length || 0}
      onOpenAttachment={(attachment) => {
        if (sessionId && selectedMessageId && attachment.downloadUrl) {
          triggerDownload(attachment.downloadUrl, attachment.downloadFilename || attachment.filename || undefined)
        }
      }}
      onOpenPrev={() => {
        void openPrevMessage()
      }}
      onOpenNext={() => {
        void openNextMessage()
      }}
      canNavigatePrev={Boolean(currentPage?.items?.length && currentPage.items.findIndex((item) => item.id === selectedMessageId) > 0)}
      canNavigateNext={Boolean(
        currentPage?.items?.length &&
          currentPage.items.findIndex((item) => item.id === selectedMessageId) < (currentPage.items.length || 0) - 1
      )}
    />
  ) : (
    <div className="panel-surface flex h-full items-center justify-center p-6">
      <EmptyState title="Preview empty" description="Open a mailbox and select a message to inspect it here." />
    </div>
  )

  const appContent = authenticated ? (
    <>
      <AppShell
        userName={username}
        authenticated={authenticated}
        settingsMenu={settingsMenu}
        breadcrumbs={breadcrumbs}
        sidebarCollapsed={sidebarCollapsed}
        previewCollapsed={previewCollapsed}
        onLogout={() => {
          void handleLogout()
        }}
        sidebar={sidebarNode}
        messagePanel={messageNode}
        preview={previewNode}
      />

      <Dialog
        open={fullViewOpen && (Boolean(selectedMessage) || messageLoading)}
        onOpenChange={(open) => {
          setFullViewOpen(open)
        }}
      >
        <DialogContent showCloseButton={false} className="h-[min(96vh,980px)] w-[min(98vw,1600px)] overflow-hidden p-0">
          <div className="flex h-full min-h-0 flex-col">
            <div className="sr-only">
              <DialogTitle>Email full view</DialogTitle>
            </div>
            <div className="flex items-center justify-end gap-2 border-b border-[color:var(--line)] px-4 py-3">
              <IconButton
                label="Previous message"
                onClick={() => {
                  void openPrevMessage()
                }}
                disabled={!Boolean(currentPage?.items?.length && currentPage.items.findIndex((item) => item.id === selectedMessageId) > 0)}
              >
                <ChevronLeft className="h-4 w-4" />
              </IconButton>
              <IconButton
                label="Next message"
                onClick={() => {
                  void openNextMessage()
                }}
                disabled={
                  !Boolean(
                    currentPage?.items?.length &&
                      currentPage.items.findIndex((item) => item.id === selectedMessageId) < (currentPage.items.length || 0) - 1
                  )
                }
              >
                <ChevronRight className="h-4 w-4" />
              </IconButton>
              <IconButton
                label="Close full view"
                onClick={() => {
                  setFullViewOpen(false)
                }}
              >
                <X className="h-4 w-4" />
              </IconButton>
            </div>
            <EmailPreview
              detail={selectedMessage}
              loading={messageLoading}
              theme={theme}
              onDownloadJson={() => {
                void downloadJson()
              }}
              onDownloadEml={() => {
                void downloadEml()
              }}
              onDownloadItem={() => {
                if (selectedMessage?.downloadUrl) {
                  triggerDownload(
                    selectedMessage.downloadUrl,
                    selectedMessage.downloadFilename || selectedMessage.archiveEntryName || selectedMessage.subject || undefined
                  )
                }
              }}
              onToggleFlag={() => {
                void toggleFlag()
              }}
              onClearReview={() => {
                void clearFlag()
              }}
              onOpenTags={() => {
                setTagsDialogOpen(true)
              }}
              tagCount={selectedMessage?.review?.tags?.length || 0}
              showNavigationControls={false}
              onOpenAttachment={(attachment) => {
                if (sessionId && selectedMessageId && attachment.downloadUrl) {
                  triggerDownload(attachment.downloadUrl, attachment.downloadFilename || attachment.filename || undefined)
                }
              }}
              onOpenPrev={() => {
                void openPrevMessage()
              }}
              onOpenNext={() => {
                void openNextMessage()
              }}
              canNavigatePrev={Boolean(currentPage?.items?.length && currentPage.items.findIndex((item) => item.id === selectedMessageId) > 0)}
              canNavigateNext={Boolean(
                currentPage?.items?.length &&
                  currentPage.items.findIndex((item) => item.id === selectedMessageId) < (currentPage.items.length || 0) - 1
              )}
            />
          </div>
        </DialogContent>
      </Dialog>

      <MfaReminderDialog
        open={showReminder}
        username={username}
        allowSkip={!mfaEnforced}
        onSetup={() => {
          void openSelfServiceMfa()
        }}
        onSkip={() => {
          void setMfaReminderSkipped()
        }}
      />

      <MfaSetupDialog
        open={selfMfaOpen}
        loading={selfMfaLoading}
        message={selfMfaMessage}
        error={selfMfaError}
        data={selfMfaSetup}
        recoveryCodes={selfMfaRecoveryCodes}
        onSubmit={(code) => {
          void submitSelfServiceMfa(code)
        }}
        onClose={() => {
          setSelfMfaOpen(false)
          setSelfMfaSetup(null)
          setSelfMfaRecoveryCodes([])
          setSelfMfaMessage('')
          setSelfMfaError('')
        }}
        onDownload={() => {
          if (selfMfaRecoveryCodes.length) {
            downloadTextFile('mfa-recovery-codes.txt', selfMfaRecoveryCodes.join('\n'))
          }
        }}
        onFinish={() => {
          void finishSelfServiceMfa()
        }}
      />

      <TagManagerDialog
        open={tagsDialogOpen && Boolean(selectedMessage)}
        tags={selectedMessage?.review?.tags || []}
        subject={selectedMessage?.subject || 'Manage tags for the selected message.'}
        onOpenChange={(open) => {
          setTagsDialogOpen(open)
        }}
        onAddTag={(tag) => {
          void addTag(tag)
        }}
        onRemoveTag={(tag) => {
          void removeTag(tag)
        }}
      />

      <UserManagementDialog
        open={usersDialogOpen}
        loading={usersLoading}
        error={usersError}
        message={usersMessage}
        users={users}
        selectedUser={selectedAdminUser}
        activity={userActivity}
        activityLoading={userActivityLoading}
        activityError={userActivityError}
        inviteUsername={inviteUsername}
        inviteEmail={inviteEmail}
        onInviteUsernameChange={setInviteUsername}
        onInviteEmailChange={setInviteEmail}
        onInvite={inviteUser}
        onClose={() => {
          setUsersDialogOpen(false)
          setCaseAccessDialogUser('')
          setCaseAccessDraftPaths([])
        }}
        onRefresh={() => {
          void loadUsers()
        }}
        onSelectUser={(value) => {
          setSelectedAdminUser(value)
          void loadUserActivity(value)
        }}
        onDeleteUser={(value) => {
          void deleteUser(value)
        }}
        onResendInvite={(value) => {
          void resendInvite(value)
        }}
        onRevokeInvite={(value) => {
          void revokeInvite(value)
        }}
        onResetMfa={(value) => {
          void resetMfa(value)
        }}
        onToggleMfaEnforcement={(value, enforced) => {
          void toggleMfaEnforcement(value, enforced)
        }}
        onOpenCaseAccess={(value) => {
          openCaseAccessDialog(value)
        }}
        onCopyInvite={(value) => {
          const selected = users.find((item) => item.username === value)
          if (selected?.inviteUrl) {
            navigator.clipboard?.writeText(selected.inviteUrl).catch(() => undefined)
          }
        }}
      />

      <CaseAccessDialog
        open={Boolean(caseAccessDialogUser)}
        loading={usersLoading}
        user={users.find((item) => item.username === caseAccessDialogUser) || null}
        caseOptions={caseSelectorOptions}
        assignedCasePaths={caseAccessDraftPaths}
        onClose={() => {
          setCaseAccessDialogUser('')
          setCaseAccessDraftPaths([])
        }}
        onAssignedCasePathsChange={setCaseAccessDraftPaths}
        onSave={() => {
          if (!caseAccessDialogUser) {
            return
          }
          void saveUserCaseAccess(caseAccessDialogUser, caseAccessDraftPaths)
        }}
      />

      <SmtpSettingsDialog
        open={smtpDialogOpen}
        loading={smtpLoading}
        error={smtpError}
        message={smtpMessage}
        form={smtpForm}
        password={smtpPassword}
        testRecipient={smtpTestRecipient}
        onClose={() => {
          setSmtpDialogOpen(false)
        }}
        onRefresh={() => {
          void loadSmtpSettings()
        }}
        onFormChange={setSmtpForm}
        onPasswordChange={setSmtpPassword}
        onTestRecipientChange={setSmtpTestRecipient}
        onSave={() => {
          void saveSmtpSettings()
        }}
        onSendTest={() => {
          void sendSmtpTest()
        }}
      />

      <ActivityLogDialog
        open={activityDialogOpen}
        loading={activityLoading}
        error={activityError}
        message={activityMessage}
        entries={activityEntries}
        filterUsername={activityFilterUser}
        onFilterUsernameChange={setActivityFilterUser}
        onClose={() => setActivityDialogOpen(false)}
        onRefresh={() => {
          void loadActivityLog(activityFilterUser)
        }}
        onExportCsv={() => {
          triggerDownload(api.activityLog.csvUrl(activityFilterUser), 'activity-log.csv')
        }}
      />

      <Dialog open={clearFlagsDialogOpen} onOpenChange={setClearFlagsDialogOpen}>
        <DialogContent className="w-[min(92vw,520px)]" showCloseButton={!clearFlagsLoading}>
          <div className="flex flex-col gap-5 p-6">
            <div>
              <div className="text-xs font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                Workspace
              </div>
              <DialogTitle className="mt-1 text-2xl">Clear all flags</DialogTitle>
              <DialogDescription>
                This clears your flagged state for the items currently loaded in the workspace.
                Tags are left unchanged.
              </DialogDescription>
            </div>

            {clearFlagsError ? (
              <div className="rounded-2xl border border-[color:rgba(220,38,38,0.22)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {clearFlagsError}
              </div>
            ) : null}

            <div className="flex items-center justify-end gap-3">
              <Button
                variant="ghost"
                onClick={() => setClearFlagsDialogOpen(false)}
                disabled={clearFlagsLoading}
              >
                Cancel
              </Button>
              <Button variant="danger" onClick={() => void clearAllFlags()} disabled={clearFlagsLoading}>
                {clearFlagsLoading ? 'Clearing...' : 'Clear flags'}
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </>
  ) : null

  if (!authReady || bootLoading) {
    return (
      <div className="grid min-h-screen place-items-center bg-[color:var(--bg)] px-4">
        <div className="panel-surface-strong max-w-md p-8 text-center">
          <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
            PST Mail Explorer
          </div>
          <div className="mt-3 text-2xl font-semibold text-[color:var(--text)]">Loading workspace...</div>
          <div className="mt-2 text-sm text-[color:var(--muted)]">Preparing authentication and mailbox data.</div>
        </div>
      </div>
    )
  }

  if (!authenticated || authFlowActive) {
    return (
      <div className="h-screen overflow-hidden bg-[color:var(--bg)] text-[color:var(--text)]">
        <AuthScreen
          view={resetFlowActive ? 'reset' : inviteFlowActive ? 'invite' : authView}
          busy={authBusy || inviteLoading || selfMfaLoading}
          message={authMessage}
          error={authError}
          passwordResetAvailable={passwordResetAvailable}
          invite={invite}
          inviteStep={inviteStep}
          inviteMfaAvailable={inviteMfaAvailable}
          resetLookup={resetLookup}
          inviteMfaEnforced={inviteMfaEnforced}
          inviteSetup={inviteSetup}
          inviteRecoveryCodes={inviteRecoveryCodes}
          onPasswordResetRequest={(usernameOrEmail) => {
            void handlePasswordResetRequest(usernameOrEmail)
          }}
          onPasswordResetConfirm={(password, confirmPassword) => {
            void handlePasswordResetConfirm(password, confirmPassword)
          }}
          onLogin={(user, pass) => {
            void handleLogin(user, pass)
          }}
          onMfaChallenge={(code) => {
            void handleMfaChallenge(code)
          }}
          onInviteAccept={(password) => {
            void handleInviteAccept(password)
          }}
          onInviteMfaStart={() => {
            void startInviteMfa()
          }}
          onInviteMfaSkip={() => {
            void continueInviteToPlatform()
          }}
          onInviteMfaSubmit={(code) => {
            void submitInviteMfa(code)
          }}
          onInviteFinish={() => {
            void continueInviteToPlatform()
          }}
          onOpenLogin={() => {
            setResetToken(null)
            setResetLookup(null)
            setPasswordResetAvailable(false)
            setAuthView('login')
            setAuthError('')
            setAuthMessage('')
            setMfaChallengeUsername('')
          }}
        />
      </div>
    )
  }

  return <>{appContent}</>
}

function UserManagementDialog({
  open,
  loading,
  error,
  message,
  users,
  selectedUser,
  activity,
  activityLoading,
  activityError,
  inviteUsername,
  inviteEmail,
  onInviteUsernameChange,
  onInviteEmailChange,
  onInvite,
  onClose,
  onRefresh,
  onSelectUser,
  onDeleteUser,
  onResendInvite,
  onRevokeInvite,
  onResetMfa,
  onToggleMfaEnforcement,
  onOpenCaseAccess,
  onCopyInvite
}: {
  open: boolean
  loading: boolean
  error: string
  message: string
  users: UserInvite[]
  selectedUser: string
  activity: ActivityLogEntry[]
  activityLoading: boolean
  activityError: string
  inviteUsername: string
  inviteEmail: string
  onInviteUsernameChange: (value: string) => void
  onInviteEmailChange: (value: string) => void
  onInvite: () => void
  onClose: () => void
  onRefresh: () => void
  onSelectUser: (username: string) => void
  onDeleteUser: (username: string) => void
  onResendInvite: (username: string) => void
  onRevokeInvite: (username: string) => void
  onResetMfa: (username: string) => void
  onToggleMfaEnforcement: (username: string, enforced: boolean) => void
  onOpenCaseAccess: (username: string) => void
  onCopyInvite: (username: string) => void
}) {
  return (
    <Dialog open={open}>
      <DialogContent className="w-[min(98vw,1540px)]">
        <div className="flex h-[86vh] flex-col overflow-hidden">
          <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
            <div>
              <div className="text-xs font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">Settings</div>
              <DialogTitle className="mt-1 text-2xl">User management</DialogTitle>
              <DialogDescription>Add users by invite, manage access, and inspect per-user activity.</DialogDescription>
            </div>
            <div className="flex items-center gap-2">
              <Button variant="ghost" onClick={onRefresh}>
                <RefreshCw className="h-4 w-4" />
                Refresh
              </Button>
              <Button variant="ghost" onClick={onClose}>
                Close
              </Button>
            </div>
          </div>
          <div className="grid min-h-0 flex-1 grid-cols-1 gap-4 overflow-hidden p-6 xl:grid-cols-[30%_70%]">
            <div className="flex min-h-0 flex-col gap-4 overflow-hidden">
              <div className="panel-surface-strong p-4">
                <div className="flex items-center gap-2">
                  <UserPlus className="h-4 w-4 text-[color:var(--accent)]" />
                  <div className="text-sm font-semibold text-[color:var(--text)]">Invite user</div>
                </div>
                <div className="mt-4 space-y-3">
                  <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                    Username
                    <Input
                      className="mt-2"
                      value={inviteUsername}
                      onChange={(event) => onInviteUsernameChange(event.target.value)}
                      placeholder="jane.doe"
                    />
                  </label>
                  <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                    Email
                    <Input
                      className="mt-2"
                      value={inviteEmail}
                      onChange={(event) => onInviteEmailChange(event.target.value)}
                      placeholder="jane@example.com"
                    />
                  </label>
                  <Button className="w-full justify-center" onClick={onInvite} disabled={loading}>
                    {loading ? 'Sending...' : 'Send invite'}
                  </Button>
                </div>
                {message ? <div className="mt-3 rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">{message}</div> : null}
                {error ? <div className="mt-3 rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">{error}</div> : null}
              </div>

              <div className="panel-surface-strong min-h-0 flex-1 overflow-hidden">
                <div className="flex items-center justify-between border-b border-[color:var(--line)] px-4 py-3">
                  <div>
                    <div className="text-sm font-semibold text-[color:var(--text)]">Users</div>
                    <div className="text-xs text-[color:var(--muted)]">{users.length} accounts</div>
                  </div>
                  <Badge>{users.length}</Badge>
                </div>
                <ScrollArea className="min-h-0 h-[calc(100%-3.5rem)]">
                  <div className="space-y-2 p-4">
                    {users.length ? (
                      users.map((user) => {
                        const active = user.username === selectedUser
                        return (
                          <div
                            key={user.username}
                            role="button"
                            tabIndex={0}
                            onClick={() => onSelectUser(user.username)}
                            onKeyDown={(event) => {
                              if (event.key === 'Enter' || event.key === ' ') {
                                event.preventDefault()
                                onSelectUser(user.username)
                              }
                            }}
                            className={`flex w-full items-start justify-between gap-3 rounded-2xl border px-4 py-3 text-left transition ${
                              active
                                ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)]'
                                : 'border-[color:var(--line)] bg-[color:var(--surface-soft)] hover:border-[color:var(--accent-soft)]'
                            }`}
                          >
                            <div className="min-w-0">
                              <div className="flex items-center gap-2">
                                <div className="truncate font-semibold text-[color:var(--text)]">{user.username}</div>
                                {user.mfaEnabled ? (
                                  <Badge className="border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">
                                    MFA on
                                  </Badge>
                                ) : (
                                  <Badge>MFA off</Badge>
                                )}
                                {user.mfaEnforced ? (
                                  <Badge className="border-[color:var(--warning-bg)] bg-[color:var(--warning-bg)] text-[color:var(--warning)]">
                                    Required
                                  </Badge>
                                ) : null}
                                <Badge>{user.inviteStatus}</Badge>
                              </div>
                              <div className="mt-1 truncate text-xs text-[color:var(--muted)]">{user.recipientEmail}</div>
                              <div className="mt-1 text-xs text-[color:var(--soft)]">
                                {user.inviteStatus === 'pending' ? 'Pending invite' : user.inviteStatus === 'active' ? 'Active account' : user.inviteStatus}
                              </div>
                              <div className="mt-1 text-xs text-[color:var(--muted)]">
                                {user.assignedCasePaths.length ? `${user.assignedCasePaths.length} assigned case${user.assignedCasePaths.length === 1 ? '' : 's'}` : 'No cases assigned'}
                              </div>
                            </div>
                            <div className="flex shrink-0 items-center gap-1">
                              <IconButton
                                label={`Manage case access for ${user.username}`}
                                className="h-9 w-9"
                                onClick={(event) => {
                                  event.stopPropagation()
                                  onOpenCaseAccess(user.username)
                                }}
                              >
                                <FolderCog className="h-4 w-4" />
                              </IconButton>
                              {user.inviteUrl ? (
                                <IconButton
                                  label="Copy invite link"
                                  className="h-9 w-9"
                                  onClick={(event) => {
                                    event.stopPropagation()
                                    onCopyInvite(user.username)
                                  }}
                                >
                                  <Link2 className="h-4 w-4" />
                                </IconButton>
                              ) : null}
                              {user.inviteStatus === 'pending' ? (
                                <IconButton
                                  label="Resend invite"
                                  className="h-9 w-9"
                                  onClick={(event) => {
                                    event.stopPropagation()
                                    onResendInvite(user.username)
                                  }}
                                >
                                  <Send className="h-4 w-4" />
                                </IconButton>
                              ) : null}
                              {user.inviteStatus === 'pending' ? (
                                <IconButton
                                  label="Revoke invite"
                                  className="h-9 w-9"
                                  onClick={(event) => {
                                    event.stopPropagation()
                                    onRevokeInvite(user.username)
                                  }}
                                >
                                  <RotateCcw className="h-4 w-4" />
                                </IconButton>
                              ) : null}
                              <IconButton
                                label={user.mfaEnforced ? 'Remove MFA enforcement' : 'Enforce MFA'}
                                className="h-9 w-9"
                                onClick={(event) => {
                                  event.stopPropagation()
                                  onToggleMfaEnforcement(user.username, !user.mfaEnforced)
                                }}
                              >
                                {user.mfaEnforced ? <ShieldCheck className="h-4 w-4" /> : <ShieldAlert className="h-4 w-4" />}
                              </IconButton>
                              {user.username !== 'admin' ? (
                                <IconButton
                                  label="Delete user"
                                  className="h-9 w-9 icon-button-danger"
                                  onClick={(event) => {
                                    event.stopPropagation()
                                    onDeleteUser(user.username)
                                  }}
                                >
                                  <Trash2 className="h-4 w-4" />
                                </IconButton>
                              ) : null}
                            </div>
                          </div>
                        )
                      })
                    ) : (
                      <div className="empty-state">No users yet.</div>
                    )}
                  </div>
                </ScrollArea>
              </div>
            </div>

            <div className="panel-surface-strong min-h-0 overflow-hidden">
              <div className="flex items-center justify-between border-b border-[color:var(--line)] px-4 py-3">
                <div>
                  <div className="text-sm font-semibold text-[color:var(--text)]">User activity</div>
                  <div className="text-xs text-[color:var(--muted)]">
                    {selectedUser ? `Viewing activity for ${selectedUser}` : 'Select a user to inspect activity.'}
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <Button variant="ghost" onClick={() => onSelectUser(selectedUser)} disabled={!selectedUser}>
                    <RefreshCw className="h-4 w-4" />
                    Refresh
                  </Button>
                  <Button
                    variant="ghost"
                    onClick={() => {
                      if (!selectedUser) {
                        return
                      }
                      triggerDownload(api.activityLog.csvUrl(selectedUser), `activity-${selectedUser}.csv`)
                    }}
                    disabled={!selectedUser}
                  >
                    <Download className="h-4 w-4" />
                    Export CSV
                  </Button>
                </div>
              </div>
              <div className="grid min-h-0 h-[calc(100%-3.5rem)] gap-4 p-4">
                {activityError ? <div className="rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">{activityError}</div> : null}
                <ScrollArea className="min-h-0">
                  {activityLoading ? (
                    <div className="empty-state">Loading activity...</div>
                  ) : selectedUser ? (
                    <div className="space-y-3">
                      {activity.length ? (
                        activity.map((entry) => (
                            <div key={`${entry.timestamp}-${entry.action}-${entry.target}`} className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4">
                            <div className="flex items-start justify-between gap-4">
                              <div>
                                <div className="text-sm font-semibold text-[color:var(--text)]">{entry.action}</div>
                                <div className="mt-1 text-xs text-[color:var(--muted)]">{entry.target}</div>
                              </div>
                              <Badge>{entry.outcome}</Badge>
                            </div>
                            <div className="mt-3 text-xs text-[color:var(--muted)]">
                              {formatDate(entry.timestamp)} · {entry.request?.method} {entry.request?.path} · {entry.request?.ip}
                            </div>
                          </div>
                        ))
                      ) : (
                        <div className="empty-state">No activity recorded for this user.</div>
                      )}
                    </div>
                  ) : (
                    <div className="empty-state">Choose a user to view their activity log.</div>
                  )}
                </ScrollArea>
              </div>
            </div>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}

function CaseAccessDialog({
  open,
  loading,
  user,
  caseOptions,
  assignedCasePaths,
  onClose,
  onAssignedCasePathsChange,
  onSave
}: {
  open: boolean
  loading: boolean
  user: UserInvite | null
  caseOptions: Array<{ label: string; value: string; count: number }>
  assignedCasePaths: string[]
  onClose: () => void
  onAssignedCasePathsChange: (value: string[]) => void
  onSave: () => void
}) {
  const assignedCaseSet = React.useMemo(() => new Set(assignedCasePaths), [assignedCasePaths])

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) => {
        if (!nextOpen) {
          onClose()
        }
      }}
    >
      <DialogContent className="w-[min(92vw,760px)]">
        <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
          <div>
            <div className="text-xs font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">Settings</div>
            <DialogTitle className="mt-1 text-2xl">Case access</DialogTitle>
            <DialogDescription>
              {user ? `Assign top-level cases for ${user.username}.` : 'Assign top-level cases for the selected user.'}
            </DialogDescription>
          </div>
          <Badge>{assignedCasePaths.length}</Badge>
        </div>

        <div className="space-y-4 px-6 py-5">
          <div className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] px-4 py-3 text-sm text-[color:var(--muted)]">
            Users start with no cases assigned. Leave everything unchecked for no access.
          </div>

          <ScrollArea className="max-h-[48vh] pr-2">
            <div className="space-y-2">
              {caseOptions.length ? (
                caseOptions.map((option) => {
                  const checked = assignedCaseSet.has(option.value)
                  return (
                    <label
                      key={option.value || 'pst-root'}
                      className="flex cursor-pointer items-center justify-between gap-3 rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] px-4 py-3 text-sm text-[color:var(--text)] transition hover:border-[color:var(--accent-soft)]"
                    >
                      <span className="min-w-0 truncate">{option.label}</span>
                      <span className="flex items-center gap-2">
                        {option.count ? <Badge>{option.count}</Badge> : null}
                        <input
                          type="checkbox"
                          checked={checked}
                          disabled={!user || loading}
                          onChange={(event) => {
                            const nextPaths = event.target.checked
                              ? Array.from(new Set([...assignedCasePaths, option.value]))
                              : assignedCasePaths.filter((path) => path !== option.value)
                            onAssignedCasePathsChange(nextPaths)
                          }}
                        />
                      </span>
                    </label>
                  )
                })
              ) : (
                <div className="empty-state">No cases available.</div>
              )}
            </div>
          </ScrollArea>
        </div>

        <div className="flex items-center justify-end gap-2 border-t border-[color:var(--line)] px-6 py-5">
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button onClick={onSave} disabled={loading || !user}>
            {loading ? 'Saving...' : 'Save access'}
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  )
}

function SmtpSettingsDialog({
  open,
  loading,
  error,
  message,
  form,
  password,
  testRecipient,
  onClose,
  onRefresh,
  onFormChange,
  onPasswordChange,
  onTestRecipientChange,
  onSave,
  onSendTest
}: {
  open: boolean
  loading: boolean
  error: string
  message: string
  form: SmtpFormState
  password: string
  testRecipient: string
  onClose: () => void
  onRefresh: () => void
  onFormChange: (value: SmtpFormState) => void
  onPasswordChange: (value: string) => void
  onTestRecipientChange: (value: string) => void
  onSave: () => void
  onSendTest: () => void
}) {
  return (
    <Dialog open={open}>
      <DialogContent className="w-[min(98vw,1200px)]">
        <div className="flex h-[84vh] flex-col overflow-hidden">
          <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
            <div>
              <div className="text-xs font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">Settings</div>
              <DialogTitle className="mt-1 text-2xl">SMTP settings</DialogTitle>
              <DialogDescription>Configure the sender profile used for invite and notification emails.</DialogDescription>
            </div>
            <div className="flex items-center gap-2">
              <Button variant="ghost" onClick={onRefresh}>
                <RefreshCw className="h-4 w-4" />
                Refresh
              </Button>
              <Button variant="ghost" onClick={onClose}>Close</Button>
            </div>
          </div>
          <div className="grid min-h-0 flex-1 gap-4 overflow-hidden p-6 xl:grid-cols-[1.2fr_0.8fr]">
            <div className="panel-surface-strong min-h-0 overflow-auto p-4">
              <div className="grid gap-4 md:grid-cols-2">
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Enabled
                  <div className="mt-2 flex items-center gap-3">
                    <input
                      type="checkbox"
                      checked={form.enabled}
                      onChange={(event) => onFormChange({ ...form, enabled: event.target.checked })}
                    />
                    <span className="text-sm text-[color:var(--muted)]">Use SMTP for outgoing emails</span>
                  </div>
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Secure
                  <div className="mt-2 flex items-center gap-3">
                    <input
                      type="checkbox"
                      checked={form.secure}
                      onChange={(event) => onFormChange({ ...form, secure: event.target.checked })}
                    />
                    <span className="text-sm text-[color:var(--muted)]">Use TLS/SSL</span>
                  </div>
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Host
                  <Input className="mt-2" value={form.host} onChange={(event) => onFormChange({ ...form, host: event.target.value })} />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Port
                  <Input
                    className="mt-2"
                    type="number"
                    min={1}
                    value={form.port}
                    onChange={(event) => onFormChange({ ...form, port: Number(event.target.value || 0) })}
                  />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Username
                  <Input className="mt-2" value={form.username} onChange={(event) => onFormChange({ ...form, username: event.target.value })} />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Password
                  <Input className="mt-2" type="password" value={password} onChange={(event) => onPasswordChange(event.target.value)} placeholder="Leave blank to keep current password" />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  From name
                  <Input className="mt-2" value={form.fromName} onChange={(event) => onFormChange({ ...form, fromName: event.target.value })} />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  From address
                  <Input className="mt-2" value={form.fromAddress} onChange={(event) => onFormChange({ ...form, fromAddress: event.target.value })} />
                </label>
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)] md:col-span-2">
                  Reply-to
                  <Input className="mt-2" value={form.replyTo} onChange={(event) => onFormChange({ ...form, replyTo: event.target.value })} />
                </label>
              </div>

              {error ? <div className="mt-4 rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">{error}</div> : null}
              {message ? <div className="mt-4 rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">{message}</div> : null}

              <div className="mt-4 flex items-center justify-end gap-2">
                <Button variant="ghost" onClick={onClose}>Cancel</Button>
                <Button onClick={onSave} disabled={loading}>
                  {loading ? 'Saving...' : 'Save settings'}
                </Button>
              </div>
            </div>

            <div className="panel-surface-strong min-h-0 overflow-hidden">
              <div className="border-b border-[color:var(--line)] px-4 py-3">
                <div className="text-sm font-semibold text-[color:var(--text)]">Send test email</div>
                <div className="text-xs text-[color:var(--muted)]">Verify the configuration without wiring notifications yet.</div>
              </div>
              <div className="space-y-4 p-4">
                <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Recipient
                  <Input className="mt-2" value={testRecipient} onChange={(event) => onTestRecipientChange(event.target.value)} placeholder="test@example.com" />
                </label>
                <div className="flex items-center gap-2">
                  <Button variant="secondary" className="w-full justify-center" onClick={onSendTest} disabled={loading}>
                    <Send className="h-4 w-4" />
                    {loading ? 'Sending...' : 'Send test email'}
                  </Button>
                </div>
                <div className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4 text-sm text-[color:var(--muted)]">
                  The SMTP password is never returned to the browser. Leave it blank to keep the stored secret.
                </div>
              </div>
            </div>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}

function ActivityLogDialog({
  open,
  loading,
  error,
  message,
  entries,
  filterUsername,
  onFilterUsernameChange,
  onClose,
  onRefresh,
  onExportCsv
}: {
  open: boolean
  loading: boolean
  error: string
  message: string
  entries: ActivityLogEntry[]
  filterUsername: string
  onFilterUsernameChange: (value: string) => void
  onClose: () => void
  onRefresh: () => void
  onExportCsv: () => void
}) {
  return (
    <Dialog open={open}>
      <DialogContent className="w-[min(98vw,1180px)]">
        <div className="flex h-[84vh] flex-col overflow-hidden">
          <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
            <div>
              <div className="text-xs font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">Settings</div>
              <DialogTitle className="mt-1 text-2xl">Activity log</DialogTitle>
              <DialogDescription>Review recent platform activity and export it as CSV.</DialogDescription>
            </div>
            <div className="flex items-center gap-2">
              <Button variant="ghost" onClick={onRefresh}>
                <RefreshCw className="h-4 w-4" />
                Refresh
              </Button>
              <Button variant="ghost" onClick={onExportCsv}>
                <Download className="h-4 w-4" />
                Export CSV
              </Button>
              <Button variant="ghost" onClick={onClose}>Close</Button>
            </div>
          </div>
          <div className="flex min-h-0 flex-1 flex-col gap-4 p-6">
            <div className="grid grid-cols-[1fr_auto] gap-3">
              <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                Username filter
                <Input className="mt-2" value={filterUsername} onChange={(event) => onFilterUsernameChange(event.target.value)} placeholder="Filter by user" />
              </label>
              <div className="self-end">
                <Button variant="secondary" onClick={onRefresh}>
                  <Search className="h-4 w-4" />
                  Search
                </Button>
              </div>
            </div>
            {message ? <div className="rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">{message}</div> : null}
            {error ? <div className="rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">{error}</div> : null}
            <ScrollArea className="min-h-0 flex-1">
              <div className="space-y-3">
                {loading ? (
                  <div className="empty-state">Loading activity...</div>
                ) : entries.length ? (
                  entries.map((entry) => (
                    <div key={`${entry.timestamp}-${entry.action}-${entry.target}`} className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4">
                      <div className="flex items-start justify-between gap-4">
                        <div>
                          <div className="text-sm font-semibold text-[color:var(--text)]">{entry.actor.username}</div>
                          <div className="mt-1 text-sm text-[color:var(--muted)]">{entry.action}</div>
                          <div className="mt-1 text-xs text-[color:var(--soft)]">{entry.target}</div>
                        </div>
                        <Badge>{entry.outcome}</Badge>
                      </div>
                      <div className="mt-3 text-xs text-[color:var(--muted)]">
                        {formatDate(entry.timestamp)} · {entry.request.method} {entry.request.path} · {entry.request.ip}
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="empty-state">No activity entries found.</div>
                )}
              </div>
            </ScrollArea>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}
