import * as React from 'react'
import {
  Activity,
  ChevronLeft,
  ChevronRight,
  Download,
  Flag,
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
import { downloadTextFile, formatDate, normalizeText } from '@/lib/utils'
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
  PageResponse,
  PstCatalogResponse,
  SessionOpenResponse,
  SmtpSettings,
  SmtpSettingsResponse,
  UserInvite,
  UsersResponse
} from '@/types'
import { useUiStore } from '@/store/ui'

type WorkspaceMode = 'folder' | 'search'
type SearchScope = 'pst' | 'search' | 'all'
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

function getInviteToken(pathname = window.location.pathname): string | null {
  const match = pathname.match(/^\/invite\/([^/?#]+)/i)
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
  const [invite, setInvite] = React.useState<UserInvite | null>(null)
  const [inviteLoading, setInviteLoading] = React.useState(false)
  const [inviteStep, setInviteStep] = React.useState<InviteStep>('password')
  const [inviteMfaAvailable, setInviteMfaAvailable] = React.useState(false)
  const [inviteSetup, setInviteSetup] = React.useState<MfaEnrollmentStartResponse | null>(null)
  const [inviteRecoveryCodes, setInviteRecoveryCodes] = React.useState<string[]>([])
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
  const [selectedPstFileName, setSelectedPstFileName] = React.useState('')
  const [isRemovedCatalog, setIsRemovedCatalog] = React.useState(false)
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
  const [searchScope, setSearchScope] = React.useState<SearchScope>('pst')
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
  const [mfaReminderDismissed, setMfaReminderDismissed] = React.useState(false)
  const [tagsDialogOpen, setTagsDialogOpen] = React.useState(false)
  const [fullViewOpen, setFullViewOpen] = React.useState(false)

  const catalogLoadKeyRef = React.useRef('')

  const username = authStatus?.user?.username || ''
  const authenticated = Boolean(authStatus?.authenticated)
  const canManageUsers = Boolean(authStatus?.canManageUsers)
  const mfaEnabled = Boolean(authStatus?.mfaEnabled)
  const inviteFlowActive = Boolean(inviteToken)
  const showReminder = authenticated && !mfaEnabled && !mfaReminderDismissed && !selfMfaOpen && !inviteFlowActive
  const workspaceReady = authenticated && (mfaEnabled || mfaReminderDismissed || !authStatus?.enabled)
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

  React.useEffect(() => {
    document.documentElement.dataset.theme = theme
    document.documentElement.style.colorScheme = theme
  }, [theme])

  React.useEffect(() => {
    let cancelled = false

    async function bootstrap(): Promise<void> {
      try {
        const token = getInviteToken()
        setInviteToken(token)
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
            setAuthReady(true)
            setAuthView('login')
            setMfaChallengeUsername('')
            setAuthError('')
            setAuthMessage('')
            if (status.user?.username && !status.mfaEnabled) {
              setMfaReminderDismissed(readReminderDismissed(status.user.username))
            }
          } else if (status.mfaRequired && status.user?.username) {
            setAuthView('mfa')
            setMfaChallengeUsername(status.user.username)
            setAuthMessage(`Enter the verification code for ${status.user.username}.`)
          } else {
            setAuthStatus(null)
            setAuthReady(true)
            setAuthView(token ? 'invite' : 'login')
          }
        } catch (error) {
          if (error instanceof ApiError && error.statusCode === 401) {
            const payload = error.payload as AuthStatus | undefined
            if (payload?.mfaRequired && payload?.user?.username) {
              setAuthView('mfa')
              setMfaChallengeUsername(payload.user.username)
              setAuthMessage(`Enter the verification code for ${payload.user.username}.`)
            } else {
              setAuthStatus(null)
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
    if (!authenticated || !username) {
      return
    }

    setIsRemovedCatalog(readWorkspaceStorageBool('catalogMode', true, username, false))
    setMfaReminderDismissed(readReminderDismissed(username))
    const rememberedCasePath = readWorkspaceStorageItem('casePath', true, username, '')
    const rememberedScopePath = readWorkspaceStorageItem('scopePath', true, username, '')
    const normalizedCasePath = getCasePathFromScopePath(rememberedCasePath || rememberedScopePath)
    setSelectedCasePath(normalizedCasePath)
    setSelectedScopePath(rememberedScopePath || normalizedCasePath)
    setSelectedPstFileName(readWorkspaceStorageItem('pstFileName', true, username, ''))
    setCurrentFolderId(readWorkspaceStorageItem('folderId', true, username, ''))
    setSelectedMessageId(readWorkspaceStorageItem('messageId', true, username, ''))
    setSearchQuery(readWorkspaceStorageItem('query', true, username, ''))
    setSearchScope((readWorkspaceStorageItem('searchScope', true, username, 'pst') as SearchScope) || 'pst')
    setMailOnly(readWorkspaceStorageBool('mailOnly', true, username, false))
    setSort(readWorkspaceStorageItem('sort', true, username, 'date-desc'))
    setReviewFlaggedOnly(readWorkspaceStorageBool('reviewFlaggedOnly', true, username, false))
    setReviewTaggedOnly(readWorkspaceStorageBool('reviewTaggedOnly', true, username, false))
    setActivityFilterUser(readWorkspaceStorageItem('activityFilterUser', true, username, ''))
    setHiddenFiltersOpen(readWorkspaceStorageBool('hiddenFiltersOpen', true, username, hiddenFiltersOpen))
  }, [authenticated, hiddenFiltersOpen, setHiddenFiltersOpen, username])

  React.useEffect(() => {
    if (!authenticated || !username) {
      return
    }

    writeWorkspaceStorageItem('catalogMode', true, username, isRemovedCatalog)
    writeWorkspaceStorageItem('casePath', true, username, selectedCasePath)
    writeWorkspaceStorageItem('scopePath', true, username, selectedScopePath)
    writeWorkspaceStorageItem('pstFileName', true, username, selectedPstFileName)
    writeWorkspaceStorageItem('folderId', true, username, currentFolderId)
    writeWorkspaceStorageItem('messageId', true, username, selectedMessageId)
    writeWorkspaceStorageItem('query', true, username, searchQuery)
    writeWorkspaceStorageItem('searchScope', true, username, searchScope)
    writeWorkspaceStorageItem('mailOnly', true, username, mailOnly)
    writeWorkspaceStorageItem('sort', true, username, sort)
    writeWorkspaceStorageItem('reviewFlaggedOnly', true, username, reviewFlaggedOnly)
    writeWorkspaceStorageItem('reviewTaggedOnly', true, username, reviewTaggedOnly)
    writeWorkspaceStorageItem('hiddenFiltersOpen', true, username, hiddenFiltersOpen)
    writeWorkspaceStorageItem('activityFilterUser', true, username, activityFilterUser)
  }, [
    activityFilterUser,
    authenticated,
    currentFolderId,
    hiddenFiltersOpen,
    isRemovedCatalog,
    mailOnly,
    reviewFlaggedOnly,
    reviewTaggedOnly,
    searchQuery,
    searchScope,
    selectedCasePath,
    selectedMessageId,
    selectedPstFileName,
    selectedScopePath,
    sort,
    username
  ])

  React.useEffect(() => {
    if (!workspaceReady || !authenticated) {
      return
    }

    let cancelled = false

    async function loadCatalog(): Promise<void> {
      const loadKey = `root|${isRemovedCatalog ? 'removed' : 'active'}`
      if (catalogLoadKeyRef.current === loadKey) {
        return
      }

      try {
        setMessagesLoading(true)
        const response = await api.pst.catalog('', isRemovedCatalog)
        if (cancelled) {
          return
        }
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
  }, [authenticated, isRemovedCatalog, selectedCasePath, selectedScopePath, username, workspaceReady])

  React.useEffect(() => {
    if (!workspaceReady || !authenticated || !catalog?.scopes?.length || !activeCatalogScope) {
      return
    }

    const nextFiles = activeCatalogScope.files || []
    setCatalogFiles(nextFiles)
    setCatalogMessage(
      `Found ${activeCatalogScope.fileCount || 0} mailbox file${
        activeCatalogScope.fileCount === 1 ? '' : 's'
      } in ${activeCatalogScope.scopeLabel || 'PST root'}.`
    )

    const nextFile =
      nextFiles.find((file) => file.fileName === selectedPstFileName)?.fileName || nextFiles[0]?.fileName || ''
    if (nextFile && nextFile !== selectedPstFileName) {
      setSelectedPstFileName(nextFile)
      void openMailbox(nextFile, activeCatalogScope.scopePath || selectedScopePath || selectedCasePath, catalog)
    }
  }, [
    activeCatalogScope,
    authenticated,
    catalog,
    selectedCasePath,
    selectedPstFileName,
    selectedScopePath,
    username,
    workspaceReady
  ])

  React.useEffect(() => {
    const activeSessionId = sessionId
    const activeFolderId = currentFolderId
    if (!workspaceReady || !activeSessionId || !activeFolderId) {
      return
    }
    const sessionToken = activeSessionId

    let cancelled = false

    async function loadMessages(): Promise<void> {
      try {
        setMessagesLoading(true)
        const queryParams =
          workspaceMode === 'search'
            ? {
                scope: searchScope,
                query: searchQuery,
                mode: deriveSearchMode(searchQuery, 'and'),
                page: pageIndex,
                pageSize: 50,
                mailOnly,
                sort,
                reviewFlagged: reviewFlaggedOnly,
                reviewTagged: reviewTaggedOnly,
                scopePath: selectedScopePath || selectedCasePath,
                sessionId: searchScope === 'pst' ? sessionToken : undefined
              }
            : null

        const response = queryParams
          ? await api.search(queryParams)
          : await api.session.folderMessages(sessionToken, activeFolderId, {
              q: searchQuery,
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

        const page = 'page' in response && response.page ? normalizeSearchResultsPage(response.page) : null
        if (page) {
          setCurrentPage(page)
          if (page.page !== pageIndex) {
            setPageIndex(page.page)
          }
          const storedMessageId = readWorkspaceStorageItem('messageId', true, username, '')
          const nextMessageId = page.items.some((item) => item.id === selectedMessageId)
            ? selectedMessageId
            : page.items.find((item) => item.id === storedMessageId)?.id || page.items[0]?.id || ''
          if (nextMessageId && nextMessageId !== selectedMessageId) {
            setSelectedMessageId(nextMessageId)
          }
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

    void loadMessages()

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
    searchQuery,
    searchScope,
    selectedCasePath,
    selectedMessageId,
    selectedScopePath,
    sessionId,
    sort,
    username,
    workspaceMode,
    workspaceReady
  ])

  React.useEffect(() => {
    const activeSessionId = sessionId
    const activeMessageId = selectedMessageId
    if (!workspaceReady || !activeSessionId || !activeMessageId) {
      setSelectedMessage(null)
      return
    }
    const sessionToken = activeSessionId

    let cancelled = false

    async function loadMessage(): Promise<void> {
      try {
        setMessageLoading(true)
        const response = await api.session.messageDetail(sessionToken, activeMessageId)
        if (cancelled) {
          return
        }
        setSelectedMessage(response.detail)
      } catch (error) {
        if (!cancelled) {
          setSelectedMessage(null)
          setAuthError(error instanceof Error ? error.message : 'Unable to load message detail')
        }
      } finally {
        if (!cancelled) {
          setMessageLoading(false)
        }
      }
    }

    void loadMessage()

    return () => {
      cancelled = true
    }
  }, [selectedMessageId, sessionId, workspaceReady])

  React.useEffect(() => {
    if (!workspaceReady || !selectedMessage) {
      setFullViewOpen(false)
    }
  }, [selectedMessage, workspaceReady])

  React.useEffect(() => {
    if (!workspaceReady || !sessionId || !currentPage?.items?.length) {
      return
    }

    const first = currentPage.items.find((item) => item.id === selectedMessageId) || currentPage.items[0]
    if (first && first.id !== selectedMessageId) {
      setSelectedMessageId(first.id)
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

  async function handleLogin(usernameInput: string, password: string): Promise<void> {
    setAuthBusy(true)
    setAuthError('')
    setAuthMessage('')
    try {
      const response = await api.auth.login(usernameInput, password)
      if (response.mfaRequired && response.user?.username) {
        setAuthView('mfa')
        setMfaChallengeUsername(response.user.username)
        setAuthMessage(`Enter the verification code for ${response.user.username}.`)
        setAuthStatus(null)
        return
      }
      setAuthStatus(response)
      setAuthView('login')
      setMfaChallengeUsername('')
      if (response.user?.username && !response.mfaEnabled) {
        writeReminderDismissed(response.user.username, readReminderDismissed(response.user.username))
      }
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Login failed')
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
      setInvite(null)
      setInviteStep('password')
      setInviteSetup(null)
      setInviteRecoveryCodes([])
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
      setSearchScope('pst')
      setMailOnly(false)
      setSort('date-desc')
      setReviewFlaggedOnly(false)
      setReviewTaggedOnly(false)
      setHiddenRules([])
      setUsers([])
      setUsersError('')
      setUsersMessage('')
      setUserActivity([])
      setUserActivityError('')
      setSmtpError('')
      setSmtpMessage('')
      setActivityEntries([])
      setActivityError('')
      setAuthView('login')
      setAuthMessage('')
      setAuthError('')
    }
  }

  async function loadInvite(token: string): Promise<void> {
    setInviteLoading(true)
    setAuthError('')
    try {
      const response: InviteLookupResponse = await api.auth.inviteLookup(token)
      setInvite(response.invite)
      setAuthView('invite')
      setAuthMessage('Invite validated. Choose a password to continue.')
    } catch (error) {
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
      setAuthMessage('Password saved. You can now continue with optional MFA setup.')
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
    if (invite?.username) {
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
    if (!sessionId || !currentFolderId) {
      return
    }
    if (workspaceMode === 'search') {
      await loadSearchPage(pageIndex)
    } else {
      await loadFolderPage(currentFolderId, pageIndex)
    }
  }

  async function loadFolderPage(folderId: string, nextPage = 1): Promise<void> {
    if (!sessionId) {
      return
    }
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
      setWorkspaceMode('folder')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to load folder messages')
    } finally {
      setMessagesLoading(false)
    }
  }

  async function loadSearchPage(nextPage = 1): Promise<void> {
    if (!sessionId) {
      return
    }
    setMessagesLoading(true)
    try {
      setPageIndex(nextPage)
      const response = await api.search({
        scope: searchScope,
        query: searchQuery,
        mode: deriveSearchMode(searchQuery, 'and'),
        page: nextPage,
        pageSize: 50,
        mailOnly,
        sort,
        reviewFlagged: reviewFlaggedOnly,
        reviewTagged: reviewTaggedOnly,
        scopePath: selectedScopePath || selectedCasePath,
        sessionId: searchScope === 'pst' && sessionId ? sessionId : undefined
      })
      setCurrentPage(normalizeSearchResultsPage(response.page))
      setWorkspaceMode('search')
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Unable to run search')
    } finally {
      setMessagesLoading(false)
    }
  }

  async function openMailbox(fileName: string, scopePath: string, catalogResponse?: PstCatalogResponse): Promise<void> {
    const effectiveScope = scopePath || selectedScopePath || selectedCasePath || catalogResponse?.scopePath || catalog?.scopePath || ''
    try {
      const response: SessionOpenResponse = await api.pst.open(effectiveScope, fileName)
      const nextCasePath = getCasePathFromScopePath(effectiveScope || response.scopePath)
      setSessionId(response.sessionId)
      setSessionSummary(response.summary)
      setFolderTree(response.tree)
      setCurrentFolderId(response.tree?.id || '')
      setSelectedPstFileName(response.fileName)
      setSelectedCasePath(nextCasePath)
      setSelectedScopePath(effectiveScope || response.scopePath)
      setWorkspaceMode('folder')
      setPageIndex(1)
      const storedFolderId = readWorkspaceStorageItem('folderId', true, username, '')
      const folderToOpen = getFolderNode(response.tree, storedFolderId)?.id || response.tree?.id || ''
      if (folderToOpen) {
        setCurrentFolderId(folderToOpen)
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
    }
    await loadFolderPage(folderId, 1)
  }

  async function openMessage(messageId: string): Promise<void> {
    setSelectedMessageId(messageId)
    writeWorkspaceStorageItem('messageId', true, username, messageId)
  }

  async function openPrevMessage(): Promise<void> {
    if (!currentPage?.items?.length) {
      return
    }
    const index = currentPage.items.findIndex((item) => item.id === selectedMessageId)
    const next = index > 0 ? currentPage.items[index - 1] : null
    if (next) {
      await openMessage(next.id)
    }
  }

  async function openNextMessage(): Promise<void> {
    if (!currentPage?.items?.length) {
      return
    }
    const index = currentPage.items.findIndex((item) => item.id === selectedMessageId)
    const next = index >= 0 && index < currentPage.items.length - 1 ? currentPage.items[index + 1] : null
    if (next) {
      await openMessage(next.id)
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
    if (!sessionId || !selectedMessageId) {
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      flagged: !(selectedMessage?.review?.flagged ?? false),
      tags: selectedMessage?.review?.tags || []
    })
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function clearFlag(): Promise<void> {
    if (!sessionId || !selectedMessageId) {
      return
    }
    await api.session.clearReview(sessionId, selectedMessageId)
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function addTag(tag: string): Promise<void> {
    if (!sessionId || !selectedMessageId) {
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      tags: Array.from(new Set([...(selectedMessage?.review?.tags || []), tag]))
    })
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function removeTag(tag: string): Promise<void> {
    if (!sessionId || !selectedMessageId) {
      return
    }
    await api.session.updateReview(sessionId, selectedMessageId, {
      tags: (selectedMessage?.review?.tags || []).filter((item) => item !== tag)
    })
    await refreshCurrentPage()
    await openMessage(selectedMessageId)
  }

  async function refreshSearchIndex(): Promise<void> {
    await api.pst.refreshSearchIndex()
    await api.hiddenFilters.list().then((result) => setHiddenRules(result.items || []))
    await refreshCurrentPage()
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
    if (!searchQuery.trim() && searchScope === 'pst') {
      await refreshCurrentPage()
      return
    }
    await loadSearchPage(1)
  }

  async function setMfaReminderSkipped(): Promise<void> {
    if (username) {
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
      isRemovedCatalog={isRemovedCatalog}
      catalogMessage={catalogMessage}
      caseOptions={caseSelectorOptions}
      selectedCasePath={selectedCasePath}
      selectedScopePath={selectedScopePath}
      searchOptions={searchSelectorOptions}
      catalogFiles={catalogFiles}
      selectedPstFileName={selectedPstFileName}
      onCaseChange={(value) => {
        const nextCasePath = normalizeText(value)
        const nextScopePath = getDefaultSearchPathForCase(caseOptions, nextCasePath)
        setSelectedCasePath(nextCasePath)
        setSelectedScopePath(nextScopePath)
        writeWorkspaceStorageItem('casePath', true, username, nextCasePath)
        writeWorkspaceStorageItem('scopePath', true, username, nextScopePath)
      }}
      onScopeChange={(value) => {
        const nextScopePath = value
        const nextCasePath = getCasePathFromScopePath(nextScopePath)
        setSelectedCasePath(nextCasePath)
        setSelectedScopePath(nextScopePath)
        writeWorkspaceStorageItem('casePath', true, username, nextCasePath)
        writeWorkspaceStorageItem('scopePath', true, username, nextScopePath)
      }}
      onCatalogModeToggle={() => {
        setIsRemovedCatalog((current) => !current)
      }}
      onRefreshCatalog={() => {
        void refreshCatalogRoot()
      }}
      onOpenMailbox={(fileName, scopePath) => {
        void openMailbox(fileName, scopePath)
      }}
      onRemoveMailbox={(fileName, scopePath) => {
        void moveMailbox(async () => {
          await api.pst.remove(scopePath, fileName)
        })
      }}
      onRestoreMailbox={(fileName, scopePath) => {
        void moveMailbox(async () => {
          await api.pst.restore(scopePath, fileName)
        })
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
      query={searchQuery}
      searchScope={searchScope}
      mailOnly={mailOnly}
      sort={sort}
      reviewFlaggedOnly={reviewFlaggedOnly}
      reviewTaggedOnly={reviewTaggedOnly}
      hiddenFiltersOpen={hiddenFiltersOpen}
      hiddenRules={hiddenRules}
      hiddenFiltersCount={hiddenRules.length}
      onQueryChange={setSearchQuery}
      onSearch={() => {
        void runSearch()
      }}
      onSearchScopeChange={(value) => setSearchScope(value)}
      onMailOnlyChange={(value) => setMailOnly(value)}
      onSortChange={(value) => setSort(value)}
      onReviewFlaggedChange={(value) => setReviewFlaggedOnly(value)}
      onReviewTaggedChange={(value) => setReviewTaggedOnly(value)}
      onToggleHiddenFilters={() => setHiddenFiltersOpen(!hiddenFiltersOpen)}
      onRemoveHiddenFilter={(filterId) => {
        void deleteHiddenFilter(filterId)
      }}
      onSelectMessage={(messageId) => {
        void openMessage(messageId)
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
          scopePath: selectedScopePath || selectedCasePath,
          sessionId,
          query: searchQuery,
          mailOnly,
          sort,
          reviewFlagged: reviewFlaggedOnly,
          reviewTagged: reviewTaggedOnly
        })
        triggerDownload(url, 'flagged-bundle.zip')
      }}
      onRefreshSearchIndex={() => {
        void refreshSearchIndex()
      }}
      onCreateHiddenFilter={(kind, value, label) => {
        void createHiddenFilter(kind, value, label)
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
      theme={theme}
      onDownloadJson={() => {
        void downloadJson()
      }}
      onDownloadEml={() => {
        void downloadEml()
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

  async function refreshCatalogRoot(): Promise<void> {
    if (!authenticated) {
      return
    }
    try {
      const response = await api.pst.catalog('', isRemovedCatalog)
      catalogLoadKeyRef.current = ''
      setCatalog(response)
      setCatalogMessage(response.message)
      setCaseOptions(response.scopes || [])
    } catch (error) {
      setCatalogMessage(error instanceof Error ? error.message : 'Unable to refresh catalog')
    }
  }

  async function moveMailbox(action: () => Promise<unknown>): Promise<void> {
    try {
      await action()
      setSelectedMessageId('')
      setSelectedMessage(null)
      setSessionId(null)
      setSessionSummary(null)
      setFolderTree(null)
      setCurrentFolderId('')
      setSelectedPstFileName('')
      setCurrentPage(null)
      setPageIndex(1)
      setWorkspaceMode('folder')
      setCatalogMessage('Mailbox moved. Refreshing catalog...')
      await refreshCatalogRoot()
    } catch (error) {
      setCatalogMessage(error instanceof Error ? error.message : 'Unable to update mailbox')
    }
  }

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
        open={fullViewOpen && Boolean(selectedMessage)}
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
              theme={theme}
              onDownloadJson={() => {
                void downloadJson()
              }}
              onDownloadEml={() => {
                void downloadEml()
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
        onCopyInvite={(value) => {
          const selected = users.find((item) => item.username === value)
          if (selected?.inviteUrl) {
            navigator.clipboard?.writeText(selected.inviteUrl).catch(() => undefined)
          }
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

  if (!authenticated || inviteFlowActive) {
    return (
      <div className="min-h-screen bg-[color:var(--bg)] text-[color:var(--text)]">
        <AuthScreen
          view={inviteFlowActive ? 'invite' : authView}
          busy={authBusy || inviteLoading || selfMfaLoading}
          message={authMessage}
          error={authError}
          invite={invite}
          inviteStep={inviteStep}
          inviteMfaAvailable={inviteMfaAvailable}
          inviteSetup={inviteSetup}
          inviteRecoveryCodes={inviteRecoveryCodes}
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
                                {user.mfaEnabled ? <Badge className="border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">MFA</Badge> : <Badge>MFA off</Badge>}
                                <Badge>{user.inviteStatus}</Badge>
                              </div>
                              <div className="mt-1 truncate text-xs text-[color:var(--muted)]">{user.recipientEmail}</div>
                              <div className="mt-1 text-xs text-[color:var(--soft)]">
                                {user.inviteStatus === 'pending' ? 'Pending invite' : user.inviteStatus === 'active' ? 'Active account' : user.inviteStatus}
                              </div>
                            </div>
                              <div className="flex shrink-0 items-center gap-1">
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
