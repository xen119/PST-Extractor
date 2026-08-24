import * as React from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { Panel, PanelGroup, PanelResizeHandle } from 'react-resizable-panels'
import {
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  Cloud,
  Download,
  Flag,
  FileText,
  Mail,
  Maximize2,
  LogOut,
  Plus,
  RefreshCw,
  Search,
  ScanLine,
  ShieldAlert,
  User,
  Tag as TagIcon
} from 'lucide-react'
import type {
  AttachmentDetail,
  CatalogEntry,
  FolderNode,
  MessageDetail,
  MessageSummary,
  SearchIndexRefreshSource,
  SearchIndexRefreshStatus,
  ReviewState
} from '@/types'
import { buildHtmlFrameSrcDoc, cn, escapeHtml, formatBytes, formatDate, getInitials } from '@/lib/utils'
import {
  Badge,
  Button,
  Dialog,
  DialogContent,
  DialogDescription,
  DialogTitle,
  IconButton,
  Input,
  Popover,
  PopoverContent,
  PopoverTrigger,
  ScrollArea,
  Separator
} from '@/components/ui'

const OFFICE_PREVIEW_CONTENT_TYPES = new Set([
  'application/msword',
  'application/vnd.ms-excel',
  'application/vnd.ms-powerpoint',
  'text/csv',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
])

function isOfficePreviewDocument(contentType?: string, fileName = ''): boolean {
  const normalizedType = (contentType || '').toLowerCase().split(';')[0].trim()
  if (normalizedType && OFFICE_PREVIEW_CONTENT_TYPES.has(normalizedType)) {
    return true
  }
  const extension = fileName.toLowerCase().split('.').pop() || ''
  return ['csv', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'].includes(extension)
}

function getMessagePreviewTitle(detail: MessageDetail | null): string {
  if (!detail) {
    return 'Message preview'
  }
  return (
    detail.archiveEntryName ||
    detail.downloadFilename ||
    detail.subject ||
    detail.archivePath ||
    detail.id ||
    'Message preview'
  )
}

export interface AppShellProps {
  userName: string
  authenticated: boolean
  settingsMenu: React.ReactNode
  breadcrumbs: Array<{ label: string; value?: string }>
  sidebarCollapsed: boolean
  previewCollapsed: boolean
  onLogout: () => void
  sidebar: React.ReactNode
  messagePanel: React.ReactNode
  preview: React.ReactNode
}

export function AppShell({
  userName,
  authenticated,
  settingsMenu,
  breadcrumbs,
  sidebarCollapsed,
  previewCollapsed,
  onLogout,
  sidebar,
  messagePanel,
  preview
}: AppShellProps) {
  const userChip = authenticated ? (
    <Popover>
      <span className="inline-flex items-stretch overflow-hidden rounded-full border border-[color:var(--line)] bg-[color:var(--surface-strong)] text-xs font-medium text-[color:var(--text)] shadow-sm">
        <PopoverTrigger asChild>
          <button
            type="button"
            className="inline-flex items-center gap-2 px-3 py-1.5 transition hover:bg-[color:var(--surface-soft)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]"
          >
            <User className="h-4 w-4 text-[color:var(--muted)]" aria-hidden="true" />
            <span>{userName || 'Signed in'}</span>
          </button>
        </PopoverTrigger>
        <span className="w-px self-stretch bg-[color:var(--line)]" aria-hidden="true" />
        <button
          type="button"
          aria-label="Logout"
          className="inline-flex items-center justify-center px-3 py-1.5 text-[color:var(--muted)] transition hover:bg-[color:var(--surface-soft)] hover:text-[color:var(--text)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]"
          onClick={onLogout}
        >
          <LogOut className="h-4 w-4" aria-hidden="true" />
        </button>
      </span>
      <PopoverContent className="w-72 p-3" align="end">
        {settingsMenu}
      </PopoverContent>
    </Popover>
  ) : null

  return (
    <div className="flex h-screen flex-col overflow-hidden">
      <header className="glass-bar sticky top-0 z-40 shrink-0">
        <div className="mx-auto flex w-full max-w-[1800px] items-center gap-3 px-4 py-3 lg:px-6">
          <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] text-[color:var(--accent)] shadow-sm">
            <Mail className="h-5 w-5" aria-hidden="true" />
          </div>
          <div className="min-w-0 flex-1 overflow-hidden">
            <div className="flex min-w-0 items-center gap-3">
              <div className="shrink-0 text-sm font-semibold tracking-[0.16em] text-[color:var(--muted)] uppercase">
                PST Mail Explorer
              </div>
              <div className="min-w-0 truncate text-xs leading-5 text-[color:var(--muted)]">
                {breadcrumbs.length ? (
                  breadcrumbs.map((crumb) => crumb.label).filter(Boolean).join(' / ')
                ) : (
                  'Select a case to begin.'
                )}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {userChip}
          </div>
        </div>
      </header>

      <main className="mx-auto flex w-full max-w-[1800px] flex-1 min-h-0 gap-3 overflow-hidden p-3 lg:p-4">
        <PanelGroup
          autoSaveId="pst-mail-explorer.layout"
          direction="horizontal"
          className="flex min-h-0 flex-1 gap-3 overflow-hidden"
        >
          <Panel
            defaultSize={22}
            minSize={16}
            collapsible
            collapsedSize={0}
            className={cn('min-h-0 overflow-hidden', sidebarCollapsed && 'hidden lg:block')}
          >
            {sidebar}
          </Panel>

          <PanelResizeHandle className="hidden w-1 rounded-full bg-transparent transition hover:bg-[color:var(--accent-soft)] lg:block" />

          <Panel defaultSize={previewCollapsed ? 100 : 42} minSize={28} className="min-h-0 overflow-hidden">
            {messagePanel}
          </Panel>

          <PanelResizeHandle className="hidden w-1 rounded-full bg-transparent transition hover:bg-[color:var(--accent-soft)] xl:block" />

          <Panel
            defaultSize={36}
            minSize={26}
            collapsible
            collapsedSize={0}
            className={cn('min-h-0 overflow-hidden', previewCollapsed && 'hidden xl:block')}
          >
            {preview}
          </Panel>
        </PanelGroup>
      </main>
    </div>
  )
}

interface SidebarProps {
  catalogMessage?: string
  caseOptions: Array<{ label: string; value: string; count: number }>
  selectedCasePath: string
  selectedScopePath: string
  sourceType: 'mailbox' | 'teams' | 'sharepoint'
  sourceCounts?: Record<'mailbox' | 'teams' | 'sharepoint', number> | null
  searchOptions: Array<{ label: string; value: string; count: number }>
  catalogFiles: CatalogEntry[]
  selectedPstFileName: string
  onCaseChange: (value: string) => void
  onScopeChange: (value: string) => void
  onSourceTypeChange: (value: 'mailbox' | 'teams' | 'sharepoint') => void
  onOpenAllItems?: () => void
  canRefreshSearchIndex: boolean
  searchIndexRefreshStatuses: Partial<Record<SearchIndexRefreshSource, SearchIndexRefreshStatus | null>>
  searchIndexRefreshBusyBySource: Partial<Record<SearchIndexRefreshSource, boolean>>
  onRefreshSearchIndex: (source: SearchIndexRefreshSource) => void
  onOpenMailbox: (fileName: string, scopePath: string) => void
  folderTree: FolderNode | null
  currentFolderId: string
  onSelectFolder: (folderId: string) => void
  summary?: MessageDetail | null
}

export function Sidebar({
  catalogMessage,
  caseOptions,
  selectedCasePath,
  selectedScopePath,
  sourceType,
  sourceCounts,
  searchOptions,
  catalogFiles,
  selectedPstFileName,
  onCaseChange,
  onScopeChange,
  onSourceTypeChange,
  onOpenAllItems,
  canRefreshSearchIndex,
  searchIndexRefreshStatuses,
  searchIndexRefreshBusyBySource,
  onRefreshSearchIndex,
  onOpenMailbox,
  folderTree,
  currentFolderId,
  onSelectFolder
}: SidebarProps) {
  const selectedMailbox = catalogFiles.find((file) => file.fileName === selectedPstFileName) || null
  const tabCounts = sourceCounts || null
  const mailboxesRefreshStatus = searchIndexRefreshStatuses.mailboxes || null
  const itemsRefreshStatus = searchIndexRefreshStatuses.items || null
  const mailboxesRefreshBusy = Boolean(searchIndexRefreshBusyBySource.mailboxes)
  const itemsRefreshBusy = Boolean(searchIndexRefreshBusyBySource.items)

  const sourceTabs: Array<{
    key: 'mailbox' | 'teams' | 'sharepoint'
    label: string
    icon: React.ReactNode
  }> = [
    {
      key: 'mailbox',
      label: 'Mailbox',
      icon: <Mail className="h-4 w-4" />
    },
    {
      key: 'teams',
      label: 'Teams',
      icon: <FileText className="h-4 w-4" />
    },
    {
      key: 'sharepoint',
      label: 'SharePoint/OneDrive',
      icon: <Cloud className="h-4 w-4" />
    }
  ]

  return (
    <div className="panel-surface flex h-full min-h-0 flex-col overflow-hidden">
      <div className="panel-heading">
        <div>
          <div className="panel-title">Case / Search</div>
          <div className="text-sm text-[color:var(--muted)]">Browse and open a mailbox</div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {canRefreshSearchIndex ? (
            <>
              {[
                {
                  key: 'mailboxes' as const,
                  label: 'Mailboxes',
                  status: mailboxesRefreshStatus,
                  busy: mailboxesRefreshBusy
                },
                {
                  key: 'items' as const,
                  label: 'Items',
                  status: itemsRefreshStatus,
                  busy: itemsRefreshBusy
                }
              ].map((entry) => (
                <div key={entry.key} className="flex items-center gap-2">
                  {entry.status?.status === 'running' ? (
                    <span className="inline-flex items-center gap-2 rounded-full border border-[color:var(--line)] bg-[color:var(--surface-soft)] px-3 py-1 text-xs font-medium text-[color:var(--muted)]">
                      <RefreshCw className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />
                      Reindexing {entry.label.toLowerCase()}
                    </span>
                  ) : entry.status?.status === 'failed' ? (
                    <span className="inline-flex items-center gap-2 rounded-full border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-3 py-1 text-xs font-medium text-[color:var(--danger)]">
                      <ShieldAlert className="h-3.5 w-3.5" aria-hidden="true" />
                      {entry.label} reindex failed
                    </span>
                  ) : null}
                  <button
                    type="button"
                    className="inline-flex items-center gap-2 rounded-full border border-[color:var(--line)] bg-[color:var(--surface-strong)] px-3 py-1.5 text-xs font-medium text-[color:var(--text)] transition hover:bg-[color:var(--surface-soft)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)] disabled:cursor-not-allowed disabled:opacity-60"
                    disabled={entry.busy || entry.status?.status === 'running'}
                    onClick={() => onRefreshSearchIndex(entry.key)}
                  >
                    <RefreshCw className={cn('h-3.5 w-3.5', entry.status?.status === 'running' && 'animate-spin')} aria-hidden="true" />
                    Reindex {entry.label.toLowerCase()}
                  </button>
                </div>
              ))}
            </>
          ) : null}
          {onOpenAllItems ? (
            <button
              type="button"
              className="inline-flex items-center gap-2 rounded-full border border-[color:var(--line)] bg-[color:var(--surface-strong)] px-3 py-1.5 text-xs font-medium text-[color:var(--text)] transition hover:bg-[color:var(--surface-soft)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]"
              onClick={onOpenAllItems}
            >
              <ScanLine className="h-3.5 w-3.5" aria-hidden="true" />
              All items
            </button>
          ) : null}
        </div>
      </div>
      {[mailboxesRefreshStatus, itemsRefreshStatus].find((status) => status?.status === 'failed' && status.error) ? (
        <div className="border-t border-[color:var(--line)] px-4 py-3 text-xs leading-5 text-[color:var(--danger)]">
          {mailboxesRefreshStatus?.status === 'failed' && mailboxesRefreshStatus.error
            ? mailboxesRefreshStatus.error
            : itemsRefreshStatus?.status === 'failed' && itemsRefreshStatus.error
              ? itemsRefreshStatus.error
              : null}
        </div>
      ) : null}

      <ScrollArea className="min-h-0 flex-1">
        <div className="flex min-h-0 flex-col">
          <div className="space-y-3 border-t border-[color:var(--line)] px-4 py-4">
            <div className="space-y-2">
              <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">Source</div>
              <div className="grid grid-cols-3 gap-2">
                {sourceTabs.map((tab) => {
                  const active = sourceType === tab.key
                  const count = tabCounts ? tabCounts[tab.key] : '…'
                  return (
                    <button
                      key={tab.key}
                      type="button"
                      className={cn(
                        'inline-flex items-center justify-center gap-2 rounded-2xl border px-3 py-2 text-xs font-semibold transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]',
                        active
                          ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)] text-[color:var(--accent)]'
                          : 'border-[color:var(--line)] bg-[color:var(--surface-strong)] text-[color:var(--text)] hover:border-[color:var(--accent-soft)] hover:bg-[color:var(--surface-soft)]'
                      )}
                      onClick={() => onSourceTypeChange(tab.key)}
                    >
                      {tab.icon}
                      <span className="truncate">{tab.label}</span>
                      <span className="rounded-full border border-[color:var(--line)] bg-[color:var(--surface)] px-2 py-0.5 text-[10px] font-medium text-[color:var(--muted)]">
                        {count}
                      </span>
                    </button>
                  )
                })}
              </div>
            </div>

            <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
              Case
              <select
                className="select mt-2"
                value={selectedCasePath || ''}
                onChange={(event) => onCaseChange(event.target.value)}
              >
                <option value="">{caseOptions.length ? 'Select a case' : 'No cases found'}</option>
                {caseOptions.length ? (
                  caseOptions.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))
                ) : null}
              </select>
            </label>

            <label className="block text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
              Search
              <select
                className="select mt-2"
                value={selectedScopePath || ''}
                disabled={!selectedCasePath || !searchOptions.length}
                onChange={(event) => onScopeChange(event.target.value)}
              >
                <option value="">{selectedCasePath ? 'Select a search' : 'Select a case first'}</option>
                {searchOptions.length ? (
                  searchOptions.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))
                ) : null}
              </select>
            </label>
            {catalogMessage ? <div className="text-xs leading-5 text-[color:var(--muted)]">{catalogMessage}</div> : null}
          </div>

          <Separator />

          <div className="flex min-h-0 flex-1 flex-col">
            {sourceType === 'mailbox' ? (
              <>
                <div className="px-4 py-3">
                  <div className="flex items-center justify-between">
                    <div className="text-sm font-semibold text-[color:var(--text)]">Mailboxes</div>
                    <Badge>{catalogFiles.length}</Badge>
                  </div>
                  {catalogFiles.length ? (
                    <div className="mt-2 flex items-center gap-2">
                      <select
                        className="select flex-1"
                        aria-label="Mailbox selector"
                        value={selectedMailbox?.fileName || ''}
                        disabled={!catalogFiles.length}
                        onChange={(event) => {
                          const nextMailbox = catalogFiles.find((file) => file.fileName === event.target.value)
                          if (nextMailbox) {
                            onOpenMailbox(nextMailbox.fileName, nextMailbox.scopePath || '')
                          }
                        }}
                      >
                        <option value="">{selectedCasePath ? 'Select a mailbox' : 'Select a case first'}</option>
                        {catalogFiles.map((file) => (
                          <option key={`${file.scopePath || ''}/${file.fileName}`} value={file.fileName}>
                            {file.fileName}
                          </option>
                        ))}
                      </select>
                    </div>
                  ) : (
                    <div className="mt-2 empty-state min-h-[112px]">No PST files are available in the current scope.</div>
                  )}
                </div>

                <Separator />

                <div className="flex min-h-0 flex-1 flex-col">
                  <div className="panel-heading">
                    <div>
                      <div className="panel-title">Folders</div>
                      <div className="text-sm text-[color:var(--muted)]">Navigate mailbox structure</div>
                    </div>
                    <Badge>{folderTree ? 'loaded' : 'empty'}</Badge>
                  </div>
                  <div className="min-h-0 flex-1 px-4 pb-4 pt-1">
                    {folderTree ? (
                      <FolderList node={folderTree} currentFolderId={currentFolderId} onSelectFolder={onSelectFolder} />
                    ) : (
                      <div className="empty-state">Open a PST to see folders.</div>
                    )}
                  </div>
                </div>
              </>
            ) : (
              <div className="px-4 pb-4 pt-3">
                <div className="empty-state min-h-[180px]">
                  Search results in {sourceType === 'teams' ? 'Teams' : 'SharePoint/OneDrive'} will appear here.
                </div>
              </div>
            )}
          </div>
        </div>
      </ScrollArea>
    </div>
  )
}

interface FolderListProps {
  node: FolderNode
  currentFolderId: string
  onSelectFolder: (folderId: string) => void
  depth?: number
}

function FolderList({ node, currentFolderId, onSelectFolder, depth = 0 }: FolderListProps) {
  const childNodes = (Array.isArray(node.children) ? node.children : [])
    .map((child) => (
      <FolderList
        key={child.id}
        node={child}
        currentFolderId={currentFolderId}
        onSelectFolder={onSelectFolder}
        depth={depth + 1}
      />
    ))
    .filter(Boolean)

  const totalCount = node.indexedMessageCount || 0
  if (totalCount <= 0 && childNodes.length === 0) {
    return null
  }

  const active = node.id === currentFolderId
  return (
    <div className="space-y-1" style={{ paddingLeft: `${depth * 12}px` }}>
      <button
        type="button"
        className={cn(
          'flex w-full items-center justify-between rounded-2xl border px-3 py-2 text-left transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]',
          active
            ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)]'
            : 'border-[color:var(--line)] bg-[color:var(--surface-strong)] hover:border-[color:var(--accent-soft)] hover:bg-[color:var(--surface-soft)]'
        )}
        onClick={() => onSelectFolder(node.id)}
      >
        <span className="min-w-0 truncate text-sm font-medium">{node.displayName || '(untitled)'}</span>
        {totalCount > 0 ? <span className="shrink-0 whitespace-nowrap text-sm text-[color:var(--soft)]">- {totalCount}</span> : null}
      </button>
      {childNodes.length ? childNodes : null}
    </div>
  )
}

interface MessageListProps {
  page: {
    items: MessageSummary[]
    total: number
    page: number
    totalPages: number
    query: string
    folder?: { displayName?: string; path?: string }
    scopeLabel?: string
    reviewFilters?: { flaggedOnly: boolean; taggedOnly: boolean; tag: string }
  } | null
  loading: boolean
  query: string
  activeQuery?: string
  sourceType: 'mailbox' | 'teams' | 'sharepoint'
  searchScope: 'pst' | 'search' | 'all'
  mailOnly: boolean
  sort: string
  reviewFlaggedOnly: boolean
  reviewTaggedOnly: boolean
  onQueryChange: (value: string) => void
  onSearch: () => void
  onSearchScopeChange: (value: 'pst' | 'search' | 'all') => void
  onMailOnlyChange: (value: boolean) => void
  onSortChange: (value: string) => void
  onReviewFlaggedChange: (value: boolean) => void
  onReviewTaggedChange: (value: boolean) => void
  onSelectMessage: (message: MessageSummary) => void
  onPrevPage: () => void
  onNextPage: () => void
  onOpenBundle: () => void
  onPageChange?: (page: number) => void
  selectedMessageId: string
  sessionId: string | null
  emptyStateTitle?: string
  emptyStateDescription?: string
}

export function MessageList({
  page,
  loading,
  query,
  activeQuery,
  sourceType,
  searchScope,
  mailOnly,
  sort,
  reviewFlaggedOnly,
  reviewTaggedOnly,
  onQueryChange,
  onSearch,
  onSearchScopeChange,
  onMailOnlyChange,
  onSortChange,
  onReviewFlaggedChange,
  onReviewTaggedChange,
  onSelectMessage,
  onPrevPage,
  onNextPage,
  onOpenBundle,
  selectedMessageId,
  sessionId,
  emptyStateTitle,
  emptyStateDescription
}: MessageListProps) {
  const parentRef = React.useRef<HTMLDivElement | null>(null)
  const rows = page?.items || []
  const virtualizer = useVirtualizer({
    count: rows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 96,
    overscan: 8
  })

  const activePage = page?.page ?? 0
  const totalPages = page?.totalPages ?? 0

  return (
    <div className="panel-surface flex h-full min-h-0 flex-col overflow-hidden">
      <div className="panel-heading">
        <div>
          <div className="panel-title">Messages</div>
          <div className="text-sm text-[color:var(--muted)]">
            {page
              ? `${page.total} results in ${page.folder?.displayName || page.scopeLabel || 'folder'}`
              : 'Select a folder to load messages'}
          </div>
        </div>
        <div className="flex items-center gap-2">
          {sourceType === 'mailbox' ? (
            <IconButton label="Download flagged bundle" onClick={onOpenBundle}>
              <Download className="h-4 w-4" />
            </IconButton>
          ) : null}
        </div>
      </div>

            <div className="flex min-h-0 flex-1 flex-col">
        <div className="border-t border-[color:var(--line)] px-4 py-4">
          <MessageSearchBar
            query={query}
            activeQuery={activeQuery}
            sourceType={sourceType}
            searchScope={searchScope}
            mailOnly={mailOnly}
            sort={sort}
            reviewFlaggedOnly={reviewFlaggedOnly}
            reviewTaggedOnly={reviewTaggedOnly}
            onQueryChange={onQueryChange}
            onSearch={onSearch}
            onSearchScopeChange={onSearchScopeChange}
            onMailOnlyChange={onMailOnlyChange}
            onSortChange={onSortChange}
            onReviewFlaggedChange={onReviewFlaggedChange}
            onReviewTaggedChange={onReviewTaggedChange}
          />
        </div>
        <div className="flex min-h-0 flex-1 flex-col border-t border-[color:var(--line)]">
          <div className="flex items-center justify-between border-b border-[color:var(--line)] px-4 py-2 text-xs text-[color:var(--muted)]">
            <div className="flex items-center gap-2">
              <Button variant="ghost" className="px-2 py-1 text-xs" onClick={onPrevPage} disabled={!page || activePage <= 1}>
                <ChevronLeft className="mr-1 h-4 w-4" />
                Prev
              </Button>
              <Button variant="ghost" className="px-2 py-1 text-xs" onClick={onNextPage} disabled={!page || activePage >= totalPages}>
                Next
                <ChevronRight className="ml-1 h-4 w-4" />
              </Button>
            </div>
            <span>
              Page {activePage} of {totalPages}
            </span>
          </div>
          <ScrollArea ref={parentRef} className="min-h-0 flex-1">
            {loading ? (
              <div className="empty-state m-4">Loading messages...</div>
            ) : page ? (
              rows.length ? (
                <div
                  className="relative w-full"
                  style={{
                    height: `${virtualizer.getTotalSize()}px`
                  }}
                >
                  {virtualizer.getVirtualItems().map((virtualRow) => {
                    const item = rows[virtualRow.index]
                    return (
                      <div
                        key={item.id}
                        ref={virtualizer.measureElement}
                        data-index={virtualRow.index}
                        className="absolute left-0 top-0 w-full px-4 py-2"
                        style={{
                          transform: `translateY(${virtualRow.start}px)`
                        }}
                      >
                        <MessageRow
                          item={item}
                          active={item.id === selectedMessageId}
                          onSelect={() => onSelectMessage(item)}
                        />
                      </div>
                    )
                  })}
                </div>
              ) : (
                <div className="empty-state m-4">No messages match the current filters.</div>
              )
            ) : (
              <div className="empty-state m-4">
                <div className="text-base font-semibold text-[color:var(--text)]">
                  {emptyStateTitle || 'Select a folder to begin.'}
                </div>
                {emptyStateDescription ? <div>{emptyStateDescription}</div> : null}
              </div>
            )}
          </ScrollArea>
        </div>
      </div>
    </div>
  )
}

function MessageSearchBar(props: {
  query: string
  activeQuery?: string
  sourceType: 'mailbox' | 'teams' | 'sharepoint'
  searchScope: 'pst' | 'search' | 'all'
  mailOnly: boolean
  sort: string
  reviewFlaggedOnly: boolean
  reviewTaggedOnly: boolean
  onQueryChange: (value: string) => void
  onSearch: () => void
  onSearchScopeChange: (value: 'pst' | 'search' | 'all') => void
  onMailOnlyChange: (value: boolean) => void
  onSortChange: (value: string) => void
  onReviewFlaggedChange: (value: boolean) => void
  onReviewTaggedChange: (value: boolean) => void
}) {
  const effectiveSearchScope = props.sourceType === 'mailbox' ? props.searchScope : 'search'
  const displayedQuery = props.activeQuery ?? props.query
  const scopeOptions =
    props.sourceType === 'mailbox'
      ? [
          { value: 'pst', label: 'Selected PST' },
          { value: 'search', label: 'Selected search' },
          { value: 'all', label: 'All cases/searches' }
        ]
      : [{ value: 'search', label: 'Selected search' }]

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-end gap-2">
        <label className="flex min-w-[18rem] flex-1 flex-col gap-1 text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
          Search
          <div className="relative">
            <Input
              value={props.query}
              onChange={(event) => props.onQueryChange(event.target.value)}
              placeholder='Keywords, "phrases", + AND, | OR'
            />
            <IconButton
              label="Run search"
              className="absolute right-1 top-1 h-8 w-8"
              onClick={props.onSearch}
            >
              <Search className="h-4 w-4" />
            </IconButton>
          </div>
        </label>

        <label className="flex w-48 flex-col gap-1 text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
          Scope
          <select
            className="select"
            value={effectiveSearchScope}
            onChange={(event) => props.onSearchScopeChange(event.target.value as 'pst' | 'search' | 'all')}
          >
            {scopeOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>

        <label className="flex w-44 flex-col gap-1 text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
          Sort
          <select
            className="select"
            value={props.sort}
            onChange={(event) => props.onSortChange(event.target.value)}
          >
            <option value="date-desc">Newest first</option>
            <option value="order">Folder order</option>
          </select>
        </label>

        <div className="flex items-center gap-2">
          <IconButton
            label={props.mailOnly ? 'Show mail and non-mail items' : 'Show mail only'}
            className={props.mailOnly ? 'icon-button-primary' : ''}
            onClick={() => props.onMailOnlyChange(!props.mailOnly)}
          >
            <Mail className="h-4 w-4" />
          </IconButton>
          <IconButton
            label={props.reviewFlaggedOnly ? 'Clear flagged filter' : 'Filter flagged items'}
            className={props.reviewFlaggedOnly ? 'icon-button-primary' : ''}
            onClick={() => props.onReviewFlaggedChange(!props.reviewFlaggedOnly)}
          >
            <Flag className="h-4 w-4" />
          </IconButton>
          <IconButton
            label={props.reviewTaggedOnly ? 'Clear tagged filter' : 'Filter tagged items'}
            className={props.reviewTaggedOnly ? 'icon-button-primary' : ''}
            onClick={() => props.onReviewTaggedChange(!props.reviewTaggedOnly)}
          >
            <TagIcon className="h-4 w-4" />
          </IconButton>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        {props.mailOnly ? <span className="chip chip-active">Mail only</span> : null}
        {props.reviewFlaggedOnly ? <span className="chip chip-active">Flagged</span> : null}
        {props.reviewTaggedOnly ? <span className="chip chip-active">Tagged</span> : null}
        {displayedQuery ? <span className="chip">Search: {displayedQuery}</span> : null}
      </div>
    </div>
  )
}

interface MessageRowProps {
  item: MessageSummary
  active: boolean
  onSelect: () => void
}

export function MessageRow({ item, active, onSelect }: MessageRowProps) {
  const sender = item.senderName || item.senderEmailAddress || '(unknown sender)'
  const snippet = item.recipientText || item.displayTo || item.folderPath || item.scopeLabel || ''
  const review = item.review
  const chips = [
    item.sourceType && item.sourceType !== 'mailbox' ? (
      <span key="source" className="chip chip-active">
        {item.sourceType === 'teams' ? 'Teams' : 'SharePoint/OneDrive'}
      </span>
    ) : null,
    item.kind ? <span key="kind" className="chip">{item.kind}</span> : null,
    item.hasAttachments ? <span key="attachments" className="chip chip-active">Attachments</span> : null,
    item.isRead ? <span key="read" className="chip">Read</span> : <span key="unread" className="chip chip-active">Unread</span>,
    review?.flagged ? <span key="flagged" className="chip chip-active">Flagged</span> : null,
    review?.tags?.length ? <span key="tagged" className="chip chip-active">{review.tags.join(', ')}</span> : null
  ].filter(Boolean)

  return (
    <button
      type="button"
      onClick={onSelect}
      onKeyDown={(event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault()
          onSelect()
        }
      }}
      className={cn(
        'w-full rounded-2xl border px-4 py-3 text-left transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]',
        active
          ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)]'
          : 'border-[color:var(--line)] bg-[color:var(--surface-strong)] hover:border-[color:var(--accent-soft)] hover:bg-[color:var(--surface-soft)]'
      )}
    >
      <div className="flex items-start justify-between gap-4">
        <div className="min-w-0">
          <div className="truncate text-[0.95rem] font-semibold text-[color:var(--text)]">
            {item.subject || '(no subject)'}
          </div>
          <div className="mt-1 truncate text-sm text-[color:var(--muted)]">{sender}</div>
        </div>
        <div className="text-xs text-[color:var(--soft)]">{formatDate(item.sortDate || item.creationTime || item.clientSubmitTime)}</div>
      </div>
      <div className="mt-2 line-clamp-2 text-sm text-[color:var(--muted)]">{snippet}</div>
      <div className="mt-3 flex flex-wrap items-center gap-2">{chips}</div>
    </button>
  )
}

interface EmailPreviewProps {
  detail: MessageDetail | null
  loading?: boolean
  theme: 'light' | 'dark'
  onDownloadJson?: () => void
  onDownloadEml?: () => void
  onDownloadItem?: () => void
  onToggleFlag?: () => void
  onClearReview?: () => void
  onOpenTags?: () => void
  onOpenFullView?: () => void
  onOpenAttachment: (attachment: AttachmentDetail) => void
  onOpenPrev: () => void
  onOpenNext: () => void
  showNavigationControls?: boolean
  tagCount: number
  canNavigatePrev: boolean
  canNavigateNext: boolean
  emptyStateTitle?: string
  emptyStateDescription?: string
}

export function EmailPreview({
  detail,
  loading = false,
  theme,
  onDownloadJson,
  onDownloadEml,
  onDownloadItem,
  onToggleFlag,
  onClearReview,
  onOpenTags,
  onOpenFullView,
  onOpenAttachment,
  onOpenPrev,
  onOpenNext,
  showNavigationControls = true,
  tagCount,
  canNavigatePrev,
  canNavigateNext,
  emptyStateTitle,
  emptyStateDescription
}: EmailPreviewProps) {
  if (!detail) {
    return (
      <div className="panel-surface flex h-full min-h-0 items-center justify-center p-6">
        <div className="empty-state">
          {loading ? <RefreshCw className="h-5 w-5 animate-spin text-[color:var(--accent)]" aria-hidden="true" /> : null}
          <div className="text-base font-semibold text-[color:var(--text)]">
            {loading ? 'Loading preview...' : emptyStateTitle || 'No message selected'}
          </div>
          <div>
            {loading
              ? 'Fetching the selected item.'
              : emptyStateDescription || 'Select a message from the list to preview it.'}
          </div>
        </div>
      </div>
    )
  }

  const senderName = detail.senderName || detail.senderEmailAddress || 'Unknown sender'
  const senderEmail =
    detail.senderEmailAddress && detail.senderEmailAddress !== senderName ? detail.senderEmailAddress : ''
  const sentTime = formatDate(detail.sortDate || detail.clientSubmitTime || detail.creationTime)
  const bodyHtml = detail.bodyHtml || ''
  const bodyText = detail.bodyText || detail.bodyPrefix || detail.parseError || ''
  const hasHtml = Boolean(bodyHtml.trim())
  const bodyFrame = hasHtml ? buildHtmlFrameSrcDoc(bodyHtml, theme === 'dark') : ''
  const isArchiveItem = Boolean(detail.archivePath)
  const downloadUrl = detail.downloadUrl || ''
  const previewUrl = detail.previewUrl || ''
  const isImagePreview = Boolean(detail.contentType && detail.contentType.startsWith('image/'))
  const isPdfPreview = Boolean(detail.contentType === 'application/pdf')
  const isOfficePreview = Boolean(
    isArchiveItem &&
      previewUrl &&
      isOfficePreviewDocument(detail.contentType, detail.downloadFilename || detail.archiveEntryName || detail.subject || '')
  )

  return (
    <div className="panel-surface flex h-full min-h-0 flex-col overflow-hidden">
      <div className="panel-heading border-b border-[color:var(--line)]">
        <div>
          <div className="panel-title">{isArchiveItem ? 'Document preview' : 'Reading pane'}</div>
          <div className="text-sm text-[color:var(--muted)]">
            {isArchiveItem ? 'Teams or document preview' : 'Message details and preview'}
          </div>
        </div>
        {showNavigationControls ? (
          <div className="flex items-center gap-2">
            <IconButton label="Previous message" onClick={onOpenPrev} disabled={!canNavigatePrev}>
              <ChevronLeft className="h-4 w-4" />
            </IconButton>
            <IconButton label="Next message" onClick={onOpenNext} disabled={!canNavigateNext}>
              <ChevronRight className="h-4 w-4" />
            </IconButton>
          </div>
        ) : null}
      </div>

      <ScrollArea className="min-h-0 flex-1">
        <article className="space-y-4 p-4">
          <div className="panel-surface-strong overflow-hidden">
          <EmailHeader
            detail={detail}
            senderName={senderName}
            senderEmail={senderEmail}
            sentTime={sentTime}
            tagCount={tagCount}
            onOpenTags={onOpenTags}
            onOpenFullView={onOpenFullView}
            onDownloadJson={isArchiveItem ? undefined : onDownloadJson}
            onDownloadEml={isArchiveItem ? undefined : onDownloadEml}
            onDownloadItem={isArchiveItem ? onDownloadItem : undefined}
            onToggleFlag={onToggleFlag}
            onClearReview={onClearReview}
          />
          </div>

          {detail.parseError ? (
            <div className="rounded-2xl border border-[color:rgba(245,158,11,0.25)] bg-[color:var(--warning-bg)] px-4 py-3 text-sm text-[color:var(--warning)]">
              <strong className="font-semibold">Load warning:</strong> {detail.parseError}
            </div>
          ) : null}

          <div className="panel-surface-strong overflow-hidden">
            <div className="border-b border-[color:var(--line)] px-4 py-3 text-sm font-semibold text-[color:var(--text)]">
              Message body
            </div>
            <div className="p-4">
              {hasHtml ? (
                <iframe
                  className="h-[42rem] w-full rounded-2xl border border-[color:var(--line)] bg-white"
                  sandbox="allow-popups"
                  srcDoc={bodyFrame}
                  title="Message body"
                />
              ) : isOfficePreview && previewUrl ? (
                <iframe
                  className="h-[42rem] w-full rounded-2xl border border-[color:var(--line)] bg-white"
                  src={previewUrl}
                  title="Document preview"
                />
              ) : isImagePreview && downloadUrl ? (
                <img
                  alt={getMessagePreviewTitle(detail)}
                  className="max-h-[42rem] w-auto max-w-full rounded-2xl border border-[color:var(--line)] object-contain"
                  src={downloadUrl}
                />
              ) : isPdfPreview && downloadUrl ? (
                <iframe
                  className="h-[42rem] w-full rounded-2xl border border-[color:var(--line)] bg-white"
                  src={downloadUrl}
                  title="Document preview"
                />
              ) : (
                <pre className="whitespace-pre-wrap rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4 text-sm leading-7 text-[color:var(--text)]">
                  {bodyText || 'No message body is available.'}
                </pre>
              )}
              {isArchiveItem && !hasHtml && !bodyText && downloadUrl && !isOfficePreview && !isImagePreview && !isPdfPreview ? (
                <div className="mt-3 flex items-center justify-between gap-3 rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] px-4 py-3 text-sm text-[color:var(--muted)]">
                  <span>No inline preview is available for this file type.</span>
                  {onDownloadItem ? (
                    <Button variant="secondary" onClick={onDownloadItem}>
                      Download file
                    </Button>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>

          {Array.isArray(detail.attachments) && detail.attachments.length ? (
            <div className="panel-surface-strong overflow-hidden">
              <div className="border-b border-[color:var(--line)] px-4 py-3 text-sm font-semibold text-[color:var(--text)]">
                Attachments
              </div>
              <div className="space-y-3 p-4">
                {detail.attachments.map((attachment) => (
                  <AttachmentCard key={`${attachment.index}-${attachment.attachmentId}`} attachment={attachment} onOpen={onOpenAttachment} />
                ))}
              </div>
            </div>
          ) : null}
        </article>
      </ScrollArea>
    </div>
  )
}

function EmailHeader({
  detail,
  senderName,
  senderEmail,
  sentTime,
  tagCount,
  onOpenTags,
  onOpenFullView,
  onDownloadJson,
  onDownloadEml,
  onDownloadItem,
  onToggleFlag,
  onClearReview
}: {
  detail: MessageDetail
  senderName: string
  senderEmail: string
  sentTime: string
  tagCount: number
  onOpenTags?: () => void
  onOpenFullView?: () => void
  onDownloadJson?: () => void
  onDownloadEml?: () => void
  onDownloadItem?: () => void
  onToggleFlag?: () => void
  onClearReview?: () => void
}) {
  const recipients = [
    detail.displayTo || detail.resolvedDisplayTo,
    detail.displayCC || detail.resolvedDisplayCC,
    detail.displayBCC || detail.resolvedDisplayBCC
  ].filter(Boolean)
  const isArchiveItem = Boolean(detail.archivePath)

  return (
    <div className="p-4">
      <div className="flex items-start justify-between gap-4">
        <div className="min-w-0">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-2xl bg-[color:var(--accent-soft)] text-sm font-semibold text-[color:var(--accent-strong)]">
              {getInitials(senderName)}
            </div>
            <div className="min-w-0">
              <div className="truncate text-lg font-semibold text-[color:var(--text)]">
                {detail.subject || '(no subject)'}
              </div>
              <div className="mt-1 flex flex-wrap items-center gap-2 text-sm text-[color:var(--muted)]">
                <span className="font-medium text-[color:var(--text)]">{senderName}</span>
                {senderEmail ? <span>&lt;{senderEmail}&gt;</span> : null}
              </div>
            </div>
          </div>
          <div className="mt-4 flex flex-wrap items-center gap-2 text-xs text-[color:var(--muted)]">
            <Badge>{sentTime}</Badge>
            {detail.review?.flagged ? <Badge className="border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">Flagged</Badge> : null}
            {(detail.review?.tags || []).map((tag) => (
              <Badge key={tag}>{tag}</Badge>
            ))}
          </div>
        </div>

        <div className="flex shrink-0 items-center gap-2">
          <Popover>
            <PopoverTrigger asChild>
              <IconButton label="Recipients">
                <ChevronDown className="h-4 w-4" />
              </IconButton>
            </PopoverTrigger>
            <PopoverContent className="w-[22rem]">
              <div className="space-y-3">
                <div className="text-sm font-semibold text-[color:var(--text)]">Recipients</div>
                {recipients.length ? (
                  recipients.map((value, index) => (
                    <div key={index} className="rounded-2xl border border-[color:var(--line)] px-3 py-2 text-sm text-[color:var(--muted)]">
                      {value}
                    </div>
                  ))
                ) : (
                  <div className="empty-state">No recipient details available.</div>
                )}
              </div>
            </PopoverContent>
          </Popover>
          {onOpenFullView ? (
            <IconButton label="Open full view" onClick={onOpenFullView}>
              <Maximize2 className="h-4 w-4" />
            </IconButton>
          ) : null}
          <div className="flex items-center gap-2">
            <Badge>{tagCount}</Badge>
            {onOpenTags ? (
              <IconButton label="Manage tags" onClick={onOpenTags}>
                <TagIcon className="h-4 w-4" />
              </IconButton>
            ) : null}
          </div>
          {isArchiveItem && onDownloadItem ? (
            <IconButton label="Download file" onClick={onDownloadItem}>
              <Download className="h-4 w-4" />
            </IconButton>
          ) : null}
          {!isArchiveItem ? (
            <>
              {onDownloadJson ? (
                <IconButton label="Download JSON" onClick={onDownloadJson}>
                  <Download className="h-4 w-4" />
                </IconButton>
              ) : null}
              {onDownloadEml ? (
                <IconButton label="Download EML" onClick={onDownloadEml}>
                  <Download className="h-4 w-4" />
                </IconButton>
              ) : null}
            </>
          ) : null}
          {detail.review?.flagged ? (
            onClearReview ? (
              <IconButton label="Clear flag" className="icon-button-danger" onClick={onClearReview}>
                <Flag className="h-4 w-4" />
              </IconButton>
            ) : null
          ) : onToggleFlag ? (
            <IconButton label="Flag message" onClick={onToggleFlag}>
              <Flag className="h-4 w-4" />
            </IconButton>
          ) : null}
        </div>
      </div>
    </div>
  )
}

export function TagManagerDialog({
  open,
  tags,
  subject,
  onOpenChange,
  onAddTag,
  onRemoveTag
}: {
  open: boolean
  tags: string[]
  subject?: string
  onOpenChange: (open: boolean) => void
  onAddTag: (tag: string) => void
  onRemoveTag: (tag: string) => void
}) {
  const [value, setValue] = React.useState('')

  React.useEffect(() => {
    if (!open) {
      setValue('')
    }
  }, [open])

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="!w-[min(92vw,720px)]">
        <div className="space-y-5 p-6">
          <div className="flex items-start justify-between gap-4">
            <div className="min-w-0">
              <DialogTitle>Tags</DialogTitle>
              <DialogDescription className="mt-1">
                {subject || 'Manage tags for the selected message.'}
              </DialogDescription>
            </div>
            <Badge>{tags.length}</Badge>
          </div>

          <Separator />

          <div className="space-y-2">
            {tags.length ? (
              <div className="flex flex-wrap gap-2">
                {tags.map((tag) => (
                  <span key={tag} className="chip chip-active">
                    <span>{tag}</span>
                    <button
                      type="button"
                      className="ml-1 rounded-full p-0.5 text-[color:var(--accent-strong)] hover:bg-white/50"
                      onClick={() => onRemoveTag(tag)}
                      aria-label={`Remove ${tag}`}
                    >
                      ×
                    </button>
                  </span>
                ))}
              </div>
            ) : (
              <div className="empty-state py-6">No tags yet.</div>
            )}
          </div>

          <form
            className="flex items-center gap-2"
            onSubmit={(event) => {
              event.preventDefault()
              const next = value.trim()
              if (!next) {
                return
              }
              onAddTag(next)
              setValue('')
            }}
          >
            <Input value={value} onChange={(event) => setValue(event.target.value)} placeholder="Add tag" />
            <Button type="submit" variant="secondary">
              <Plus className="h-4 w-4" />
              Add
            </Button>
          </form>
        </div>
      </DialogContent>
    </Dialog>
  )
}

function AttachmentCard({
  attachment,
  onOpen
}: {
  attachment: AttachmentDetail
  onOpen: (attachment: AttachmentDetail) => void
}) {
  const name =
    attachment.downloadFilename ||
    attachment.longFilename ||
    attachment.filename ||
    `Attachment ${attachment.index + 1}`
  const meta = [
    attachment.mimeTag ? attachment.mimeTag : '',
    Number.isFinite(attachment.size || NaN) && (attachment.size || 0) > 0 ? formatBytes(attachment.size || 0) : '',
    attachment.isEmbeddedMessage ? 'Email' : '',
    attachment.parseError ? 'Error' : ''
  ]
    .filter(Boolean)
    .join(' · ')

  const downloadable = attachment.isDownloadable !== false && Boolean(attachment.downloadUrl)

  return (
    <div className="flex items-center justify-between gap-4 rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-3">
      <div className="min-w-0">
        <div className="truncate text-sm font-medium text-[color:var(--text)]">{name}</div>
        <div className="mt-0.5 text-xs text-[color:var(--muted)]">{meta}</div>
        {attachment.parseError ? <div className="mt-1 text-xs text-[color:var(--danger)]">{attachment.parseError}</div> : null}
      </div>
      {downloadable ? (
        <a
          className="button button-ghost whitespace-nowrap"
          href={attachment.downloadUrl}
          download={name}
          onClick={() => onOpen(attachment)}
        >
          Download
        </a>
      ) : (
        <span className="rounded-full border border-[color:var(--line)] px-3 py-1.5 text-xs text-[color:var(--muted)]">
          Unavailable
        </span>
      )}
    </div>
  )
}

export function EmptyState({ title, description }: { title: string; description: string }) {
  return (
    <div className="empty-state">
      <div className="text-base font-semibold text-[color:var(--text)]">{title}</div>
      <div>{description}</div>
    </div>
  )
}

