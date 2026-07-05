import * as React from 'react'
import { useVirtualizer } from '@tanstack/react-virtual'
import { ArrowDown, ArrowUp, ArrowUpDown, Cloud, FileText, Flag, Mail, RefreshCw, Search, Tag as TagIcon, X } from 'lucide-react'
import { api } from '@/api'
import { cn, formatDate, normalizeText } from '@/lib/utils'
import { normalizeSearchResultsPage, resolveSelectionScope } from '@/lib/search'
import type { MessageSummary, ReviewState, SearchSourceType } from '@/types'
import {
  Badge,
  Button,
  Dialog,
  DialogContent,
  DialogDescription,
  DialogTitle,
  IconButton,
  Input,
  ScrollArea
} from '@/components/ui'

type AllItemsSourceTab = SearchSourceType

const PAGE_SIZE = 100
const SOURCE_TABS: Array<{
  key: AllItemsSourceTab
  label: string
  icon: React.ReactNode
}> = [
  {
    key: 'mailbox',
    label: 'Mailbox',
    icon: <Mail className="h-4 w-4" aria-hidden="true" />
  },
  {
    key: 'teams',
    label: 'Teams',
    icon: <FileText className="h-4 w-4" aria-hidden="true" />
  },
  {
    key: 'sharepoint',
    label: 'SharePoint / OneDrive',
    icon: <Cloud className="h-4 w-4" aria-hidden="true" />
  }
]

interface TabState {
  items: MessageSummary[]
  page: number
  total: number
  totalPages: number
  loading: boolean
  loadingMore: boolean
  error: string
}

const INITIAL_TAB_STATE: TabState = {
  items: [],
  page: 0,
  total: 0,
  totalPages: 0,
  loading: false,
  loadingMore: false,
  error: ''
}

function createInitialTabStates(): Record<AllItemsSourceTab, TabState> {
  return {
    mailbox: { ...INITIAL_TAB_STATE, items: [] },
    teams: { ...INITIAL_TAB_STATE, items: [] },
    sharepoint: { ...INITIAL_TAB_STATE, items: [] }
  }
}

function getSourceLabel(sourceType?: SearchSourceType): string {
  if (sourceType === 'teams') {
    return 'Teams'
  }
  if (sourceType === 'sharepoint') {
    return 'SharePoint / OneDrive'
  }
  return 'Mailbox'
}

function getRowLocation(item: MessageSummary): string {
  if (item.sourceType === 'mailbox') {
    return [item.scopeLabel, item.folderPath || item.mailboxName || item.fileName || item.scopePath]
      .filter(Boolean)
      .join(' · ')
  }

  return [item.scopeLabel, item.archiveEntryPath || item.archivePath || item.fileName || item.scopePath]
    .filter(Boolean)
    .join(' · ')
}

function appendUniqueItems(existing: MessageSummary[], incoming: MessageSummary[]): MessageSummary[] {
  const seen = new Set(existing.map((item) => item.id))
  const next = [...existing]
  for (const item of incoming) {
    if (seen.has(item.id)) {
      continue
    }
    next.push(item)
    seen.add(item.id)
  }
  return next
}

type AllItemsSortColumn = 'subject' | 'sender' | 'location' | 'date'
type AllItemsSortDirection = 'asc' | 'desc'
type AllItemsSelectionState = Record<AllItemsSourceTab, Record<string, true>>

interface AllItemsSortState {
  column: AllItemsSortColumn
  direction: AllItemsSortDirection
}

const INITIAL_SORT_STATE: AllItemsSortState = {
  column: 'date',
  direction: 'desc'
}

const ITEM_GRID_TEMPLATE =
  'grid grid-cols-[48px_110px_minmax(0,1.6fr)_minmax(0,1.2fr)_minmax(0,1.4fr)_160px_110px]'

function createPageTrackingState(): Record<AllItemsSourceTab, Set<number>> {
  return {
    mailbox: new Set<number>(),
    teams: new Set<number>(),
    sharepoint: new Set<number>()
  }
}

function createSelectionState(): AllItemsSelectionState {
  return {
    mailbox: {},
    teams: {},
    sharepoint: {}
  }
}

function getSortKey(state: AllItemsSortState): string {
  return `${state.column}-${state.direction}`
}

function getSortLabel(state: AllItemsSortState): string {
  const label =
    state.column === 'subject'
      ? 'Subject'
      : state.column === 'sender'
        ? 'Sender'
        : state.column === 'location'
          ? 'Location'
          : 'Date'
  return `${label} ${state.direction === 'asc' ? 'ascending' : 'descending'}`
}

function getNextSortState(current: AllItemsSortState, column: AllItemsSortColumn): AllItemsSortState {
  if (current.column === column) {
    return {
      column,
      direction: current.direction === 'asc' ? 'desc' : 'asc'
    }
  }

  return {
    column,
    direction: column === 'date' ? 'desc' : 'asc'
  }
}

export function AllItemsDialog({
  open,
  selectedItemId,
  selectedCasePath,
  selectedScopePath,
  sourceCounts,
  onOpenChange,
  onSelectItem,
  onReviewChange,
  onRefreshCounts
}: {
  open: boolean
  selectedItemId: string
  selectedCasePath: string
  selectedScopePath: string
  sourceCounts: Record<SearchSourceType, number> | null
  onOpenChange: (open: boolean) => void
  onSelectItem: (item: MessageSummary) => void
  onReviewChange: (itemId: string, review: ReviewState) => void
  onRefreshCounts: () => void
}) {
  const parentRef = React.useRef<HTMLDivElement | null>(null)
  const selectAllRef = React.useRef<HTMLInputElement | null>(null)
  const scrollPositionsRef = React.useRef<Record<AllItemsSourceTab, number>>({
    mailbox: 0,
    teams: 0,
    sharepoint: 0
  })
  const loadedPagesRef = React.useRef<Record<AllItemsSourceTab, Set<number>>>(createPageTrackingState())
  const inFlightPagesRef = React.useRef<Record<AllItemsSourceTab, Set<number>>>(createPageTrackingState())
  const reviewRequestRef = React.useRef(0)
  const pagingGenerationRef = React.useRef(0)
  const activeTabRef = React.useRef<AllItemsSourceTab>('mailbox')

  const [activeTab, setActiveTab] = React.useState<AllItemsSourceTab>('mailbox')
  const [tabState, setTabState] = React.useState<Record<AllItemsSourceTab, TabState>>(createInitialTabStates)
  const [dialogError, setDialogError] = React.useState('')
  const [bulkActionError, setBulkActionError] = React.useState('')
  const [bulkActionLoading, setBulkActionLoading] = React.useState(false)
  const [pendingReviewIds, setPendingReviewIds] = React.useState<Record<string, boolean>>({})
  const [draftQuery, setDraftQuery] = React.useState('')
  const [appliedQuery, setAppliedQuery] = React.useState('')
  const [mailOnly, setMailOnly] = React.useState(false)
  const [reviewFlaggedOnly, setReviewFlaggedOnly] = React.useState(false)
  const [reviewTaggedOnly, setReviewTaggedOnly] = React.useState(false)
  const [sortState, setSortState] = React.useState<AllItemsSortState>(INITIAL_SORT_STATE)
  const [scrollResetToken, setScrollResetToken] = React.useState(0)
  const [selectedItemIdsByTab, setSelectedItemIdsByTab] = React.useState<AllItemsSelectionState>(
    createSelectionState
  )

  const selectedTotalsScope = React.useMemo(
    () => resolveSelectionScope(selectedCasePath, selectedScopePath),
    [selectedCasePath, selectedScopePath]
  )
  const activeSelection = selectedItemIdsByTab[activeTab]
  const activeState = tabState[activeTab]
  const sortKey = getSortKey(sortState)
  const sortLabel = getSortLabel(sortState)
  const hasMoreItems = activeState.page > 0 && activeState.page < activeState.totalPages
  const rows = activeState.items
  const selectedRows = React.useMemo(
    () => rows.filter((item) => Boolean(activeSelection[normalizeText(item.id)])),
    [activeSelection, rows]
  )
  const selectedCount = selectedRows.length
  const allLoadedSelected = rows.length > 0 && selectedCount === rows.length
  const someLoadedSelected = selectedCount > 0 && !allLoadedSelected
  const selectedHasPending = selectedRows.some((item) => Boolean(pendingReviewIds[normalizeText(item.id)]))
  const virtualizer = useVirtualizer({
    count: rows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 84,
    overscan: 8
  })

  function markReviewPending(itemId: string, pending: boolean): void {
    const key = normalizeText(itemId)
    if (!key) {
      return
    }
    setPendingReviewIds((current) => {
      const next = { ...current }
      if (pending) {
        next[key] = true
      } else {
        delete next[key]
      }
      return next
    })
  }

  function patchReviewState(itemId: string, review: ReviewState): void {
    const key = normalizeText(itemId)
    if (!key) {
      return
    }
    setTabState((current) => {
      const next = { ...current }
      for (const tab of SOURCE_TABS) {
        const tabKey = tab.key
        const items = next[tabKey].items
        if (!items.length) {
          continue
        }
        let changed = false
        const patched = items.map((item) => {
          if (normalizeText(item.id) !== key) {
            return item
          }
          changed = true
          return { ...item, review }
        })
        if (changed) {
          next[tabKey] = {
            ...next[tabKey],
            items: patched
          }
        }
      }
      return next
    })
    onReviewChange(key, review)
  }

  const clearSelection = React.useCallback((): void => {
    setSelectedItemIdsByTab(createSelectionState())
    setBulkActionError('')
  }, [])

  function updateSelectionForTab(
    tab: AllItemsSourceTab,
    updater: (selection: Record<string, true>) => Record<string, true>
  ): void {
    setSelectedItemIdsByTab((current) => ({
      ...current,
      [tab]: updater(current[tab])
    }))
  }

  function setItemSelected(tab: AllItemsSourceTab, itemId: string, selected: boolean): void {
    const key = normalizeText(itemId)
    if (!key) {
      return
    }
    updateSelectionForTab(tab, (current) => {
      const next = { ...current }
      if (selected) {
        next[key] = true
      } else {
        delete next[key]
      }
      return next
    })
  }

  function setAllItemsSelected(tab: AllItemsSourceTab, items: MessageSummary[], selected: boolean): void {
    updateSelectionForTab(tab, () => {
      if (!selected) {
        return {}
      }
      const next: Record<string, true> = {}
      for (const item of items) {
        const key = normalizeText(item.id)
        if (key) {
          next[key] = true
        }
      }
      return next
    })
  }

  function handleTabChange(tab: AllItemsSourceTab): void {
    if (tab === activeTab) {
      return
    }
    clearSelection()
    setActiveTab(tab)
  }

  const resetPagingState = React.useCallback(() => {
    pagingGenerationRef.current += 1
    loadedPagesRef.current = createPageTrackingState()
    inFlightPagesRef.current = createPageTrackingState()
    scrollPositionsRef.current = {
      mailbox: 0,
      teams: 0,
      sharepoint: 0
    }
    setTabState(createInitialTabStates())
    setDialogError('')
    clearSelection()
    setScrollResetToken((current) => current + 1)
  }, [clearSelection])

  async function loadPage(tab: AllItemsSourceTab, page: number): Promise<void> {
    const normalizedTab = tab
    const normalizedPage = Math.max(1, Math.floor(page || 1))
    const loadedPages = loadedPagesRef.current[normalizedTab]
    const inFlightPages = inFlightPagesRef.current[normalizedTab]
    if (loadedPages.has(normalizedPage) || inFlightPages.has(normalizedPage)) {
      return
    }

    const requestGeneration = pagingGenerationRef.current
    inFlightPages.add(normalizedPage)
    setDialogError('')
    setTabState((current) => {
      const prev = current[normalizedTab]
      return {
        ...current,
        [normalizedTab]: {
          ...prev,
          loading: normalizedPage === 1 && prev.items.length === 0,
          loadingMore: normalizedPage > 1 || (normalizedPage === 1 && prev.items.length > 0),
          error: ''
        }
      }
    })

    try {
      const response = await api.search({
        scope: selectedTotalsScope.scope,
        sourceType: normalizedTab,
        casePath: selectedTotalsScope.casePath || undefined,
        scopePath: selectedTotalsScope.scopePath || undefined,
        query: appliedQuery,
        mode: 'and',
        page: normalizedPage,
        pageSize: PAGE_SIZE,
        mailOnly,
        sort: sortKey,
        reviewFlagged: reviewFlaggedOnly,
        reviewTagged: reviewTaggedOnly
      })
      if (pagingGenerationRef.current !== requestGeneration) {
        return
      }
      const pageResponse = normalizeSearchResultsPage(response.page)
      loadedPages.add(pageResponse.page)
      setTabState((current) => {
        const prev = current[normalizedTab]
        const mergedItems =
          pageResponse.page === 1 ? pageResponse.items : appendUniqueItems(prev.items, pageResponse.items)
        return {
          ...current,
          [normalizedTab]: {
            ...prev,
            items: mergedItems,
            page: pageResponse.page,
            total: pageResponse.total,
            totalPages: pageResponse.totalPages,
            loading: false,
            loadingMore: false,
            error: ''
          }
        }
      })
    } catch (error) {
      if (pagingGenerationRef.current !== requestGeneration) {
        return
      }
      const message = error instanceof Error ? error.message : 'Unable to load items'
      setTabState((current) => {
        const prev = current[normalizedTab]
        return {
          ...current,
          [normalizedTab]: {
            ...prev,
            loading: false,
            loadingMore: false,
            error: message
          }
        }
      })
      setDialogError(message)
    } finally {
      inFlightPages.delete(normalizedPage)
    }
  }

  function handleQuerySubmit(event: React.FormEvent<HTMLFormElement>): void {
    event.preventDefault()
    const nextQuery = normalizeText(draftQuery)
    setDraftQuery(nextQuery)
    setAppliedQuery(nextQuery)
  }

  function handleClearSearch(): void {
    setDraftQuery('')
    setAppliedQuery('')
  }

  function handleSortColumn(column: AllItemsSortColumn): void {
    setSortState((current) => getNextSortState(current, column))
  }

  function handleScroll(event: React.UIEvent<HTMLDivElement>): void {
    scrollPositionsRef.current[activeTab] = event.currentTarget.scrollTop
  }

  async function handleLoadMore(): Promise<void> {
    if (!hasMoreItems || activeState.loadingMore) {
      return
    }
    await loadPage(activeTab, activeState.page + 1)
  }

  async function handleReviewToggle(item: MessageSummary): Promise<void> {
    const key = normalizeText(item.id)
    if (!key || pendingReviewIds[key]) {
      return
    }

    const currentReview = item.review || {
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    }
    const nextReview: ReviewState = {
      ...currentReview,
      flagged: !currentReview.flagged
    }
    const requestId = ++reviewRequestRef.current

    markReviewPending(key, true)
    patchReviewState(key, nextReview)

    try {
      await api.item.updateReview(key, {
        flagged: nextReview.flagged,
        tags: [...(currentReview.tags || [])]
      })
    } catch (error) {
      if (reviewRequestRef.current === requestId) {
        patchReviewState(key, currentReview)
        setDialogError(error instanceof Error ? error.message : 'Unable to update review')
      }
    } finally {
      if (reviewRequestRef.current === requestId) {
        markReviewPending(key, false)
      }
    }
  }

  async function handleBulkFlagSelected(): Promise<void> {
    if (bulkActionLoading || !selectedRows.length || selectedHasPending) {
      return
    }

    const operationTab = activeTabRef.current
    const operationGeneration = pagingGenerationRef.current
    const itemsToFlag = [...selectedRows]
    setBulkActionLoading(true)
    setBulkActionError('')

    try {
      const results = await Promise.allSettled(
        itemsToFlag.map(async (item) => {
          const key = normalizeText(item.id)
          const currentReview = item.review || {
            flagged: false,
            tags: [],
            createdAt: '',
            updatedAt: ''
          }
          const response = await api.item.updateReview(key, {
            flagged: true,
            tags: [...(currentReview.tags || [])]
          })
          return {
            itemId: key,
            review: response.review
          }
        })
      )

      if (pagingGenerationRef.current !== operationGeneration || activeTabRef.current !== operationTab) {
        return
      }

      const successfulIds: string[] = []
      const failedCount = results.reduce((count, result, index) => {
        if (result.status === 'fulfilled') {
          successfulIds.push(result.value.itemId)
          patchReviewState(result.value.itemId, result.value.review)
          return count
        }

        return count + 1
      }, 0)

      if (!failedCount) {
        clearSelection()
        return
      }

      updateSelectionForTab(operationTab, (current) => {
        const next = { ...current }
        for (const itemId of successfulIds) {
          delete next[itemId]
        }
        return next
      })

      setBulkActionError(
        failedCount === 1
          ? '1 selected item could not be flagged.'
          : `${failedCount} selected items could not be flagged.`
      )
    } catch (error) {
      setBulkActionError(error instanceof Error ? error.message : 'Unable to flag selected items')
    } finally {
      setBulkActionLoading(false)
    }
  }

  React.useEffect(() => {
    if (!open) {
      return
    }
    resetPagingState()
  }, [
    appliedQuery,
    mailOnly,
    open,
    resetPagingState,
    reviewFlaggedOnly,
    reviewTaggedOnly,
    selectedCasePath,
    selectedScopePath,
    sortKey
  ])

  React.useEffect(() => {
    activeTabRef.current = activeTab
  }, [activeTab])

  React.useEffect(() => {
    if (!open) {
      return
    }

    const state = tabState[activeTab]
    if (!state.items.length && !state.loading && !state.loadingMore) {
      void loadPage(activeTab, 1)
    }
  }, [activeTab, activeState.loading, activeState.loadingMore, activeState.items.length, open, sortKey, tabState])

  React.useEffect(() => {
    if (open) {
      return
    }
    clearSelection()
    setBulkActionLoading(false)
    setDialogError('')
  }, [open, clearSelection])

  React.useLayoutEffect(() => {
    if (!open) {
      return
    }
    const el = parentRef.current
    if (!el) {
      return
    }
    el.scrollTop = scrollPositionsRef.current[activeTab] || 0
  }, [activeTab, open])

  React.useLayoutEffect(() => {
    const el = parentRef.current
    if (!el) {
      return
    }
    el.scrollTop = 0
  }, [scrollResetToken])

  React.useLayoutEffect(() => {
    if (!selectAllRef.current) {
      return
    }
    selectAllRef.current.indeterminate = someLoadedSelected
  }, [someLoadedSelected, activeTab, selectedCount, rows.length])

  const virtualItems = virtualizer.getVirtualItems()

  function renderSortHeader(column: AllItemsSortColumn, label: string): React.ReactNode {
    const active = sortState.column === column
    const direction = active ? sortState.direction : null
    const directionLabel = direction === 'asc' ? 'ascending' : direction === 'desc' ? 'descending' : ''
    return (
      <button
        type="button"
        className={cn(
          'flex w-full items-center justify-between gap-2 rounded-xl px-2 py-1.5 text-left transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]',
          active ? 'text-[color:var(--accent)]' : 'text-[color:var(--muted)] hover:text-[color:var(--text)]'
        )}
        aria-label={`Sort by ${label}${directionLabel ? ` ${directionLabel}` : ''}`}
        aria-pressed={active}
        onClick={() => handleSortColumn(column)}
      >
        <span>{label}</span>
        {active ? (
          direction === 'asc' ? (
            <ArrowUp className="h-3.5 w-3.5" aria-hidden="true" />
          ) : (
            <ArrowDown className="h-3.5 w-3.5" aria-hidden="true" />
          )
        ) : (
          <ArrowUpDown className="h-3.5 w-3.5 opacity-60" aria-hidden="true" />
        )}
      </button>
    )
  }

  return (
    <Dialog
      open={open}
      onOpenChange={(nextOpen) => {
        onOpenChange(nextOpen)
      }}
    >
      <DialogContent
        showCloseButton={false}
        className="h-[min(96vh,980px)] max-h-[96vh] w-[min(98vw,1680px)] overflow-hidden p-0"
      >
        <div className="flex h-full min-h-0 flex-col">
          <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
            <div className="min-w-0">
              <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                Workspace
              </div>
              <DialogTitle className="mt-1 text-2xl">All items review</DialogTitle>
              <DialogDescription>
                Browse the accessible index, switch sources, and flag items in place.
              </DialogDescription>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                <span className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Case scope
                </span>
                <Badge className="border-[color:var(--line)] bg-[color:var(--surface)] text-[color:var(--text)]">
                  {selectedTotalsScope.scopePath || selectedTotalsScope.casePath || 'All cases'}
                </Badge>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <Button variant="ghost" onClick={onRefreshCounts}>
                <RefreshCw className="h-4 w-4" />
                Refresh counts
              </Button>
              <Button variant="ghost" onClick={() => onOpenChange(false)}>
                <X className="h-4 w-4" />
                Close
              </Button>
            </div>
          </div>

          <div className="space-y-4 border-b border-[color:var(--line)] px-6 py-4">
            <div className="flex flex-wrap items-center gap-2">
              {SOURCE_TABS.map((tab) => {
                const active = tab.key === activeTab
                const count = sourceCounts ? sourceCounts[tab.key] : '…'
                return (
                  <button
                    key={tab.key}
                    type="button"
                    className={cn(
                      'inline-flex items-center gap-2 rounded-2xl border px-4 py-2 text-sm font-semibold transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]',
                      active
                        ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)] text-[color:var(--accent)]'
                        : 'border-[color:var(--line)] bg-[color:var(--surface-strong)] text-[color:var(--text)] hover:border-[color:var(--accent-soft)] hover:bg-[color:var(--surface-soft)]'
                    )}
                    onClick={() => handleTabChange(tab.key)}
                  >
                    {tab.icon}
                    <span>{tab.label}</span>
                    <Badge className="border-[color:var(--line)] bg-[color:var(--surface)] text-[color:var(--muted)]">
                      {count}
                    </Badge>
                  </button>
                )
              })}
            </div>

            <div className="grid gap-3 xl:grid-cols-[minmax(0,1fr)_auto]">
              <form className="flex min-w-0 items-center gap-2" onSubmit={handleQuerySubmit}>
                <Input
                  className="min-w-0"
                  value={draftQuery}
                  onChange={(event) => setDraftQuery(event.target.value)}
                  placeholder="Search within this case"
                />
                <Button type="submit">
                  <Search className="h-4 w-4" />
                  Search
                </Button>
                {draftQuery || appliedQuery ? (
                  <Button type="button" variant="ghost" onClick={handleClearSearch}>
                    Clear search
                  </Button>
                ) : null}
              </form>

              <div className="flex flex-wrap items-center gap-2">
                <Button
                  type="button"
                  variant={mailOnly ? 'secondary' : 'ghost'}
                  aria-pressed={mailOnly}
                  onClick={() => setMailOnly((current) => !current)}
                >
                  <Mail className="h-4 w-4" />
                  Mail only
                </Button>
                <Button
                  type="button"
                  variant={reviewFlaggedOnly ? 'secondary' : 'ghost'}
                  aria-pressed={reviewFlaggedOnly}
                  onClick={() => setReviewFlaggedOnly((current) => !current)}
                >
                  <Flag className="h-4 w-4" />
                  Flagged
                </Button>
                <Button
                  type="button"
                  variant={reviewTaggedOnly ? 'secondary' : 'ghost'}
                  aria-pressed={reviewTaggedOnly}
                  onClick={() => setReviewTaggedOnly((current) => !current)}
                >
                  <TagIcon className="h-4 w-4" />
                  Tagged
                </Button>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <span className="chip">Search: {appliedQuery || 'All items'}</span>
              <span className="chip">Sorted by {sortLabel}</span>
              {mailOnly ? <span className="chip chip-active">Mail only</span> : null}
              {reviewFlaggedOnly ? <span className="chip chip-active">Flagged</span> : null}
              {reviewTaggedOnly ? <span className="chip chip-active">Tagged</span> : null}
            </div>
          </div>

          {selectedCount > 0 ? (
            <div className="border-b border-[color:var(--line)] px-6 py-4">
              <div className="rounded-2xl border border-[color:rgba(59,130,246,0.22)] bg-[color:rgba(59,130,246,0.08)] px-4 py-3">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex flex-wrap items-center gap-3">
                    <Badge className="border-[color:rgba(59,130,246,0.22)] bg-[color:var(--surface-strong)] text-[color:var(--accent)]">
                      {selectedCount}
                    </Badge>
                    <div className="text-sm">
                      <div className="font-semibold text-[color:var(--text)]">
                        {selectedCount === 1 ? 'item selected' : 'items selected'}
                      </div>
                      <div className="text-[color:var(--muted)]">
                        Selection stays with loaded rows in this tab.
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <Button
                      type="button"
                      variant="secondary"
                      onClick={() => void handleBulkFlagSelected()}
                      disabled={bulkActionLoading || selectedHasPending}
                    >
                      {bulkActionLoading ? <RefreshCw className="h-4 w-4 animate-spin" /> : <Flag className="h-4 w-4" />}
                      {bulkActionLoading ? 'Flagging...' : 'Flag selected'}
                    </Button>
                    <Button type="button" variant="ghost" onClick={clearSelection} disabled={bulkActionLoading}>
                      Clear selection
                    </Button>
                  </div>
                </div>

                {bulkActionError ? (
                  <div className="mt-3 text-sm text-[color:var(--danger)]">{bulkActionError}</div>
                ) : null}
              </div>
            </div>
          ) : null}

          <div
            className={`${ITEM_GRID_TEMPLATE} border-b border-[color:var(--line)] bg-[color:var(--surface)] px-4 py-3 text-[11px] font-semibold uppercase tracking-[0.2em] text-[color:var(--muted)]`}
          >
            <div className="flex items-center justify-center">
              <input
                ref={selectAllRef}
                type="checkbox"
                className="h-4 w-4 rounded border-[color:var(--line)] accent-[color:var(--accent)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]"
                aria-label="Select all loaded items"
                checked={allLoadedSelected}
                disabled={!rows.length || bulkActionLoading}
                onChange={(event) => {
                  setAllItemsSelected(activeTab, rows, event.target.checked)
                }}
              />
            </div>
            <div>Source</div>
            {renderSortHeader('subject', 'Subject')}
            {renderSortHeader('sender', 'Sender')}
            {renderSortHeader('location', 'Location')}
            {renderSortHeader('date', 'Date')}
            <div className="text-right">Flagged</div>
          </div>

          <ScrollArea ref={parentRef} className="min-h-0 flex-1" data-testid="all-items-scroll" onScroll={handleScroll}>
            {dialogError ? (
              <div className="border-b border-[color:var(--line)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {dialogError}
              </div>
            ) : null}
            {activeState.error ? (
              <div className="border-b border-[color:var(--line)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {activeState.error}
              </div>
            ) : null}

            {!rows.length && activeState.loading ? (
              <div className="empty-state m-4">Loading items...</div>
            ) : !rows.length ? (
              <div className="empty-state m-4">No items found in this tab.</div>
            ) : (
              <div
                className="relative w-full"
                style={{
                  height: `${virtualizer.getTotalSize()}px`
                }}
              >
                {virtualItems.map((virtualRow) => {
                  const item = rows[virtualRow.index]
                  if (!item) {
                    return null
                  }
                  const review = item.review || {
                    flagged: false,
                    tags: [],
                    createdAt: '',
                    updatedAt: ''
                  }
                  const flagged = Boolean(review.flagged)
                  const selected = Boolean(activeSelection[normalizeText(item.id)])
                  const active = normalizeText(item.id) === normalizeText(selectedItemId)
                  const pending = Boolean(pendingReviewIds[normalizeText(item.id)])
                  return (
                    <div
                      key={item.id}
                      ref={virtualizer.measureElement}
                      data-index={virtualRow.index}
                      className="absolute left-0 top-0 w-full px-3 py-1.5"
                      style={{
                        transform: `translateY(${virtualRow.start}px)`
                      }}
                    >
                      <div
                        role="button"
                        tabIndex={0}
                        aria-selected={active}
                        className={cn(
                          `${ITEM_GRID_TEMPLATE} items-center gap-4 rounded-2xl border px-4 py-3 text-left transition focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]`,
                          selected
                            ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)]'
                            : active
                            ? 'border-[color:var(--accent)] bg-[color:var(--accent-soft)]'
                            : 'border-[color:var(--line)] bg-[color:var(--surface-strong)] hover:border-[color:var(--accent-soft)] hover:bg-[color:var(--surface-soft)]',
                          pending && 'opacity-80'
                        )}
                        onClick={() => {
                          onSelectItem(item)
                        }}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault()
                            onSelectItem(item)
                          }
                        }}
                      >
                        <div className="flex items-center justify-center">
                          <input
                            type="checkbox"
                            className="h-4 w-4 rounded border-[color:var(--line)] accent-[color:var(--accent)] focus:outline-none focus:ring-2 focus:ring-[color:var(--focus-ring)]"
                            aria-label={`Select item ${item.subject || item.messageId || item.id}`}
                            checked={selected}
                            disabled={bulkActionLoading}
                            onClick={(event) => {
                              event.stopPropagation()
                            }}
                            onKeyDown={(event) => {
                              event.stopPropagation()
                            }}
                            onChange={(event) => {
                              setItemSelected(activeTab, item.id, event.target.checked)
                            }}
                          />
                        </div>

                        <div className="min-w-0">
                          <span
                            className={cn(
                              'inline-flex items-center gap-2 rounded-full border px-2.5 py-1 text-[11px] font-semibold',
                              item.sourceType === 'mailbox'
                                ? 'border-[color:rgba(59,130,246,0.2)] bg-[color:rgba(59,130,246,0.08)] text-[color:var(--accent)]'
                                : item.sourceType === 'teams'
                                  ? 'border-[color:rgba(16,185,129,0.2)] bg-[color:rgba(16,185,129,0.08)] text-[color:#0f766e]'
                                  : 'border-[color:rgba(245,158,11,0.2)] bg-[color:rgba(245,158,11,0.08)] text-[color:var(--warning)]'
                            )}
                          >
                            {item.sourceType === 'mailbox' ? (
                              <Mail className="h-3.5 w-3.5" aria-hidden="true" />
                            ) : item.sourceType === 'teams' ? (
                              <FileText className="h-3.5 w-3.5" aria-hidden="true" />
                            ) : (
                              <Cloud className="h-3.5 w-3.5" aria-hidden="true" />
                            )}
                            <span>{getSourceLabel(item.sourceType)}</span>
                          </span>
                        </div>

                        <div className="min-w-0">
                          <div
                            className="truncate text-sm font-semibold text-[color:var(--text)]"
                            title={item.subject || '(no subject)'}
                          >
                            {item.subject || '(no subject)'}
                          </div>
                          <div className="mt-1 truncate text-xs text-[color:var(--muted)]">
                            {item.kind || item.messageClass || item.previewKind || ''}
                          </div>
                        </div>

                        <div className="min-w-0">
                          <div
                            className="truncate text-sm text-[color:var(--text)]"
                            title={item.senderName || item.senderEmailAddress || '(unknown sender)'}
                          >
                            {item.senderName || item.senderEmailAddress || '(unknown sender)'}
                          </div>
                          {item.senderEmailAddress ? (
                            <div className="mt-1 truncate text-xs text-[color:var(--muted)]" title={item.senderEmailAddress}>
                              {item.senderEmailAddress}
                            </div>
                          ) : null}
                        </div>

                        <div className="min-w-0">
                          <div className="truncate text-sm text-[color:var(--text)]" title={getRowLocation(item)}>
                            {getRowLocation(item) || item.scopeLabel || item.scopePath || '(unknown location)'}
                          </div>
                          <div className="mt-1 truncate text-xs text-[color:var(--muted)]" title={item.scopePath || ''}>
                            {item.scopePath || ''}
                          </div>
                        </div>

                        <div
                          className="text-sm text-[color:var(--text)]"
                          title={item.sortDate || item.creationTime || item.clientSubmitTime || ''}
                        >
                          {formatDate(item.sortDate || item.creationTime || item.clientSubmitTime)}
                        </div>

                        <div className="flex justify-end">
                          <IconButton
                            label={flagged ? 'Unflag item' : 'Flag item'}
                            className={cn(
                              'h-9 w-9',
                              flagged ? 'icon-button-primary' : '',
                              (pending || bulkActionLoading) && 'opacity-60'
                            )}
                            onClick={(event) => {
                              event.stopPropagation()
                              void handleReviewToggle(item)
                            }}
                            disabled={pending || bulkActionLoading}
                          >
                            <Flag className={cn('h-4 w-4', flagged && 'fill-current')} />
                          </IconButton>
                        </div>
                      </div>
                    </div>
                  )
                })}
              </div>
            )}

            {hasMoreItems ? (
              <div className="flex justify-center px-4 py-6">
                <Button variant="secondary" onClick={() => void handleLoadMore()} disabled={activeState.loadingMore}>
                  {activeState.loadingMore ? (
                    <>
                      <RefreshCw className="h-4 w-4 animate-spin" />
                      Loading more items...
                    </>
                  ) : (
                    'Load more items'
                  )}
                </Button>
              </div>
            ) : null}
          </ScrollArea>
        </div>
      </DialogContent>
    </Dialog>
  )
}
