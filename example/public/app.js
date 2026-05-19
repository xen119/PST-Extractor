(function () {
  const STORAGE_KEYS = {
    casePath: 'pst-mail-explorer.casePath',
    scopePath: 'pst-mail-explorer.scopePath',
    mailboxScopeView: 'pst-mail-explorer.mailboxScopeView',
    pstFileName: 'pst-mail-explorer.pstFileName',
    folderId: 'pst-mail-explorer.folderId',
    messageId: 'pst-mail-explorer.messageId',
    query: 'pst-mail-explorer.query',
    searchScope: 'pst-mail-explorer.searchScope',
    mailOnly: 'pst-mail-explorer.mailOnly',
    sort: 'pst-mail-explorer.sort',
    reviewFlaggedOnly: 'pst-mail-explorer.reviewFlaggedOnly',
    reviewTaggedOnly: 'pst-mail-explorer.reviewTaggedOnly',
    flaggedBundleScope: 'pst-mail-explorer.flaggedBundleScope'
  }

  const state = {
    sessionId: null,
    catalogScopes: [],
    catalog: [],
    catalogLoaded: false,
    catalogMessage: '',
    selectedCasePath: null,
    selectedScopePath: null,
    selectedScopeLabel: '',
    mailboxScopeView: 'search',
    selectedPstFileName: null,
    mailboxFilter: '',
    summary: null,
    tree: null,
    folderMap: new Map(),
    currentFolderId: null,
    currentFolderPage: null,
    currentSearchPage: null,
    currentMessageDetail: null,
    selectedMessageId: null,
    query: '',
    searchScope: 'pst',
    activeSearch: null,
    hiddenRules: [],
    hiddenRulesLoaded: false,
    hiddenFiltersOpen: false,
    mailOnly: true,
    sort: 'date-desc',
    reviewFlaggedOnly: false,
    reviewTaggedOnly: false,
    bundleScope: 'all',
    pageSize: 50
  }

  const ui = {}

  function getElement(id) {
    const element = document.getElementById(id)
    if (!element) {
      throw new Error(`Missing expected element: ${id}`)
    }
    return element
  }

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, (char) => {
      switch (char) {
        case '&':
          return '&amp;'
        case '<':
          return '&lt;'
        case '>':
          return '&gt;'
        case '"':
          return '&quot;'
        case "'":
          return '&#39;'
        default:
          return char
      }
    })
  }

  function escapeAttr(value) {
    return escapeHtml(value).replace(/`/g, '&#96;')
  }

  function hasText(value) {
    return Boolean(String(value ?? '').trim())
  }

  function normalizeScopePath(value) {
    const text = String(value ?? '').trim().replace(/\\/g, '/')
    if (!text || text === '.') {
      return ''
    }

    const parts = text
      .split('/')
      .map((part) => part.trim())
      .filter(Boolean)

    if (!parts.length || parts.some((part) => part === '..')) {
      return ''
    }

    return parts.join('/')
  }

  function localNormalizeMessageId(item) {
    if (!item || typeof item !== 'object') {
      return ''
    }
    return String(item.id || item.messageId || '').trim()
  }

  function localNormalizeSearchResultItem(item) {
    if (!item || typeof item !== 'object') {
      return item
    }
    const id = localNormalizeMessageId(item)
    if (!id) {
      return { ...item }
    }
    return {
      ...item,
      id,
      messageId: String(item.messageId || item.id || id).trim()
    }
  }

  function localNormalizeSearchResultsPage(page) {
    if (!page || typeof page !== 'object') {
      return page
    }
    const items = Array.isArray(page.items) ? page.items : []
    return {
      ...page,
      items: items.map((item) => localNormalizeSearchResultItem(item))
    }
  }

  const SEARCH_RESULT_HELPERS =
    (typeof window !== 'undefined' && window.PstExplorerSearch) || {
      normalizeMessageId: localNormalizeMessageId,
      normalizeSearchResultItem: localNormalizeSearchResultItem,
      normalizeSearchResultsPage: localNormalizeSearchResultsPage
    }

  function resolveMessageId(item) {
    return SEARCH_RESULT_HELPERS.normalizeMessageId(item)
  }

  function normalizeSearchResultItem(item) {
    return SEARCH_RESULT_HELPERS.normalizeSearchResultItem(item)
  }

  function normalizeSearchResultsPage(page) {
    return SEARCH_RESULT_HELPERS.normalizeSearchResultsPage(page)
  }

  function getCasePathFromScopePath(scopePath) {
    const normalized = normalizeScopePath(scopePath)
    if (!normalized) {
      return ''
    }

    return normalized.split('/')[0] || ''
  }

  function getCaseLabel(casePath) {
    return casePath ? String(casePath).split('/').join(' / ') : 'PST root'
  }

  function getSearchLabel(scopePath) {
    const normalized = normalizeScopePath(scopePath)
    if (!normalized) {
      return 'PST root'
    }

    const parts = normalized.split('/')
    return parts[parts.length - 1] || 'PST root'
  }

  function getScopeLabel(scopePath) {
    return scopePath ? String(scopePath).split('/').join(' / ') : 'PST root'
  }

  const ALL_PSTS_SCOPE_VALUE = '__all_psts__'

  function isAllPstsScopeValue(value) {
    return String(value || '') === ALL_PSTS_SCOPE_VALUE
  }

  function normalizeMailboxFilter(value) {
    return String(value || '')
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase()
  }

  function deriveSearchMode(query) {
    const text = String(query || '').trim()
    if (!text) {
      return 'and'
    }

    const pattern = /"([^"]+)"|(\S+)/g
    let match = null
    while ((match = pattern.exec(text))) {
      const raw = String(match[1] || match[2] || '').trim()
      if (!raw) {
        continue
      }

      const token = raw.toLowerCase()
      if (token === '|' || token.startsWith('|')) {
        return 'or'
      }
      if (token === '+' || token.startsWith('+')) {
        return 'and'
      }
    }

    return 'and'
  }

  function normalizeHiddenRuleValue(value) {
    return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase()
  }

  function extractEmailAddress(value) {
    const match = String(value || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)
    return match ? match[0].trim().toLowerCase() : ''
  }

  function splitRecipientTokens(value) {
    return String(value || '')
      .split(';')
      .map((item) => String(item || '').trim())
      .filter(Boolean)
  }

  function renderInlineHideButton(kind, value, label) {
    const normalizedValue = normalizeHiddenRuleValue(value)
    if (!normalizedValue) {
      return ''
    }

    return `
      <button
        class="inline-filter-button"
        type="button"
        data-action="hide-${kind}"
        data-filter-value="${escapeAttr(normalizedValue)}"
        data-filter-label="${escapeAttr(label || value)}"
      >
        -
      </button>
    `
  }

  function normalizeHiddenRule(rule) {
    if (!rule || typeof rule !== 'object') {
      return null
    }

    return {
      filterId: String(rule.filterId || rule.id || ''),
      kind: String(rule.kind || ''),
      value: normalizeHiddenRuleValue(rule.value),
      label: String(rule.label || rule.value || ''),
      createdAt: String(rule.createdAt || ''),
      updatedAt: String(rule.updatedAt || '')
    }
  }

  function isSearchResultView() {
    return Boolean(state.currentSearchPage && state.activeSearch && hasSearchCriteria(state.activeSearch))
  }

  function hasSearchCriteria(criteria) {
    if (!criteria) {
      return false
    }
    return Boolean(String(criteria.query || '').trim())
  }

  function collectSearchDraft() {
    const query = String(ui.searchInput.value || '').trim()
    return {
      query,
      searchScope: ui.searchScopeSelect.value || 'pst',
      mode: deriveSearchMode(query)
    }
  }

  function syncSearchDraftToState() {
    const draft = collectSearchDraft()
    state.query = draft.query
    state.searchScope = draft.searchScope
    saveState()
    return draft
  }

  function getCatalogScope(scopePath) {
    const normalized = normalizeScopePath(scopePath)
    return (
      state.catalogScopes.find((scope) => normalizeScopePath(scope.scopePath) === normalized) ||
      null
    )
  }

  function getCatalogCases() {
    const cases = new Map()
    for (const scope of Array.isArray(state.catalogScopes) ? state.catalogScopes : []) {
      const casePath = getCasePathFromScopePath(scope.scopePath)
      const caseLabel = getCaseLabel(casePath)
      if (!cases.has(casePath)) {
        cases.set(casePath, {
          casePath,
          caseLabel,
          searches: []
        })
      }
      cases.get(casePath).searches.push(scope)
    }

    return [...cases.values()].sort((left, right) => {
      if (left.casePath === right.casePath) {
        return 0
      }
      if (left.casePath === '') {
        return -1
      }
      if (right.casePath === '') {
        return 1
      }
      return left.caseLabel.localeCompare(right.caseLabel, undefined, { sensitivity: 'base' })
    })
  }

  function getSearchesForCase(casePath) {
    const normalizedCasePath = normalizeScopePath(casePath)
    return (Array.isArray(state.catalogScopes) ? state.catalogScopes : [])
      .filter((scope) => getCasePathFromScopePath(scope.scopePath) === normalizedCasePath)
      .sort((left, right) =>
        getSearchLabel(left.scopePath).localeCompare(getSearchLabel(right.scopePath), undefined, {
          sensitivity: 'base'
        })
      )
  }

  function getMailboxEntriesForCase(casePath, viewMode = 'search') {
    const normalizedCasePath = normalizeScopePath(casePath)
    const searches = getSearchesForCase(normalizedCasePath)
    const selectedScopePath = normalizeScopePath(state.selectedScopePath || '')
    const entries = []

    const appendScopeFiles = (scope) => {
      for (const file of scope.files || []) {
        entries.push({
          ...file,
          casePath: normalizedCasePath,
          scopePath: scope.scopePath,
          scopeLabel: scope.scopeLabel,
          displayPath: scope.scopeLabel || scope.scopePath || ''
        })
      }
    }

    if (viewMode === 'all') {
      for (const scope of searches) {
        appendScopeFiles(scope)
      }
    } else {
      const selectedScope =
        searches.find((scope) => normalizeScopePath(scope.scopePath) === selectedScopePath) ||
        searches[0] ||
        null
      if (selectedScope) {
        appendScopeFiles(selectedScope)
      }
    }

    entries.sort((left, right) => {
      if (viewMode === 'all') {
        const scopeCompare = String(left.displayPath || '').localeCompare(
          String(right.displayPath || ''),
          undefined,
          { sensitivity: 'base' }
        )
        if (scopeCompare !== 0) {
          return scopeCompare
        }
      }
      return String(left.fileName || '').localeCompare(String(right.fileName || ''), undefined, {
        sensitivity: 'base'
      })
    })

    return entries
  }

  function mailboxMatchesFilter(entry, filterText) {
    if (!filterText) {
      return true
    }

    const haystack = normalizeMailboxFilter(
      [entry.fileName, entry.displayPath, entry.scopePath].filter(Boolean).join(' ')
    )
    return haystack.includes(filterText)
  }

  function getMailboxEntriesForDisplay(casePath, viewMode = 'search') {
    const entries = getMailboxEntriesForCase(casePath, viewMode)
    const filterText = normalizeMailboxFilter(state.mailboxFilter)
    return filterText ? entries.filter((entry) => mailboxMatchesFilter(entry, filterText)) : entries
  }

  function formatDate(value) {
    if (!value) {
      return 'Not available'
    }
    const date = new Date(value)
    if (Number.isNaN(date.getTime())) {
      return value
    }
    return new Intl.DateTimeFormat(undefined, {
      dateStyle: 'medium',
      timeStyle: 'short'
    }).format(date)
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes < 0) {
      return 'Unknown size'
    }
    if (bytes === 0) {
      return '0 B'
    }
    const units = ['B', 'KB', 'MB', 'GB']
    const index = Math.min(
      units.length - 1,
      Math.floor(Math.log(bytes) / Math.log(1024))
    )
    const value = bytes / 1024 ** index
    return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`
  }

  function formatImportance(value) {
    if (value === 2) {
      return 'High'
    }
    if (value === 1) {
      return 'Normal'
    }
    if (value === 0) {
      return 'Low'
    }
    return String(value ?? 'Unknown')
  }

  function formatChipList(values) {
    return values.filter(Boolean).join(' ')
  }

  function getSenderInitials(detail) {
    const source =
      detail.senderName ||
      detail.senderEmailAddress ||
      detail.displayTo ||
      detail.subject ||
      'M'
    const parts = String(source)
      .replace(/<[^>]+>/g, ' ')
      .trim()
      .split(/[\s.@]+/)
      .filter(Boolean)

    if (!parts.length) {
      return 'M'
    }
    if (parts.length === 1) {
      return parts[0].slice(0, 2).toUpperCase()
    }
    return `${parts[0][0]}${parts[1][0]}`.toUpperCase()
  }

  function normalizeReviewState(review) {
    if (!review || typeof review !== 'object') {
      return {
        flagged: false,
        tags: [],
        createdAt: '',
        updatedAt: ''
      }
    }
    return {
      flagged: Boolean(review.flagged),
      tags: Array.isArray(review.tags) ? review.tags.filter(Boolean) : [],
      createdAt: String(review.createdAt || ''),
      updatedAt: String(review.updatedAt || '')
    }
  }

  function canReviewMessage(detail) {
    if (!detail || typeof detail !== 'object') {
      return false
    }
    return Boolean(detail.isMailLike || detail.kind === 'appointment')
  }

  function renderReviewBadges(review) {
    const state = normalizeReviewState(review)
    const badges = []

    if (state.flagged) {
      badges.push('<span class="chip review-flag">Flagged</span>')
    }

    for (const tag of state.tags.slice(0, 2)) {
      badges.push(`<span class="chip review-tag">${escapeHtml(tag)}</span>`)
    }

    if (state.tags.length > 2) {
      badges.push(
        `<span class="chip review-tag">+${escapeHtml(String(state.tags.length - 2))}</span>`
      )
    }

    return badges.join('')
  }

  function renderReviewPanel(detail) {
    const review = normalizeReviewState(detail.review)
    if (!canReviewMessage(detail)) {
      return `
        <section class="review-panel disabled">
          <div class="review-note">Review controls are available for mail and appointment items only.</div>
        </section>
      `
    }

    const tagChips = review.tags.length
      ? review.tags
          .map(
            (tag) => `
              <button class="review-chip" type="button" data-action="remove-review-tag" data-tag="${escapeAttr(
                tag
              )}">
                <span>${escapeHtml(tag)}</span>
                <span aria-hidden="true">×</span>
              </button>
            `
          )
          .join('')
      : '<span class="review-empty">No tags yet</span>'

    return `
      <section class="review-panel" data-message-id="${escapeAttr(detail.id)}">
        <div class="review-panel-top">
          <button
            class="ghost-button small review-flag-button${review.flagged ? ' active' : ''}"
            type="button"
            data-action="toggle-review-flag"
          >
            ${review.flagged ? 'Flagged' : 'Flag'}
          </button>
          <button class="ghost-button small" type="button" data-action="clear-review">
            Clear
          </button>
          <form class="review-tag-form review-tag-form-inline" data-action="add-review-tag">
            <input
              class="review-tag-input"
              data-review-tag-input
              type="text"
              maxlength="48"
              placeholder="Add tag"
              autocomplete="off"
            />
            <button class="ghost-button small" type="submit">Add</button>
          </form>
        </div>
        <div class="review-tags">${tagChips}</div>
      </section>
    `
  }

  function applyReviewToState(messageId, review) {
    const normalized = normalizeReviewState(review)
    if (
      state.currentMessageDetail &&
      state.currentMessageDetail.id === messageId
    ) {
      state.currentMessageDetail = {
        ...state.currentMessageDetail,
        review: normalized
      }
    }

    if (state.currentFolderPage && Array.isArray(state.currentFolderPage.items)) {
      state.currentFolderPage = {
        ...state.currentFolderPage,
        items: state.currentFolderPage.items.map((item) =>
          resolveMessageId(item) === messageId
            ? {
                ...item,
                review: normalized
              }
            : item
        )
      }
    }

    if (state.currentSearchPage && Array.isArray(state.currentSearchPage.items)) {
      state.currentSearchPage = {
        ...state.currentSearchPage,
        items: state.currentSearchPage.items.map((item) =>
          resolveMessageId(item) === messageId
            ? {
                ...item,
                review: normalized
              }
            : item
        )
      }
    }
  }

  async function saveReviewState(messageId, reviewPatch, options = {}) {
    if (!state.sessionId) {
      return null
    }

    const method = options.deleteReview ? 'DELETE' : 'PATCH'
    const url = `/api/sessions/${encodeURIComponent(state.sessionId)}/messages/${encodeURIComponent(
      messageId
    )}/review`
    const payload = options.deleteReview ? null : reviewPatch
    const response = await fetchJson(url, {
      method,
      headers: options.deleteReview
        ? undefined
        : {
            'Content-Type': 'application/json'
          },
      body: options.deleteReview ? undefined : JSON.stringify(payload)
    })

    applyReviewToState(messageId, response.review)
    renderMessageList()
    renderMessageDetail()
    return response.review
  }

  async function toggleReviewFlag() {
    if (!canReviewMessage(state.currentMessageDetail)) {
      return
    }

    const review = normalizeReviewState(state.currentMessageDetail.review)
    await saveReviewState(state.currentMessageDetail.id, {
      flagged: !review.flagged
    })
  }

  async function clearReview() {
    if (!canReviewMessage(state.currentMessageDetail)) {
      return
    }
    await saveReviewState(state.currentMessageDetail.id, {}, { deleteReview: true })
  }

  async function addReviewTag(tag) {
    if (!canReviewMessage(state.currentMessageDetail)) {
      return
    }

    const normalizedTag = String(tag || '').trim()
    if (!normalizedTag) {
      return
    }

    const review = normalizeReviewState(state.currentMessageDetail.review)
    const tags = [...review.tags]
    if (!tags.some((value) => value.toLowerCase() === normalizedTag.toLowerCase())) {
      tags.push(normalizedTag)
    }
    await saveReviewState(state.currentMessageDetail.id, {
      tags
    })
  }

  async function removeReviewTag(tag) {
    if (!canReviewMessage(state.currentMessageDetail)) {
      return
    }

    const review = normalizeReviewState(state.currentMessageDetail.review)
    const nextTags = review.tags.filter((value) => value.toLowerCase() !== String(tag).toLowerCase())
    await saveReviewState(state.currentMessageDetail.id, {
      tags: nextTags
    })
  }

  async function addHiddenRule(kind, value, label) {
    const normalizedValue = normalizeHiddenRuleValue(value)
    if (!normalizedValue) {
      return
    }

    await fetchJson('/api/search/filters', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        kind,
        value: normalizedValue,
        label: label || normalizedValue
      })
    })
    await loadHiddenFilters()
    await refreshActiveVisibleMessages()
  }

  async function removeHiddenRule(filterId) {
    const normalizedId = String(filterId || '').trim()
    if (!normalizedId) {
      return
    }

    await fetchJson(`/api/search/filters/${encodeURIComponent(normalizedId)}`, {
      method: 'DELETE'
    })
    await loadHiddenFilters()
    await refreshActiveVisibleMessages()
  }

  async function refreshActiveVisibleMessages() {
    const activePage = getActivePage()
    if (!state.sessionId || !activePage) {
      return
    }

    const preferredMessageId = state.selectedMessageId || null
    await loadVisibleMessages(activePage.page || 1, {
      selectPreferred: false,
      preferredMessageId
    })
  }

  async function refreshSearchIndex() {
    const response = await fetchJson('/api/search/index/refresh', {
      method: 'POST'
    })
    await loadHiddenFilters()
    return response
  }

  function makeApiPath(...parts) {
    return parts
      .map((part) => encodeURIComponent(String(part)))
      .join('/')
  }

  function setStatus(message, tone = 'neutral') {
    ui.statusBar.textContent = message
    ui.statusBar.dataset.tone = tone
  }

  function setBodyBusy(isBusy) {
    document.body.classList.toggle('is-busy', isBusy)
  }

  function readStorageBool(key, fallback) {
    const raw = localStorage.getItem(key)
    if (raw == null) {
      return fallback
    }
    return !['0', 'false', 'no', 'off'].includes(raw.toLowerCase())
  }

  function saveState() {
    if (state.selectedCasePath) {
      localStorage.setItem(STORAGE_KEYS.casePath, state.selectedCasePath)
    } else {
      localStorage.removeItem(STORAGE_KEYS.casePath)
    }

    if (state.selectedScopePath) {
      localStorage.setItem(STORAGE_KEYS.scopePath, state.selectedScopePath)
    } else {
      localStorage.removeItem(STORAGE_KEYS.scopePath)
    }

    if (state.selectedPstFileName) {
      localStorage.setItem(STORAGE_KEYS.pstFileName, state.selectedPstFileName)
    } else {
      localStorage.removeItem(STORAGE_KEYS.pstFileName)
    }

    if (state.currentFolderId) {
      localStorage.setItem(STORAGE_KEYS.folderId, state.currentFolderId)
    } else {
      localStorage.removeItem(STORAGE_KEYS.folderId)
    }

    if (state.selectedMessageId) {
      localStorage.setItem(STORAGE_KEYS.messageId, state.selectedMessageId)
    } else {
      localStorage.removeItem(STORAGE_KEYS.messageId)
    }

    localStorage.setItem(STORAGE_KEYS.query, state.query)
    localStorage.setItem(STORAGE_KEYS.mailOnly, state.mailOnly ? '1' : '0')
    localStorage.setItem(STORAGE_KEYS.sort, state.sort)
    localStorage.setItem(
      STORAGE_KEYS.reviewFlaggedOnly,
      state.reviewFlaggedOnly ? '1' : '0'
    )
    localStorage.setItem(
      STORAGE_KEYS.reviewTaggedOnly,
      state.reviewTaggedOnly ? '1' : '0'
    )
    localStorage.setItem(STORAGE_KEYS.searchScope, state.searchScope)
    localStorage.setItem(STORAGE_KEYS.mailboxScopeView, state.mailboxScopeView)
    localStorage.setItem(STORAGE_KEYS.flaggedBundleScope, state.bundleScope)
  }

  function applyStateToControls() {
    ui.searchInput.value = state.query
    ui.mailOnlyToggle.checked = state.mailOnly
    ui.sortSelect.value = state.sort
    ui.reviewFlaggedToggle.checked = state.reviewFlaggedOnly
    ui.reviewTaggedToggle.checked = state.reviewTaggedOnly
    ui.searchScopeSelect.value = state.searchScope
    if (ui.flaggedBundleScopeSelect) {
      ui.flaggedBundleScopeSelect.value = state.bundleScope
    }
    ui.scopeSelect.value = state.mailboxScopeView === 'all' ? ALL_PSTS_SCOPE_VALUE : state.selectedScopePath || ''
    ui.pstFilter.value = state.mailboxFilter
  }

  function resetSessionState(message = 'Select a PST file from the list on the left.') {
    state.sessionId = null
    state.summary = null
    state.tree = null
    state.folderMap = new Map()
    state.currentFolderId = null
    state.currentFolderPage = null
    state.currentSearchPage = null
    state.currentMessageDetail = null
    state.selectedMessageId = null
    state.activeSearch = null
    ui.folderTree.innerHTML = '<div class="panel-empty tree-empty">No mailbox loaded.</div>'
    ui.messageList.innerHTML =
      '<div class="panel-empty">Select a folder to begin.</div>'
    ui.messageDetail.innerHTML =
      '<div class="panel-empty">Select a message to inspect it.</div>'
    ui.folderCountBadge.textContent = '0'
    ui.messageCountBadge.textContent = '0'
    ui.messageResultCount.textContent = 'Select a folder to view messages.'
    ui.pageInfo.textContent = 'Page 0 of 0'
    ui.pagePrev.disabled = true
    ui.pageNext.disabled = true
    ui.messagePrev.disabled = true
    ui.messageNext.disabled = true
    ui.sessionSummary.innerHTML =
      '<div class="summary-empty">Select a PST file from the list to load its folders and messages.</div>'
    setStatus(message, 'neutral')
    saveState()
  }

  function indexFolders(node) {
    state.folderMap.set(node.id, node)
    for (const child of node.children || []) {
      indexFolders(child)
    }
  }

  function getFolderById(folderId) {
    return state.folderMap.get(folderId) || null
  }

  function renderSummary() {
    if (!state.summary) {
      ui.sessionSummary.innerHTML =
        '<div class="summary-empty">Select a PST file from the list to load its folders and messages.</div>'
      ui.folderCountBadge.textContent = '0'
      ui.messageCountBadge.textContent = '0'
      return
    }

    const stats = state.summary.stats || {}
    const warnings = Array.isArray(state.summary.warnings)
      ? state.summary.warnings
      : []
    ui.folderCountBadge.textContent = String(stats.folderCount ?? 0)
    ui.messageCountBadge.textContent = String(stats.messageCount ?? 0)

    ui.sessionSummary.innerHTML = `
      <div class="summary-header">
        <div class="summary-title">
          <strong>${escapeHtml(state.summary.mailboxName || 'Mailbox')}</strong>
          <span class="summary-subtitle">
            ${escapeHtml(state.summary.fileName || '')}
            ${state.summary.createdAt ? `· ${escapeHtml(formatDate(state.summary.createdAt))}` : ''}
          </span>
        </div>
        <span class="chip accent">${escapeHtml(String(stats.warningCount ?? 0))} warnings</span>
      </div>
      <div class="summary-metrics">
        <div class="metric">
          <span class="metric-value">${escapeHtml(String(stats.folderCount ?? 0))}</span>
          <span class="metric-label">Folders</span>
        </div>
        <div class="metric">
          <span class="metric-value">${escapeHtml(String(stats.messageCount ?? 0))}</span>
          <span class="metric-label">Messages</span>
        </div>
        <div class="metric">
          <span class="metric-value">${escapeHtml(String(stats.mailCount ?? 0))}</span>
          <span class="metric-label">Mail items</span>
        </div>
        <div class="metric">
          <span class="metric-value">${escapeHtml(String(stats.warningCount ?? 0))}</span>
          <span class="metric-label">Warnings</span>
        </div>
      </div>
      <details class="warning-panel">
        <summary>${warnings.length ? 'View session warnings' : 'No warnings reported'}</summary>
        ${
          warnings.length
            ? `<ul class="warning-list">${warnings
                .map((warning) => `<li>${escapeHtml(warning)}</li>`)
                .join('')}</ul>`
            : ''
        }
      </details>
    `
  }

  function renderHiddenFiltersPanel() {
    if (!ui.hiddenFiltersPanel || !ui.hiddenFiltersToggle || !ui.hiddenFiltersCount) {
      return
    }

    const rules = (Array.isArray(state.hiddenRules) ? state.hiddenRules : [])
      .map(normalizeHiddenRule)
      .filter(Boolean)
    const isOpen = Boolean(state.hiddenFiltersOpen)

    ui.hiddenFiltersCount.textContent = String(rules.length)
    ui.hiddenFiltersToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false')
    ui.hiddenFiltersDropdown?.classList.toggle('is-open', isOpen)
    ui.hiddenFiltersPanel.hidden = !isOpen

    if (!isOpen) {
      return
    }

    const bindCloseHandler = () => {
      const closeButtons = ui.hiddenFiltersPanel.querySelectorAll('[data-action="close-hidden-filters"]')
      closeButtons.forEach((button) => {
        if (button.dataset.boundHiddenFiltersClose === 'true') {
          return
        }
        button.dataset.boundHiddenFiltersClose = 'true'
        button.addEventListener('click', (event) => {
          event.preventDefault()
          event.stopPropagation()
          closeHiddenFiltersDropdown()
        })
      })
    }

    if (!rules.length) {
      ui.hiddenFiltersPanel.innerHTML = `
        <div class="hidden-filters-head">
          <div class="hidden-filters-title">Hidden filters</div>
          <button
            class="ghost-button small hidden-filters-close"
            type="button"
            data-action="close-hidden-filters"
            aria-label="Close hidden filters"
          >
            ×
          </button>
        </div>
        <div class="panel-empty compact-note">
          Hidden filters will appear here. Click <strong>-</strong> on an address or subject to
          hide it globally.
        </div>
      `
      bindCloseHandler()
      return
    }

    ui.hiddenFiltersPanel.innerHTML = `
      <div class="hidden-filters-head">
        <div class="hidden-filters-title">Hidden filters</div>
        <div class="hidden-filters-head-actions">
          <span class="badge">${escapeHtml(String(rules.length))}</span>
          <button
            class="ghost-button small hidden-filters-close"
            type="button"
            data-action="close-hidden-filters"
            aria-label="Close hidden filters"
          >
            ×
          </button>
        </div>
      </div>
      <div class="hidden-filters-list">
        ${rules
          .map(
            (rule) => `
              <div class="hidden-filter-chip">
                <span class="hidden-filter-kind">${escapeHtml(rule.kind || 'filter')}</span>
                <span class="hidden-filter-value">${escapeHtml(rule.label || rule.value)}</span>
                <button
                  class="ghost-button small hidden-filter-toggle"
                  type="button"
                  data-action="remove-hidden-filter"
                  data-filter-id="${escapeAttr(rule.filterId)}"
                >
                  +
                </button>
              </div>
            `
          )
          .join('')}
      </div>
    `
    bindCloseHandler()
  }

  function setHiddenFiltersOpen(isOpen) {
    state.hiddenFiltersOpen = Boolean(isOpen)
    renderHiddenFiltersPanel()
  }

  function toggleHiddenFiltersDropdown() {
    setHiddenFiltersOpen(!state.hiddenFiltersOpen)
  }

  function closeHiddenFiltersDropdown() {
    if (!state.hiddenFiltersOpen) {
      return
    }
    setHiddenFiltersOpen(false)
  }

  async function loadHiddenFilters() {
    try {
      const response = await fetchJson('/api/search/filters')
      state.hiddenRules = Array.isArray(response.items) ? response.items : []
      state.hiddenRulesLoaded = true
      renderHiddenFiltersPanel()
    } catch (error) {
      state.hiddenRulesLoaded = true
      state.hiddenRules = []
      renderHiddenFiltersPanel()
      setStatus(`Unable to load hidden filters: ${error.message}`, 'error')
    }
  }

  function renderPstCatalog() {
    const cases = getCatalogCases()
    const selectedCase =
      cases.find((entry) => normalizeScopePath(entry.casePath) === normalizeScopePath(state.selectedCasePath)) ||
      cases[0] ||
      null
    const searches = selectedCase ? getSearchesForCase(selectedCase.casePath) : []
    const selectedSearch =
      searches.find((entry) => normalizeScopePath(entry.scopePath) === normalizeScopePath(state.selectedScopePath)) ||
      searches[0] ||
      null

    state.selectedCasePath = selectedCase ? selectedCase.casePath : ''
    state.selectedScopePath = selectedSearch ? selectedSearch.scopePath : ''
    state.selectedScopeLabel = selectedSearch ? selectedSearch.scopeLabel : getScopeLabel(state.selectedScopePath)

    const allMailboxCount = selectedCase
      ? getMailboxEntriesForCase(selectedCase.casePath, 'all').length
      : 0
    const visibleMailboxes = selectedCase
      ? getMailboxEntriesForDisplay(selectedCase.casePath, state.mailboxScopeView)
      : []
    const selectedMailboxVisible = visibleMailboxes.some(
      (entry) =>
        entry.fileName === state.selectedPstFileName &&
        normalizeScopePath(entry.scopePath) === normalizeScopePath(state.selectedScopePath)
    )

    if (state.mailboxFilter && state.selectedPstFileName && !selectedMailboxVisible) {
      state.selectedPstFileName = null
      saveState()
    }

    ui.scopeCountBadge.textContent = String(searches.length + (searches.length ? 1 : 0))
    ui.caseSelect.innerHTML = cases.length
      ? cases
          .map((entry) => {
            const isSelected =
              normalizeScopePath(entry.casePath) === normalizeScopePath(state.selectedCasePath)
            const label = `${entry.caseLabel} (${entry.searches.length})`
            return `<option value="${escapeAttr(entry.casePath)}"${
              isSelected ? ' selected' : ''
            }>${escapeHtml(label)}</option>`
          })
          .join('')
      : '<option value="">No cases found</option>'

    ui.scopeSelect.innerHTML = searches.length
      ? [
          `<option value="${escapeAttr(ALL_PSTS_SCOPE_VALUE)}"${
            state.mailboxScopeView === 'all' ? ' selected' : ''
          }>All PSTs (${escapeHtml(String(allMailboxCount))})</option>`,
          ...searches.map((entry) => {
            const isSelected =
              state.mailboxScopeView !== 'all' &&
              normalizeScopePath(entry.scopePath) === normalizeScopePath(state.selectedScopePath)
            const label = `${getSearchLabel(entry.scopePath)} (${entry.fileCount})`
            return `<option value="${escapeAttr(entry.scopePath)}"${
              isSelected ? ' selected' : ''
            }>${escapeHtml(label)}</option>`
          })
        ].join('')
      : '<option value="">No searches available</option>'

    ui.pstCountBadge.textContent = String(visibleMailboxes.length)

    if (!state.catalogLoaded) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML =
        '<strong>Loading PST files...</strong> Scanning the project <code>PST/</code> folder.'
      ui.pstList.innerHTML = ''
      return
    }

    if (!searches.length) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML = escapeHtml(state.catalogMessage || 'No searches available.')
      ui.pstList.innerHTML = ''
      return
    }

    if (!visibleMailboxes.length) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML = escapeHtml(
        state.mailboxFilter
          ? 'No PST files match the current mailbox filter.'
          : state.mailboxScopeView === 'all'
            ? 'No PST files were found for this case.'
            : state.catalogMessage || 'No PST files found.'
      )
      ui.pstList.innerHTML = ''
      return
    }

    ui.pstEmpty.classList.add('hidden')
    ui.pstList.innerHTML = visibleMailboxes
      .map((file) => {
        const isActive =
          file.fileName === state.selectedPstFileName &&
          normalizeScopePath(file.scopePath) === normalizeScopePath(state.selectedScopePath)
            ? ' active'
            : ''
        const modifiedAt = file.modifiedAt ? formatDate(file.modifiedAt) : 'Unknown date'
        const size = Number.isFinite(file.size) ? formatBytes(file.size) : 'Unknown size'
        const pathLine =
          state.mailboxScopeView === 'all' && file.displayPath
            ? `<span class="pst-item-path">${escapeHtml(file.displayPath)}</span>`
            : ''
        return `
          <button
            class="pst-item${isActive}"
            data-pst-file-name="${escapeAttr(file.fileName)}"
            data-scope-path="${escapeAttr(file.scopePath || '')}"
            title="${escapeAttr(file.fileName)}"
          >
            <span class="pst-item-name">${escapeHtml(file.fileName)}</span>
            ${pathLine}
            <span class="pst-item-meta">${escapeHtml(size)} · ${escapeHtml(modifiedAt)}</span>
          </button>
        `
      })
      .join('')
  }

  function collectFoldersWithContent(node, depth = 0, output = []) {
    if (!node) {
      return output
    }

    if (Number(node.indexedMessageCount || 0) > 0) {
      output.push({
        id: node.id,
        displayName: node.displayName,
        path: node.path,
        indexedMessageCount: node.indexedMessageCount || 0,
        mailMessageCount: node.mailMessageCount || 0,
        depth
      })
    }

    for (const child of node.children || []) {
      collectFoldersWithContent(child, depth + 1, output)
    }

    return output
  }

  function renderFolderTree() {
    if (!state.tree) {
      ui.folderTree.innerHTML =
        '<div class="panel-empty tree-empty">No mailbox loaded.</div>'
      return
    }

    const folders = collectFoldersWithContent(state.tree)
    if (!folders.length) {
      ui.folderTree.innerHTML =
        '<div class="panel-empty tree-empty">No folders with content were found.</div>'
      return
    }

    ui.folderTree.innerHTML = `
      <ul class="folder-list flat-folder-list">
        ${folders
          .map((node) => {
            const isActive = node.id === state.currentFolderId ? ' active' : ''
            const badges = [
              `<span class="folder-badge" title="Indexed messages">${escapeHtml(
                String(node.indexedMessageCount ?? 0)
              )}</span>`,
              `<span class="folder-badge" title="Mail-like messages">${escapeHtml(
                String(node.mailMessageCount ?? 0)
              )}</span>`
            ].join('')

            return `
              <li class="folder-item">
                <button
                  class="folder-button${isActive}"
                  style="--folder-depth:${Number(node.depth || 0)}"
                  data-folder-id="${escapeAttr(node.id)}"
                  title="${escapeAttr(node.path || node.displayName)}"
                >
                  <span class="folder-name">${escapeHtml(node.displayName || '(untitled)')}</span>
                  <span class="folder-badges">${badges}</span>
                </button>
              </li>
            `
          })
          .join('')}
      </ul>
    `
  }

  function renderMessageRow(item) {
    const messageId = resolveMessageId(item)
    const isActive = messageId && messageId === state.selectedMessageId ? ' active' : ''
    const isSearchResult = item.scopeLabel ? ' search-result-row' : ''
    const sender = item.senderName || item.senderEmailAddress || '(unknown sender)'
    const time = item.sortDate || item.creationTime || item.clientSubmitTime
    const review = normalizeReviewState(item.review)
    const subjectHide = renderInlineHideButton('subject', item.subject, item.subject)
    const senderHide = item.senderEmailAddress
      ? renderInlineHideButton('address', item.senderEmailAddress, item.senderEmailAddress)
      : ''
    const chips = [
      `<span class="chip ${item.kind === 'mail' ? 'green' : 'accent'}">${escapeHtml(
        item.kind || 'other'
      )}</span>`,
      item.hasAttachments ? '<span class="chip accent">Attachments</span>' : '',
      item.isRead ? '<span class="chip">Read</span>' : '<span class="chip danger">Unread</span>',
      renderReviewBadges(review),
      item.parseError ? '<span class="chip danger">Parse error</span>' : ''
    ]
      .filter(Boolean)
      .join('')
    const contextLine =
      item.scopeLabel || item.fileName
        ? `
          <div class="message-context">
            ${item.scopeLabel ? `<span>${escapeHtml(item.scopeLabel)}</span>` : ''}
            ${
              item.fileName
                ? `<span>${escapeHtml(item.fileName)}</span>`
                : ''
            }
            ${
              item.folderPath
                ? `<span>${escapeHtml(item.folderPath)}</span>`
                : ''
            }
          </div>
        `
        : ''

    return `
      <article
        class="message-row${isActive}${isSearchResult}"
        tabindex="0"
        role="button"
        data-message-id="${escapeAttr(messageId)}"
        ${item.scopePath ? `data-scope-path="${escapeAttr(item.scopePath)}"` : ''}
        ${item.fileName ? `data-file-name="${escapeAttr(item.fileName)}"` : ''}
        ${item.folderId ? `data-folder-id="${escapeAttr(item.folderId)}"` : ''}
      >
        <div class="message-row-top">
          <div class="message-subject">
            <span class="message-subject-text">${escapeHtml(item.subject || '(no subject)')}</span>
            ${subjectHide}
          </div>
          <div class="message-date">${escapeHtml(formatDate(time))}</div>
        </div>
        <div class="message-sender">
          <span class="message-sender-text">${escapeHtml(sender)}</span>
          ${senderHide}
        </div>
        ${contextLine}
        <div class="message-row-meta">
          ${chips}
          <span>${escapeHtml(item.recipientText || item.displayTo || 'No recipients')}</span>
        </div>
      </article>
    `
  }

  function getActivePage() {
    return state.currentSearchPage || state.currentFolderPage
  }

  function isSearchResultsActive() {
    return Boolean(state.currentSearchPage)
  }

  function updatePagingButtons() {
    const page = getActivePage()
    const hasPage = Boolean(page)
    ui.pagePrev.disabled = !hasPage || page.page <= 1
    ui.pageNext.disabled = !hasPage || page.page >= page.totalPages

    const items = page ? page.items || [] : []
    const currentIndex = state.selectedMessageId
      ? items.findIndex((item) => resolveMessageId(item) === state.selectedMessageId)
      : -1
    ui.messagePrev.disabled =
      !hasPage || (currentIndex <= 0 && page.page <= 1)
    ui.messageNext.disabled =
      !hasPage ||
      (currentIndex === -1
        ? true
        : currentIndex >= items.length - 1 && page.page >= page.totalPages)
  }

  function renderMessageList() {
    const page = getActivePage()
    if (!page) {
      ui.messageList.innerHTML =
        '<div class="panel-empty">Select a folder to view messages.</div>'
      ui.messageResultCount.textContent = 'Select a folder to view messages.'
      ui.pageInfo.textContent = 'Page 0 of 0'
      updatePagingButtons()
      return
    }

    ui.messageCountBadge.textContent = String(state.summary?.stats?.messageCount ?? 0)
    ui.messageList.innerHTML = page.items.length
      ? page.items.map((item) => renderMessageRow(item)).join('')
      : '<div class="panel-empty">No messages match the current filters.</div>'

    const folderName = page.folder?.displayName || page.scopeLabel || 'folder'
    const queryLabel = page.query ? ` filtered by "${page.query}"` : ''
    ui.messageResultCount.textContent = page.items.length
      ? `Showing ${page.items.length} of ${page.total} messages in ${folderName}${queryLabel}.`
      : `No messages found in ${folderName}${queryLabel}.`
    ui.pageInfo.textContent = `Page ${page.page} of ${page.totalPages}`
    updatePagingButtons()
  }

  function buildHtmlFrame(html) {
    const srcdoc = `<!doctype html>
      <html>
        <head>
          <meta charset="utf-8" />
          <base target="_blank" />
          <style>
            :root {
              color-scheme: light;
            }
            body {
              margin: 0;
              padding: 8px 4px 12px;
              font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
              font-size: 13px;
              line-height: 1.5;
              color: #1f2a37;
              background: #ffffff;
            }
            img { max-width: 100%; height: auto; }
            table { border-collapse: collapse; }
            a { color: #2b6cb0; }
            blockquote {
              border-left: 3px solid #c2ccd9;
              margin-left: 0;
              padding-left: 1rem;
              color: #526072;
            }
          </style>
        </head>
        <body>${html}</body>
      </html>`
    return `<iframe class="body-frame" sandbox="allow-popups" srcdoc="${escapeAttr(srcdoc)}"></iframe>`
  }

  function renderBodySection(detail) {
    const htmlBody = detail.bodyHtml || ''
    const textBody = detail.bodyText || detail.bodyPrefix || detail.parseError || ''
    const hasHtml = hasText(htmlBody)
    const hasTextBody = hasText(textBody)

    if (!hasHtml && !hasTextBody) {
      return '<div class="plain-body">No message body is available.</div>'
    }

    return `
      <div class="body-panel">
        ${
          hasHtml
            ? buildHtmlFrame(htmlBody)
            : `<pre class="plain-body">${escapeHtml(textBody)}</pre>`
        }
      </div>
    `
  }

  function renderAttachments(detail) {
    if (!Array.isArray(detail.attachments) || !detail.attachments.length) {
      return ''
    }

    return `
      <div class="attachments-block">
        <div class="attachments-label">Attachments</div>
        <div class="attachments">
          ${detail.attachments.map((attachment) => renderAttachment(attachment)).join('')}
        </div>
      </div>
    `
  }

  function renderAttachment(attachment) {
    const name =
      attachment.downloadFilename ||
      attachment.longFilename ||
      attachment.filename ||
      `Attachment ${attachment.index + 1}`
    const meta = [
      attachment.mimeTag ? escapeHtml(attachment.mimeTag) : '',
      Number.isFinite(attachment.size) && attachment.size > 0
        ? escapeHtml(formatBytes(attachment.size))
        : '',
      attachment.isEmbeddedMessage ? 'Email' : '',
      attachment.parseError ? 'Error' : ''
    ]
      .filter(Boolean)
      .join(' · ')

    const downloadLink = attachment.downloadUrl
      ? `<a class="attachment-link" href="${escapeAttr(attachment.downloadUrl)}" download="${escapeAttr(
          name
        )}">Download</a>`
      : ''

    return `
      <article class="attachment-item">
        <div class="attachment-main">
          <div class="attachment-icon">📎</div>
          <div class="attachment-copy">
            <div class="attachment-name">${escapeHtml(name)}</div>
            <div class="attachment-meta">${meta}</div>
            ${
              attachment.parseError
                ? `<div class="attachment-error">${escapeHtml(attachment.parseError)}</div>`
                : ''
            }
          </div>
        </div>
        ${downloadLink}
      </article>
    `
  }

  function renderRecipientTokens(value, allowHide = false) {
    const tokens = splitRecipientTokens(value)
    if (!tokens.length) {
      return ''
    }

    return tokens
      .map((token) => {
        const email = extractEmailAddress(token)
        const hideButton = allowHide && email ? renderInlineHideButton('address', email, email) : ''
        return `
          <span class="recipient-chip">
            <span class="recipient-chip-text">${escapeHtml(token)}</span>
            ${hideButton}
          </span>
        `
      })
      .join('')
  }

  function renderDetailCard(detail) {
    if (!detail) {
      return '<div class="panel-empty">Select a message to inspect it.</div>'
    }
    const actions = `
      <div class="detail-actions">
        <button class="ghost-button small" type="button" data-action="download-json">
          JSON
        </button>
        <button class="ghost-button small" type="button" data-action="download-eml">
          EML
        </button>
      </div>
    `
    const subjectHide = renderInlineHideButton('subject', detail.subject, detail.subject)

    const bodySection = renderBodySection({
      ...detail,
      bodyHtml: detail.bodyHtml,
      bodyText: detail.bodyText,
      parseError: detail.parseError
    })
    const senderName = detail.senderName || detail.senderEmailAddress || 'Unknown sender'
    const senderEmail =
      detail.senderEmailAddress && detail.senderEmailAddress !== senderName
        ? detail.senderEmailAddress
        : ''
    const sentTime = formatDate(detail.sortDate || detail.clientSubmitTime || detail.creationTime)
    const reviewPanel = renderReviewPanel(detail)

    return `
      <article class="detail-card" data-message-id="${escapeAttr(detail.id)}">
        <header class="outlook-header">
          <h3 class="detail-title">
            <span class="detail-title-text">${escapeHtml(detail.subject || '(no subject)')}</span>
            ${subjectHide}
          </h3>
          ${actions}
        </header>

        ${
          detail.parseError
            ? `<div class="message-warning"><strong>Load warning:</strong> ${escapeHtml(detail.parseError)}</div>`
            : ''
        }

        <section class="sender-strip">
          <div class="sender-avatar" aria-hidden="true">${escapeHtml(getSenderInitials(detail))}</div>
          <div class="sender-copy">
            <div class="sender-line">
              <span class="sender-name">${escapeHtml(senderName)}</span>
              ${
              senderEmail
                  ? `<span class="sender-address">&lt;${escapeHtml(senderEmail)}&gt;${renderInlineHideButton('address', senderEmail, senderEmail)}</span>`
                  : ''
              }
            </div>
            ${
              detail.resolvedDisplayTo || detail.displayTo
                ? `
                  <div class="recipient-line">
                    <span class="recipient-label">To</span>
                    <span class="recipient-values">${renderRecipientTokens(
                      detail.resolvedDisplayTo || detail.displayTo,
                      true
                    )}</span>
                  </div>
                `
                : ''
            }
            ${
              detail.resolvedDisplayCC || detail.displayCC
                ? `
                  <div class="recipient-line">
                    <span class="recipient-label">Cc</span>
                    <span class="recipient-values">${renderRecipientTokens(
                      detail.resolvedDisplayCC || detail.displayCC,
                      true
                    )}</span>
                  </div>
                `
                : ''
            }
            ${
              detail.resolvedDisplayBCC || detail.displayBCC
                ? `
                  <div class="recipient-line">
                    <span class="recipient-label">Bcc</span>
                    <span class="recipient-values">${renderRecipientTokens(
                      detail.resolvedDisplayBCC || detail.displayBCC,
                      true
                    )}</span>
                  </div>
                `
                : ''
            }
          </div>
          <div class="sent-time">${escapeHtml(sentTime)}</div>
        </section>

        ${reviewPanel}

        <section class="message-body">
          ${bodySection}
        </section>

        ${renderAttachments(detail)}
      </article>
    `
  }

  function renderMessageDetail() {
    if (!state.currentMessageDetail) {
      ui.messageDetail.innerHTML =
        '<div class="panel-empty">Select a message to inspect it.</div>'
      updatePagingButtons()
      return
    }

    ui.messageDetail.innerHTML = renderDetailCard(state.currentMessageDetail)
    updatePagingButtons()
  }

  function setActiveFolder(folderId) {
    state.currentFolderId = folderId
    renderFolderTree()
  }

  async function fetchJson(url, init = {}) {
    const response = await fetch(url, {
      headers: {
        Accept: 'application/json',
        ...(init.headers || {})
      },
      ...init
    })

    const contentType = response.headers.get('content-type') || ''
    const payload = contentType.includes('application/json')
      ? await response.json()
      : await response.text()

    if (!response.ok) {
      const message =
        (payload && typeof payload === 'object' && payload.error) ||
        (typeof payload === 'string' ? payload : '') ||
        response.statusText ||
        'Request failed'
      throw new Error(message)
    }

    return payload
  }

  async function openMailbox(fileName, options = {}) {
    const showBusy = options.showBusy !== false
    if (showBusy) {
      setBodyBusy(true)
    }
    const scopePath = normalizeScopePath(options.scopePath || state.selectedScopePath || '')
    setStatus(`Opening ${fileName}...`)
    try {
      const response = await fetchJson('/api/psts/open', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          scopePath,
          fileName
        })
      })

      state.sessionId = response.sessionId
      state.selectedPstFileName = response.fileName || fileName
      state.selectedCasePath = getCasePathFromScopePath(response.scopePath || scopePath)
      state.selectedScopePath = normalizeScopePath(response.scopePath || scopePath)
      state.selectedScopeLabel = response.scopeLabel || getScopeLabel(state.selectedScopePath)
      state.summary = response.summary
      state.tree = response.tree
      state.folderMap = new Map()
      indexFolders(state.tree)
      const restoreFolderId = options.restoreFolderId || localStorage.getItem(STORAGE_KEYS.folderId)
      const restoreMessageId =
        options.restoreMessageId || localStorage.getItem(STORAGE_KEYS.messageId)
      state.currentFolderId =
        restoreFolderId && state.folderMap.has(restoreFolderId)
          ? restoreFolderId
          : state.summary.rootFolderId || state.tree.id
      state.currentFolderPage = null
      state.currentSearchPage = null
      state.currentMessageDetail = null
      state.selectedMessageId = restoreMessageId || null
      renderSummary()
      renderFolderTree()
      renderPstCatalog()
      saveState()
      setStatus(`Loaded ${fileName}.`, 'success')
      await loadFolderPage(1, {
        selectPreferred: true,
        preferredMessageId: state.selectedMessageId
      })
    } catch (error) {
      const hadSession = Boolean(state.sessionId)
      setStatus(`Unable to open ${fileName}: ${error.message}`, 'error')
      if (!hadSession) {
        ui.messageDetail.innerHTML = `
          <div class="panel-empty">
            <strong>Mailbox failed to open.</strong> ${escapeHtml(error.message)}
          </div>
        `
        ui.messageList.innerHTML =
          '<div class="panel-empty">Select a folder to view messages.</div>'
      }
    } finally {
      if (showBusy) {
        setBodyBusy(false)
      }
    }
  }

  async function loadMailboxCatalog(options = {}) {
    const showBusy = options.showBusy !== false
    const hadSession = Boolean(state.sessionId)
    if (showBusy) {
      setBodyBusy(true)
    }
    setStatus('Loading PST catalog...')
    try {
      const casePath = normalizeScopePath(
        options.casePath ||
          state.selectedCasePath ||
          localStorage.getItem(STORAGE_KEYS.casePath) ||
          ''
      )
      const scopePath = normalizeScopePath(
        options.scopePath ||
          state.selectedScopePath ||
          localStorage.getItem(STORAGE_KEYS.scopePath) ||
          ''
      )
      const response = await fetchJson(
        `/api/psts${scopePath ? `?scopePath=${encodeURIComponent(scopePath)}` : ''}`
      )
      state.catalogLoaded = true
      state.catalogMessage = response.message || ''
      state.catalogScopes = Array.isArray(response.scopes) ? response.scopes : []
      state.catalog = Array.isArray(response.files) ? response.files : []
      state.selectedCasePath = normalizeScopePath(casePath || getCasePathFromScopePath(response.scopePath))
      state.selectedScopePath = normalizeScopePath(response.scopePath || scopePath)
      state.selectedScopeLabel = response.scopeLabel || getScopeLabel(state.selectedScopePath)
      renderPstCatalog()

      const preferredFileName =
        options.preferredFileName ||
        state.selectedPstFileName ||
        localStorage.getItem(STORAGE_KEYS.pstFileName) ||
        ''
      const preferredScopePath = normalizeScopePath(
        options.preferredScopePath ||
          state.selectedScopePath ||
          localStorage.getItem(STORAGE_KEYS.scopePath) ||
          ''
      )
      const mailboxEntries = getMailboxEntriesForCase(
        state.selectedCasePath || casePath || getCasePathFromScopePath(response.scopePath),
        state.mailboxScopeView
      )
      const preferred = preferredFileName
        ? mailboxEntries.find(
            (item) =>
              item.fileName === preferredFileName &&
              normalizeScopePath(item.scopePath) === preferredScopePath
          ) ||
          mailboxEntries.find((item) => item.fileName === preferredFileName)
        : null
      const nextMailbox = preferred || mailboxEntries[0] || null

      if (!nextMailbox) {
        renderPstCatalog()
        if (hadSession) {
          setStatus(response.message || 'No PST files were found in the PST folder.', 'neutral')
          return
        }
        resetSessionState(response.message || 'Select a PST file from the list on the left.')
        ui.sessionSummary.innerHTML =
          '<div class="summary-empty">No PST files are available in the project PST folder.</div>'
        renderPstCatalog()
        return
      }

      await openMailbox(nextMailbox.fileName, {
        showBusy: false,
        scopePath: nextMailbox.scopePath || state.selectedScopePath || scopePath || '',
        restoreFolderId: localStorage.getItem(STORAGE_KEYS.folderId) || undefined,
        restoreMessageId: localStorage.getItem(STORAGE_KEYS.messageId) || undefined
      })
    } catch (error) {
      state.catalogLoaded = true
      state.catalog = []
      state.catalogMessage = error.message
      renderPstCatalog()
      if (!hadSession) {
        resetSessionState(`Unable to load the PST catalog: ${error.message}`)
        ui.sessionSummary.innerHTML = `
          <div class="summary-empty">
            <strong>Catalog unavailable.</strong> ${escapeHtml(error.message)}
          </div>
        `
      }
      setStatus(`Unable to load the PST catalog: ${error.message}`, 'error')
    } finally {
      if (showBusy) {
        setBodyBusy(false)
      }
    }
  }

  async function loadFolderPage(page = 1, options = {}) {
    if (!state.sessionId || !state.currentFolderId) {
      return
    }

    const token = (state.folderLoadToken = (state.folderLoadToken || 0) + 1)
    const folderId = state.currentFolderId
    const params = new URLSearchParams()
    params.set('page', String(page))
    params.set('pageSize', String(state.pageSize))
    params.set('mailOnly', state.mailOnly ? '1' : '0')
    params.set('sort', state.sort)
    if (state.reviewFlaggedOnly) {
      params.set('reviewFlagged', '1')
    }
    if (state.reviewTaggedOnly) {
      params.set('reviewTagged', '1')
    }
    if (options.query) {
      params.set('q', String(options.query).trim())
    }

    setStatus(`Loading messages for ${getFolderById(folderId)?.displayName || 'folder'}...`)
    try {
      const response = await fetchJson(
        `/api/sessions/${encodeURIComponent(state.sessionId)}/folders/${encodeURIComponent(
          folderId
        )}/messages?${params.toString()}`
      )

      if (token !== state.folderLoadToken) {
        return
      }

      state.currentSearchPage = null
      state.currentFolderPage = response.page
      renderMessageList()

      const preferredMessageId =
        options.preferredMessageId || localStorage.getItem(STORAGE_KEYS.messageId) || null
      const pageItems = response.page?.items || []
      const preferredVisible = preferredMessageId
        ? pageItems.find((item) => item.id === preferredMessageId)
        : null

      if (preferredVisible) {
        await selectMessage(preferredVisible.id, { preserveBodyMode: true })
      } else if (options.selectPreferred !== false && pageItems.length) {
        await selectMessage(pageItems[0].id, { preserveBodyMode: false })
      } else {
        state.selectedMessageId = null
        state.currentMessageDetail = null
        ui.messageDetail.innerHTML =
          '<div class="panel-empty">Select a message to inspect it.</div>'
        saveState()
        renderMessageList()
      }

      const currentFolder = getFolderById(folderId)
      const pageTotal = response.page?.total ?? 0
      setStatus(
        `${currentFolder ? currentFolder.displayName : 'Folder'} loaded: ${pageTotal} message(s).`,
        'success'
      )
    } catch (error) {
      setStatus(`Unable to load messages: ${error.message}`, 'error')
      ui.messageList.innerHTML = `
        <div class="panel-empty">
          <strong>Unable to load messages.</strong> ${escapeHtml(error.message)}
        </div>
      `
    }
  }

  async function loadSearchResults(page = 1, options = {}) {
    const criteria = options.criteria || state.activeSearch
    if (!criteria || !hasSearchCriteria(criteria)) {
      state.activeSearch = null
      state.currentSearchPage = null
      return loadFolderPage(page, options)
    }

    const scope = criteria.searchScope || state.searchScope || 'pst'
    if (scope === 'pst' && !state.sessionId) {
      return
    }
    const params = new URLSearchParams()
    params.set('scope', scope)
    params.set('query', String(criteria.query || '').trim())
    params.set('mode', criteria.mode || deriveSearchMode(criteria.query))
    params.set('page', String(page))
    params.set('pageSize', String(state.pageSize))
    params.set('mailOnly', state.mailOnly ? '1' : '0')
    params.set('sort', state.sort)
    if (state.reviewFlaggedOnly) {
      params.set('reviewFlagged', '1')
    }
    if (state.reviewTaggedOnly) {
      params.set('reviewTagged', '1')
    }
    const searchScopePath = String(
      scope === 'search'
        ? criteria.scopePath || state.selectedScopePath || ''
        : criteria.scopePath || ''
    ).trim()
    if (searchScopePath) {
      params.set('scopePath', searchScopePath)
    }
    if (scope === 'pst' && state.sessionId) {
      params.set('sessionId', state.sessionId)
    }

    setStatus(
      `Searching ${
        scope === 'all'
          ? 'all cases/searches'
          : scope === 'search'
            ? getScopeLabel(criteria.scopePath || state.selectedScopePath || '') || state.selectedScopeLabel || 'selected search'
            : 'selected PST'
      }...`
    )
    try {
      const response = await fetchJson(`/api/search?${params.toString()}`)
      const normalizedPage = normalizeSearchResultsPage(response.page)
      state.currentFolderPage = null
      state.currentSearchPage = normalizedPage
      state.hiddenRules = Array.isArray(response.page?.hiddenRules)
        ? response.page.hiddenRules
        : state.hiddenRules
      state.hiddenRulesLoaded = true
      state.activeSearch = {
        query: String(response.page?.query || criteria.query || '').trim(),
        searchScope: scope,
        scopePath:
          scope === 'search'
            ? String(response.page?.scopePath || criteria.scopePath || state.selectedScopePath || '').trim()
            : String(criteria.scopePath || '').trim(),
        mode: String(response.page?.mode || criteria.mode || deriveSearchMode(criteria.query) || 'and')
      }
      renderHiddenFiltersPanel()
      renderMessageList()

      const pageItems = normalizedPage?.items || []
      const preferredMessageId =
        options.preferredMessageId || localStorage.getItem(STORAGE_KEYS.messageId) || null
      const preferredVisible = preferredMessageId
        ? pageItems.find((item) => resolveMessageId(item) === preferredMessageId)
        : null

      if (scope === 'pst' && preferredVisible) {
        await selectMessage(resolveMessageId(preferredVisible), { refresh: true })
      } else if (
        scope === 'pst' &&
        options.selectPreferred !== false &&
        pageItems.length &&
        pageItems[0].fileName === state.selectedPstFileName
      ) {
        await selectMessage(resolveMessageId(pageItems[0]), { refresh: true })
      } else {
        state.selectedMessageId = null
        state.currentMessageDetail = null
        ui.messageDetail.innerHTML =
          '<div class="panel-empty">Select a result to inspect it.</div>'
        saveState()
        renderMessageList()
      }

      setStatus(
        `Search completed: ${response.page?.total ?? 0} message(s) found in ${
          response.page?.scopeLabel || 'selected scope'
        }.`,
        'success'
      )
    } catch (error) {
      setStatus(`Unable to search messages: ${error.message}`, 'error')
      ui.messageList.innerHTML = `
        <div class="panel-empty">
          <strong>Unable to search messages.</strong> ${escapeHtml(error.message)}
        </div>
      `
    }
  }

  async function loadVisibleMessages(page = 1, options = {}) {
    if (state.currentSearchPage && state.activeSearch && hasSearchCriteria(state.activeSearch)) {
      return loadSearchResults(page, options)
    }
    return loadFolderPage(page, options)
  }

  async function executeSearch(page = 1, options = {}) {
    const draft = syncSearchDraftToState()
    const criteria = {
      query: draft.query,
      searchScope: draft.searchScope,
      scopePath: draft.searchScope === 'search' ? state.selectedScopePath || '' : '',
      mode: draft.mode
    }

    if (!hasSearchCriteria(criteria)) {
      state.activeSearch = null
      state.currentSearchPage = null
      await loadFolderPage(page, options)
      return
    }

    await loadSearchResults(page, {
      ...options,
      criteria
    })
  }

  async function selectMessage(messageId, options = {}) {
    if (!state.sessionId) {
      return
    }

    if (state.selectedMessageId === messageId && state.currentMessageDetail && !options.refresh) {
      renderMessageDetail()
      return
    }

    const token = (state.detailLoadToken = (state.detailLoadToken || 0) + 1)
    setStatus(`Loading message ${messageId}...`)
    try {
      const response = await fetchJson(
        `/api/sessions/${encodeURIComponent(state.sessionId)}/messages/${encodeURIComponent(
          messageId
        )}`
      )
      if (token !== state.detailLoadToken) {
        return
      }

      state.selectedMessageId = messageId
      state.currentMessageDetail = response.detail
      renderMessageList()
      renderMessageDetail()
      saveState()
      setStatus(`Loaded message "${response.detail.subject || '(no subject)'}".`, 'success')
    } catch (error) {
      setStatus(`Unable to load message: ${error.message}`, 'error')
      ui.messageDetail.innerHTML = `
        <div class="panel-empty">
          <strong>Unable to load message.</strong> ${escapeHtml(error.message)}
        </div>
      `
    }
  }

  async function downloadCurrentMessage(kind) {
    if (!state.sessionId || !state.selectedMessageId) {
      return
    }
    const route =
      kind === 'json'
        ? 'export.json'
        : 'export.eml'
    const url = `/api/sessions/${encodeURIComponent(state.sessionId)}/messages/${encodeURIComponent(
      state.selectedMessageId
    )}/${route}`
    window.location.assign(url)
  }

  function buildFlaggedBundleDownloadUrl() {
    const params = new URLSearchParams()
    const scope = String(state.bundleScope || 'all')
    params.set('scope', scope)

    if (scope === 'pst') {
      if (!state.sessionId) {
        return ''
      }
      params.set('sessionId', state.sessionId)
    } else if (scope === 'search') {
      const scopePath = normalizeScopePath(state.selectedScopePath || '')
      if (scopePath) {
        params.set('scopePath', scopePath)
      }
    }

    return `/api/exports/flagged.zip?${params.toString()}`
  }

  function downloadFlaggedBundle() {
    const url = buildFlaggedBundleDownloadUrl()
    if (!url) {
      setStatus('Select a PST session before downloading a selected-PST bundle.', 'error')
      return
    }
    window.location.assign(url)
  }

  async function navigateFolderPage(delta) {
    const page = getActivePage()
    if (!page) {
      return
    }
    const nextPage = page.page + delta
    if (nextPage < 1 || nextPage > page.totalPages) {
      return
    }
    await loadVisibleMessages(nextPage, { selectPreferred: true })
  }

  async function navigateMessage(delta) {
    const page = getActivePage()
    if (!page || !state.selectedMessageId) {
      return
    }

    const items = page.items || []
    const index = items.findIndex((item) => resolveMessageId(item) === state.selectedMessageId)
    if (index >= 0) {
      const targetIndex = index + delta
      if (targetIndex >= 0 && targetIndex < items.length) {
        await selectMessage(resolveMessageId(items[targetIndex]), { refresh: true })
        return
      }
    }

    const nextPage = page.page + (delta > 0 ? 1 : -1)
    if (nextPage < 1 || nextPage > page.totalPages) {
      return
    }

    await loadVisibleMessages(nextPage, { selectPreferred: true })
  }

  async function openSearchResult(item) {
    const messageId = resolveMessageId(item)
    if (!item || !item.fileName || !messageId) {
      return
    }

    const currentScopePath = normalizeScopePath(state.selectedScopePath || '')
    const targetScopePath = normalizeScopePath(item.scopePath || '')
    const isCurrentMailbox =
      state.sessionId &&
      item.fileName === state.selectedPstFileName &&
      targetScopePath === currentScopePath

    if (isCurrentMailbox) {
      await selectMessage(messageId, { refresh: true })
      return
    }

    await openMailbox(item.fileName, {
      showBusy: true,
      scopePath: targetScopePath,
      restoreFolderId: item.folderId,
      restoreMessageId: messageId
    })
    await selectMessage(messageId, { refresh: true })
  }

  async function chooseFolder(folderId) {
    if (!folderId || folderId === state.currentFolderId) {
      return
    }
    state.currentFolderId = folderId
    state.currentFolderPage = null
    state.currentSearchPage = null
    state.currentMessageDetail = null
    state.selectedMessageId = null
    renderFolderTree()
    ui.messageDetail.innerHTML =
      '<div class="panel-empty">Select a message to inspect it.</div>'
    saveState()
    await loadFolderPage(1, { selectPreferred: true })
  }

  function refreshControls() {
    applyStateToControls()
    updatePagingButtons()
  }

  function wireEvents() {
    ui.refreshCatalog.addEventListener('click', () => {
      void (async () => {
        setStatus('Refreshing mailbox catalog and search cache...')
        try {
          await refreshSearchIndex()
        } catch (error) {
          setStatus(`Unable to refresh search cache: ${error.message}`, 'error')
        }
        await loadMailboxCatalog({
          showBusy: true,
          casePath: state.selectedCasePath || undefined,
          scopePath: state.selectedScopePath || undefined,
          preferredFileName: state.selectedPstFileName || undefined
        })
      })()
    })

    ui.caseSelect.addEventListener('change', () => {
      const nextCasePath = normalizeScopePath(ui.caseSelect.value)
      const nextSearch = getSearchesForCase(nextCasePath)[0] || null
      state.selectedCasePath = nextCasePath
      state.selectedScopePath = nextSearch ? nextSearch.scopePath : ''
      state.selectedScopeLabel = nextSearch ? nextSearch.scopeLabel : getScopeLabel(state.selectedScopePath)
      saveState()
      void loadMailboxCatalog({
        showBusy: true,
        casePath: state.selectedCasePath || undefined,
        scopePath: state.selectedScopePath || undefined,
        preferredFileName: state.selectedPstFileName || undefined,
        preferredScopePath: state.selectedScopePath || undefined
      })
    })

    ui.scopeSelect.addEventListener('change', () => {
      const nextScopeValue = String(ui.scopeSelect.value || '')
      if (isAllPstsScopeValue(nextScopeValue)) {
        state.mailboxScopeView = 'all'
      } else {
        const nextScopePath = normalizeScopePath(nextScopeValue)
        const nextScope = getCatalogScope(nextScopePath)
        state.mailboxScopeView = 'search'
        state.selectedScopePath = nextScopePath
        state.selectedCasePath = getCasePathFromScopePath(nextScopePath)
        state.selectedScopeLabel = nextScope ? nextScope.scopeLabel : getScopeLabel(nextScopePath)
      }
      saveState()
      void loadMailboxCatalog({
        showBusy: true,
        casePath: state.selectedCasePath || undefined,
        scopePath: state.selectedScopePath || undefined,
        preferredFileName: state.selectedPstFileName || undefined,
        preferredScopePath: state.selectedScopePath || undefined
      })
    })

    ui.searchScopeSelect.addEventListener('change', () => {
      state.searchScope = ui.searchScopeSelect.value || 'pst'
      saveState()
    })

    if (ui.flaggedBundleScopeSelect) {
      ui.flaggedBundleScopeSelect.addEventListener('change', () => {
        state.bundleScope = ui.flaggedBundleScopeSelect.value || 'all'
        saveState()
      })
    }

    if (ui.flaggedBundleButton) {
      ui.flaggedBundleButton.addEventListener('click', () => {
        downloadFlaggedBundle()
      })
    }

    ui.pstList.addEventListener('click', (event) => {
      const button = event.target.closest('[data-pst-file-name]')
      if (!button) {
        return
      }
      const fileName = button.dataset.pstFileName
      const scopePath = normalizeScopePath(button.dataset.scopePath || '')
      if (!fileName) {
        return
      }
      void openMailbox(fileName, { showBusy: true, scopePath })
    })

    ui.folderTree.addEventListener('click', (event) => {
      const button = event.target.closest('[data-folder-id]')
      if (!button) {
        return
      }
      void chooseFolder(button.dataset.folderId)
    })

    ui.messageList.addEventListener('click', (event) => {
      const actionButton = event.target.closest('[data-action]')
      if (actionButton && actionButton.dataset.action) {
        const action = actionButton.dataset.action
        if (action === 'hide-subject') {
          void addHiddenRule('subject', actionButton.dataset.filterValue || '', actionButton.dataset.filterLabel || '')
        } else if (action === 'hide-address') {
          void addHiddenRule('address', actionButton.dataset.filterValue || '', actionButton.dataset.filterLabel || '')
        }
        return
      }

      const button = event.target.closest('[data-message-id]')
      if (!button) {
        return
      }
      if (button.dataset.fileName && button.dataset.scopePath) {
        void openSearchResult({
          id: button.dataset.messageId,
          messageId: button.dataset.messageId,
          fileName: button.dataset.fileName,
          scopePath: button.dataset.scopePath,
          folderId: button.dataset.folderId || '',
          displayName: ''
        })
        return
      }
      void selectMessage(button.dataset.messageId, { refresh: true })
    })

    ui.messageList.addEventListener('keydown', (event) => {
      const row = event.target.closest('[data-message-id]')
      if (!row) {
        return
      }
      if (event.key !== 'Enter' && event.key !== ' ') {
        return
      }
      event.preventDefault()
      if (row.dataset.fileName && row.dataset.scopePath) {
        void openSearchResult({
          id: row.dataset.messageId,
          messageId: row.dataset.messageId,
          fileName: row.dataset.fileName,
          scopePath: row.dataset.scopePath,
          folderId: row.dataset.folderId || '',
          displayName: ''
        })
        return
      }
      void selectMessage(row.dataset.messageId, { refresh: true })
    })

    ui.messageDetail.addEventListener('click', (event) => {
      const actionButton = event.target.closest('[data-action]')
      if (!actionButton) {
        return
      }

      const action = actionButton.dataset.action
      if (action === 'hide-subject') {
        void addHiddenRule('subject', actionButton.dataset.filterValue || '', actionButton.dataset.filterLabel || '')
      } else if (action === 'hide-address') {
        void addHiddenRule('address', actionButton.dataset.filterValue || '', actionButton.dataset.filterLabel || '')
      } else if (action === 'download-json') {
        void downloadCurrentMessage('json')
      } else if (action === 'download-eml') {
        void downloadCurrentMessage('eml')
      } else if (action === 'toggle-review-flag') {
        void toggleReviewFlag()
      } else if (action === 'clear-review') {
        void clearReview()
      } else if (action === 'remove-review-tag') {
        const tag = actionButton.dataset.tag || ''
        void removeReviewTag(tag)
      }
    })

    ui.messageDetail.addEventListener('submit', (event) => {
      const form = event.target.closest('[data-action="add-review-tag"]')
      if (!form) {
        return
      }
      event.preventDefault()
      const input = form.querySelector('[data-review-tag-input]')
      const value = input ? input.value : ''
      void addReviewTag(value)
    })

    ui.hiddenFiltersPanel.addEventListener('click', (event) => {
      const actionButton = event.target.closest('[data-action="remove-hidden-filter"]')
      if (!actionButton) {
        const closeButton = event.target.closest('[data-action="close-hidden-filters"]')
        if (closeButton) {
          closeHiddenFiltersDropdown()
        }
        return
      }
      void removeHiddenRule(actionButton.dataset.filterId || '')
    })

    ui.hiddenFiltersToggle.addEventListener('click', (event) => {
      event.preventDefault()
      toggleHiddenFiltersDropdown()
    })

    ui.searchInput.addEventListener(
      'input',
      debounce(() => {
        state.query = ui.searchInput.value
        saveState()
      }, 50)
    )

    ui.searchInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault()
        void executeSearch(1, { selectPreferred: true })
      }
    })

    document.addEventListener('click', (event) => {
      if (!state.hiddenFiltersOpen || !ui.hiddenFiltersDropdown) {
        return
      }
      const target = event.target
      if (target instanceof Node && ui.hiddenFiltersDropdown.contains(target)) {
        return
      }
      closeHiddenFiltersDropdown()
    })

    document.addEventListener('keydown', (event) => {
      if (event.key !== 'Escape') {
        return
      }
      closeHiddenFiltersDropdown()
    })

    ui.pstFilter.addEventListener('input', () => {
      state.mailboxFilter = ui.pstFilter.value
      renderPstCatalog()
    })

    ui.searchButton.addEventListener('click', () => {
      void executeSearch(1, { selectPreferred: true })
    })

    ui.mailOnlyToggle.addEventListener('change', () => {
      state.mailOnly = ui.mailOnlyToggle.checked
      saveState()
      if (state.sessionId && (state.currentFolderId || state.currentSearchPage)) {
        void loadVisibleMessages(1, { selectPreferred: true })
      }
    })

    ui.reviewFlaggedToggle.addEventListener('change', () => {
      state.reviewFlaggedOnly = ui.reviewFlaggedToggle.checked
      saveState()
      if (state.sessionId && (state.currentFolderId || state.currentSearchPage)) {
        void loadVisibleMessages(1, { selectPreferred: true })
      }
    })

    ui.reviewTaggedToggle.addEventListener('change', () => {
      state.reviewTaggedOnly = ui.reviewTaggedToggle.checked
      saveState()
      if (state.sessionId && (state.currentFolderId || state.currentSearchPage)) {
        void loadVisibleMessages(1, { selectPreferred: true })
      }
    })

    ui.sortSelect.addEventListener('change', () => {
      state.sort = ui.sortSelect.value
      saveState()
      if (state.sessionId && (state.currentFolderId || state.currentSearchPage)) {
        void loadVisibleMessages(1, { selectPreferred: true })
      }
    })

    ui.pagePrev.addEventListener('click', () => {
      void navigateFolderPage(-1)
    })

    ui.pageNext.addEventListener('click', () => {
      void navigateFolderPage(1)
    })

    ui.messagePrev.addEventListener('click', () => {
      void navigateMessage(-1)
    })

    ui.messageNext.addEventListener('click', () => {
      void navigateMessage(1)
    })

    window.addEventListener('beforeunload', saveState)
  }

  function debounce(fn, delay) {
    let timer = null
    return (...args) => {
      window.clearTimeout(timer)
      timer = window.setTimeout(() => fn(...args), delay)
    }
  }

  async function bootstrap() {
    ui.refreshCatalog = getElement('refresh-catalog')
    ui.sessionSummary = getElement('session-summary')
    ui.pstCountBadge = getElement('pst-count-badge')
    ui.pstEmpty = getElement('pst-empty')
    ui.pstList = getElement('pst-list')
    ui.folderTree = getElement('folder-tree')
    ui.folderCountBadge = getElement('folder-count-badge')
    ui.messageCountBadge = getElement('message-count-badge')
    ui.scopeCountBadge = getElement('scope-count-badge')
    ui.caseSelect = getElement('case-select')
    ui.scopeSelect = getElement('scope-select')
    ui.pstFilter = getElement('pst-filter')
    ui.searchInput = getElement('message-search')
    ui.searchButton = getElement('search-button')
    ui.searchScopeSelect = getElement('search-scope-select')
    ui.flaggedBundleScopeSelect = getElement('flagged-bundle-scope-select')
    ui.flaggedBundleButton = getElement('flagged-bundle-button')
    ui.hiddenFiltersDropdown = getElement('hidden-filters-dropdown')
    ui.hiddenFiltersToggle = getElement('hidden-filters-toggle')
    ui.hiddenFiltersCount = getElement('hidden-filters-count')
    ui.hiddenFiltersPanel = getElement('hidden-filters-panel')
    ui.mailOnlyToggle = getElement('mail-only-toggle')
    ui.reviewFlaggedToggle = getElement('review-flagged-toggle')
    ui.reviewTaggedToggle = getElement('review-tagged-toggle')
    ui.sortSelect = getElement('message-sort')
    ui.messageList = getElement('message-list')
    ui.messageDetail = getElement('message-detail')
    ui.messageResultCount = getElement('message-result-count')
    ui.pageInfo = getElement('page-info')
    ui.pagePrev = getElement('page-prev')
    ui.pageNext = getElement('page-next')
    ui.messagePrev = getElement('message-prev')
    ui.messageNext = getElement('message-next')
    ui.statusBar = getElement('status-bar')

    state.query = localStorage.getItem(STORAGE_KEYS.query) || ''
    state.searchScope = ['pst', 'search', 'all'].includes(
      localStorage.getItem(STORAGE_KEYS.searchScope) || ''
    )
      ? localStorage.getItem(STORAGE_KEYS.searchScope) || 'pst'
      : 'pst'
    state.mailboxScopeView =
      localStorage.getItem(STORAGE_KEYS.mailboxScopeView) === 'all' ? 'all' : 'search'
    state.selectedCasePath = normalizeScopePath(localStorage.getItem(STORAGE_KEYS.casePath) || '')
    state.selectedScopePath = normalizeScopePath(localStorage.getItem(STORAGE_KEYS.scopePath) || '')
    state.selectedScopeLabel = getScopeLabel(state.selectedScopePath)
    state.mailOnly = readStorageBool(STORAGE_KEYS.mailOnly, true)
    state.sort = localStorage.getItem(STORAGE_KEYS.sort) || 'date-desc'
    state.reviewFlaggedOnly = readStorageBool(STORAGE_KEYS.reviewFlaggedOnly, false)
    state.reviewTaggedOnly = readStorageBool(STORAGE_KEYS.reviewTaggedOnly, false)
    state.bundleScope = ['all', 'search', 'pst'].includes(
      localStorage.getItem(STORAGE_KEYS.flaggedBundleScope) || ''
    )
      ? localStorage.getItem(STORAGE_KEYS.flaggedBundleScope) || 'all'
      : 'all'
    state.hiddenFiltersOpen = false
    state.selectedPstFileName = localStorage.getItem(STORAGE_KEYS.pstFileName) || null
    state.mailboxFilter = ''
    refreshControls()
    wireEvents()
    renderPstCatalog()
    await loadMailboxCatalog({
      showBusy: true,
      casePath: state.selectedCasePath || undefined,
      scopePath: state.selectedScopePath || undefined,
      preferredFileName: state.selectedPstFileName || undefined
    })
    await loadHiddenFilters()
  }

  bootstrap().catch((error) => {
    console.error(error)
    setStatus(`Unable to initialize the viewer: ${error.message}`, 'error')
    ui.sessionSummary.innerHTML = `
      <div class="panel-empty">
        <strong>Initialization failed.</strong> ${escapeHtml(error.message)}
      </div>
    `
  })
})()
