(function () {
  const STORAGE_KEYS = {
    casePath: 'pst-mail-explorer.casePath',
    scopePath: 'pst-mail-explorer.scopePath',
    mailboxScopeView: 'pst-mail-explorer.mailboxScopeView',
    mailboxFilter: 'pst-mail-explorer.mailboxFilter',
    catalogMode: 'pst-mail-explorer.catalogMode',
    pstFileName: 'pst-mail-explorer.pstFileName',
    folderId: 'pst-mail-explorer.folderId',
    messageId: 'pst-mail-explorer.messageId',
    query: 'pst-mail-explorer.query',
    searchScope: 'pst-mail-explorer.searchScope',
    mailOnly: 'pst-mail-explorer.mailOnly',
    sort: 'pst-mail-explorer.sort',
    reviewFlaggedOnly: 'pst-mail-explorer.reviewFlaggedOnly',
    reviewTaggedOnly: 'pst-mail-explorer.reviewTaggedOnly',
    flaggedBundleScope: 'pst-mail-explorer.flaggedBundleScope',
    theme: 'pst-mail-explorer.theme'
  }

  const state = {
    authEnabled: true,
    authenticated: false,
    authChecked: false,
    authLoading: false,
    authUser: '',
    authUsers: [],
    authUsersLoading: false,
    authUsersLoaded: false,
    authCanManageUsers: false,
    authMfaEnabled: false,
    authView: 'login',
    authInviteToken: '',
    authInviteLoading: false,
    authInviteLoaded: false,
    authInvite: null,
    authInviteError: '',
    authInviteMfaAvailable: false,
    authInviteStep: 'password',
    authInviteSetup: null,
    authInviteRecoveryCodes: [],
    authInviteUsername: '',
    authMfaChallengeUsername: '',
    authMfaChallengeError: '',
    authMfaReminderOpen: false,
    authMfaReminderUsername: '',
    authMfaReminderDeferred: false,
    authMfaSetupOpen: false,
    authMfaSetupLoaded: false,
    authMfaSetupStep: 'idle',
    authMfaSetupUsername: '',
    authMfaSetupData: null,
    authMfaSetupRecoveryCodes: [],
    authMfaSetupError: '',
    selectedUserActivityUsername: '',
    userActivityLoading: false,
    userActivityLoaded: false,
    userActivityEntries: [],
    smtpSettingsOpen: false,
    smtpSettingsLoading: false,
    smtpSettingsSaving: false,
    smtpSettingsTesting: false,
    smtpSettingsLoaded: false,
    smtpSettingsHasPassword: false,
    activityLogOpen: false,
    activityLogLoading: false,
    activityLogLoaded: false,
    activityLogEntries: [],
    settingsMenuOpen: false,
    userManagementOpen: false,
    theme: 'light',
    sessionId: null,
    catalogScopes: [],
    catalog: [],
    catalogLoaded: false,
    catalogMessage: '',
    catalogMode: 'active',
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
    viewerInitialized: false,
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

  function renderDownloadIcon() {
    return `
      <svg class="button-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M12 4.5v8.25" />
        <path d="m8.25 9 3.75 3.75L15.75 9" />
        <path d="M5 17.25h14" />
      </svg>
    `
  }

  function renderTrashIcon() {
    return `
      <svg class="button-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M4.75 7h14.5" />
        <path d="M9 7V5.75A1.75 1.75 0 0 1 10.75 4h2.5A1.75 1.75 0 0 1 15 5.75V7" />
        <path d="m8.5 7 .65 10.95A1.75 1.75 0 0 0 10.9 19.5h2.2a1.75 1.75 0 0 0 1.75-1.55L15.5 7" />
        <path d="M10.5 10v5m3-5v5" />
      </svg>
    `
  }

  function hasText(value) {
    return Boolean(String(value ?? '').trim())
  }

  function normalizeAuthUserKey(value) {
    return String(value || '')
      .trim()
      .toLowerCase()
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

  function getCatalogModeLabel(mode) {
    return mode === 'removed' ? 'Removed PSTs' : 'Mailboxes'
  }

  function getCatalogEndpoint() {
    return state.catalogMode === 'removed' ? '/api/psts/removed' : '/api/psts'
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

  function buildHiddenRuleLookup(rules) {
    const addresses = new Set()
    const subjects = new Set()

    for (const rule of Array.isArray(rules) ? rules : []) {
      const normalizedValue = normalizeHiddenRuleValue(rule && rule.value)
      if (!normalizedValue) {
        continue
      }

      if (rule && rule.kind === 'address') {
        addresses.add(normalizedValue)
        for (const email of normalizedValue.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []) {
          addresses.add(normalizeHiddenRuleValue(email))
        }
      } else if (rule && rule.kind === 'subject') {
        subjects.add(normalizedValue)
      }
    }

    return { addresses, subjects }
  }

  function messageMatchesHiddenRules(item, hiddenLookup) {
    if (!hiddenLookup || (!hiddenLookup.addresses.size && !hiddenLookup.subjects.size)) {
      return false
    }

    if (hiddenLookup.addresses.size) {
      const addressValues = [
        item && item.senderEmailAddress,
        item && item.displayTo,
        item && item.displayCC,
        item && item.displayBCC,
        item && item.resolvedDisplayTo,
        item && item.resolvedDisplayCC,
        item && item.resolvedDisplayBCC,
        item && item.recipientText
      ]
        .map((value) => normalizeHiddenRuleValue(value))
        .filter(Boolean)

      for (const value of addressValues) {
        if (hiddenLookup.addresses.has(value)) {
          return true
        }
        for (const email of value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []) {
          if (hiddenLookup.addresses.has(normalizeHiddenRuleValue(email))) {
            return true
          }
        }
      }
    }

    if (hiddenLookup.subjects.size) {
      const subjectValues = [
        normalizeHiddenRuleValue(item && item.subject),
        normalizeHiddenRuleValue(item && item.originalSubject)
      ].filter(Boolean)
      if (subjectValues.some((value) => hiddenLookup.subjects.has(value))) {
        return true
      }
    }

    return false
  }

  function filterVisibleMessageItems(items) {
    const hiddenLookup = buildHiddenRuleLookup(state.hiddenRules)
    if (!hiddenLookup.addresses.size && !hiddenLookup.subjects.size) {
      return Array.isArray(items) ? items : []
    }
    return (Array.isArray(items) ? items : []).filter((item) => !messageMatchesHiddenRules(item, hiddenLookup))
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
    if (!activePage) {
      return
    }

    const preferredMessageId = state.selectedMessageId || null
    if (state.currentSearchPage && state.activeSearch && hasSearchCriteria(state.activeSearch)) {
      if (state.activeSearch.searchScope === 'pst' && !state.sessionId) {
        return
      }
      await loadSearchResults(activePage.page || 1, {
        selectPreferred: false,
        preferredMessageId
      })
      return
    }

    if (!state.sessionId) {
      return
    }
    await loadVisibleMessages(activePage.page || 1, {
      selectPreferred: false,
      preferredMessageId
    })
  }

  function clearMailboxSessionState(options = {}) {
    const preserveSearchResults = Boolean(options.preserveSearchResults)
    const suppressSave = Boolean(options.suppressSave)
    state.sessionId = null
    state.summary = null
    state.tree = null
    state.folderMap = new Map()
    state.currentFolderId = null
    state.currentFolderPage = null
    state.currentMessageDetail = null
    state.selectedMessageId = null
    state.selectedPstFileName = null
    if (!preserveSearchResults) {
      state.currentSearchPage = null
      state.activeSearch = null
    }
    renderSummary()
    renderFolderTree()
    renderMessageList()
    renderMessageDetail()
    if (!suppressSave) {
      saveState()
    }
  }

  async function refreshSearchIndex() {
    const response = await fetchJson('/api/search/index/refresh', {
      method: 'POST'
    })
    await loadHiddenFilters()
    return response
  }

  async function removeMailboxFromPlatform(fileName, scopePath) {
    const normalizedScopePath = normalizeScopePath(scopePath || state.selectedScopePath || '')
    const currentSessionId = state.sessionId
    const hasGlobalSearch =
      Boolean(state.currentSearchPage && state.activeSearch && hasSearchCriteria(state.activeSearch)) &&
      state.activeSearch.searchScope !== 'pst'
    const response = await fetchJson('/api/psts/remove', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: normalizedScopePath,
        fileName
      })
    })

    const removedCurrentSession =
      Boolean(currentSessionId) &&
      Array.isArray(response.closedSessionIds) &&
      response.closedSessionIds.includes(currentSessionId)

    if (removedCurrentSession) {
      if (hasGlobalSearch) {
        clearMailboxSessionState({ preserveSearchResults: true })
      } else {
        clearMailboxSessionState()
      }
    }

    await loadMailboxCatalog({
      showBusy: false,
      refreshOnly: true,
      casePath: state.selectedCasePath || undefined,
      scopePath: state.selectedScopePath || undefined,
      preferredFileName: state.selectedPstFileName || undefined,
      preferredScopePath: state.selectedScopePath || undefined
    })

    if (
      hasGlobalSearch &&
      state.activeSearch &&
      state.activeSearch.searchScope === 'search'
    ) {
      const activeSearchScope = normalizeScopePath(
        state.activeSearch.scopePath || state.selectedScopePath || ''
      )
      if (activeSearchScope && !getCatalogScope(activeSearchScope)) {
        state.currentSearchPage = null
        state.activeSearch = null
        state.selectedMessageId = null
        state.currentMessageDetail = null
        renderMessageList()
        renderMessageDetail()
        saveState()
      }
    }

    if (hasGlobalSearch && state.currentSearchPage && state.activeSearch) {
      await refreshActiveVisibleMessages()
    }
    setStatus(`Removed ${fileName} from the platform.`, 'success')
  }

  async function restoreMailboxToPlatform(fileName, scopePath) {
    const normalizedScopePath = normalizeScopePath(scopePath || '')
    await fetchJson('/api/psts/restore', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: normalizedScopePath,
        fileName
      })
    })

    await loadMailboxCatalog({
      showBusy: false,
      refreshOnly: true,
      casePath: state.selectedCasePath || undefined,
      scopePath: state.selectedScopePath || undefined,
      preferredFileName: state.selectedPstFileName || undefined,
      preferredScopePath: state.selectedScopePath || undefined
    })

    if (state.currentSearchPage || state.currentFolderPage) {
      await refreshActiveVisibleMessages()
    }
    setStatus(`Restored ${fileName} to the active catalog.`, 'success')
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

  function setAuthBusy(isBusy) {
    state.authLoading = isBusy
    if (ui.authSubmit) {
      ui.authSubmit.disabled = isBusy
      ui.authSubmit.textContent = isBusy ? 'Signing in...' : 'Sign in'
    }
    if (ui.authUsername) {
      ui.authUsername.disabled = isBusy
    }
    if (ui.authPassword) {
      ui.authPassword.disabled = isBusy
    }
    if (ui.authMfaSubmit) {
      ui.authMfaSubmit.disabled = isBusy
      ui.authMfaSubmit.textContent = isBusy ? 'Verifying...' : 'Verify code'
    }
    if (ui.authMfaCode) {
      ui.authMfaCode.disabled = isBusy
    }
    if (ui.inviteSubmit) {
      ui.inviteSubmit.disabled = isBusy
      ui.inviteSubmit.textContent = isBusy ? 'Saving...' : 'Set password'
    }
    if (ui.invitePassword) {
      ui.invitePassword.disabled = isBusy
    }
    if (ui.inviteConfirmPassword) {
      ui.inviteConfirmPassword.disabled = isBusy
    }
    if (ui.inviteMfaStart) {
      ui.inviteMfaStart.disabled = isBusy
    }
    if (ui.inviteMfaSkip) {
      ui.inviteMfaSkip.disabled = isBusy
    }
    if (ui.inviteMfaSubmit) {
      ui.inviteMfaSubmit.disabled = isBusy
      ui.inviteMfaSubmit.textContent = isBusy ? 'Verifying...' : 'Verify code'
    }
    if (ui.inviteMfaCode) {
      ui.inviteMfaCode.disabled = isBusy
    }
    if (ui.inviteMfaDownload) {
      ui.inviteMfaDownload.disabled = isBusy
    }
    if (ui.inviteFinish) {
      ui.inviteFinish.disabled = isBusy
    }
    if (ui.authLogout) {
      ui.authLogout.disabled = isBusy
    }
    if (ui.authScreen) {
      const authControls = ui.authScreen.querySelectorAll('input, button, textarea, select')
      authControls.forEach((control) => {
        if (control === ui.authLogout) {
          return
        }
        if (control === ui.authSubmit) {
          control.disabled = isBusy
          control.textContent = isBusy ? 'Signing in...' : 'Sign in'
          return
        }
        if (control === ui.authMfaSubmit) {
          control.disabled = isBusy
          control.textContent = isBusy ? 'Verifying...' : 'Verify code'
          return
        }
        if (control === ui.inviteSubmit) {
          control.disabled = isBusy
          control.textContent = isBusy ? 'Saving...' : 'Set password'
          return
        }
        if (control === ui.inviteMfaSubmit) {
          control.disabled = isBusy
          control.textContent = isBusy ? 'Verifying...' : 'Verify code'
          return
        }
        control.disabled = isBusy
      })
    }
  }

  function setAuthError(message = '') {
    if (!ui.authMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.authMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.authMessage.dataset.tone = 'error'
    } else {
      delete ui.authMessage.dataset.tone
    }
  }

  function setAuthMessage(message = '', tone = 'neutral') {
    if (!ui.authMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.authMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.authMessage.dataset.tone = tone
    } else {
      delete ui.authMessage.dataset.tone
    }
  }

  function getInviteTokenFromLocation() {
    const match = String(window.location.pathname || '').match(/^\/invite\/([^/?#]+)/i)
    return match ? decodeURIComponent(match[1] || '') : ''
  }

  function setAuthView(view) {
    state.authView = view
    if (ui.authLoginView) {
      ui.authLoginView.hidden = view !== 'login'
    }
    if (ui.authMfaView) {
      ui.authMfaView.hidden = view !== 'mfa'
    }
    if (ui.inviteView) {
      ui.inviteView.hidden = view !== 'invite'
    }
  }

  function resetInviteState() {
    state.authInviteLoading = false
    state.authInviteLoaded = false
    state.authInvite = null
    state.authInviteError = ''
    state.authInviteMfaAvailable = false
    state.authInviteStep = 'password'
    state.authInviteSetup = null
    state.authInviteRecoveryCodes = []
    state.authInviteUsername = ''
    if (ui.inviteDetails) {
      ui.inviteDetails.innerHTML = ''
    }
    if (ui.inviteMfaPrompt) {
      ui.inviteMfaPrompt.hidden = true
    }
    if (ui.inviteMfaSetup) {
      ui.inviteMfaSetup.hidden = true
    }
    if (ui.inviteMfaComplete) {
      ui.inviteMfaComplete.hidden = true
    }
    if (ui.inviteMfaSecret) {
      ui.inviteMfaSecret.textContent = ''
    }
    if (ui.inviteMfaUri) {
      ui.inviteMfaUri.textContent = ''
    }
    if (ui.inviteMfaQr) {
      ui.inviteMfaQr.removeAttribute('src')
    }
    if (ui.inviteMfaRecoveryList) {
      ui.inviteMfaRecoveryList.innerHTML = ''
    }
    if (ui.invitePassword) {
      ui.invitePassword.value = ''
    }
    if (ui.inviteConfirmPassword) {
      ui.inviteConfirmPassword.value = ''
    }
    if (ui.inviteMfaCode) {
      ui.inviteMfaCode.value = ''
    }
  }

  function renderInviteDetails(invite) {
    if (!ui.inviteDetails) {
      return
    }

    if (!invite || !invite.username) {
      ui.inviteDetails.innerHTML = ''
      return
    }

    const status = String(invite.inviteStatus || 'pending')
    const chips = [
      `<span class="chip accent">${escapeHtml(status === 'pending' ? 'Pending invite' : status)}</span>`,
      invite.mfaEnabled ? '<span class="chip green">MFA enabled</span>' : ''
    ].filter(Boolean)

    ui.inviteDetails.innerHTML = `
      <div class="invite-details-card">
        <div class="invite-details-row">
          <strong>${escapeHtml(invite.username)}</strong>
          <span class="invite-details-email">${escapeHtml(invite.recipientEmail || '')}</span>
        </div>
        <div class="invite-details-meta">
          ${chips.join('')}
        </div>
        <div class="invite-details-expiry">
          Expires ${escapeHtml(invite.inviteExpiresAt || 'soon')}
        </div>
      </div>
    `
  }

  function renderInviteRecoveryCodes(codes) {
    if (!ui.inviteMfaRecoveryList) {
      return
    }

    const normalizedCodes = Array.isArray(codes)
      ? codes.map((code) => String(code || '').trim()).filter(Boolean)
      : []
    state.authInviteRecoveryCodes = normalizedCodes

    if (!normalizedCodes.length) {
      ui.inviteMfaRecoveryList.innerHTML = '<div class="panel-empty">No recovery codes available.</div>'
      return
    }

    ui.inviteMfaRecoveryList.innerHTML = `
      <div class="recovery-code-grid">
        ${normalizedCodes
          .map((code) => `<code class="recovery-code">${escapeHtml(code)}</code>`)
          .join('')}
      </div>
    `
  }

  function downloadTextFile(filename, content) {
    const blob = new Blob([String(content || '')], { type: 'text/plain;charset=utf-8' })
    const url = window.URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = filename
    anchor.rel = 'noopener'
    document.body.appendChild(anchor)
    anchor.click()
    anchor.remove()
    window.URL.revokeObjectURL(url)
  }

  function downloadInviteRecoveryCodes() {
    if (!Array.isArray(state.authInviteRecoveryCodes) || !state.authInviteRecoveryCodes.length) {
      return
    }

    downloadTextFile(
      'mfa-recovery-codes.txt',
      state.authInviteRecoveryCodes.map((code) => code.trim()).filter(Boolean).join('\n')
    )
  }

  function getMfaReminderDismissalKey(username) {
    const normalizedUsername = normalizeAuthUserKey(username)
    if (!normalizedUsername) {
      return ''
    }

    return `pst-mail-explorer.mfa-reminder.dismissed::${normalizedUsername}`
  }

  function hasDismissedMfaReminder(username) {
    const key = getMfaReminderDismissalKey(username)
    if (!key) {
      return false
    }

    try {
      return sessionStorage.getItem(key) === '1'
    } catch (error) {
      return false
    }
  }

  function dismissMfaReminder(username) {
    const key = getMfaReminderDismissalKey(username)
    if (!key) {
      return
    }

    try {
      sessionStorage.setItem(key, '1')
    } catch (error) {
      // Ignore storage errors in restricted browser contexts.
    }
  }

  function clearMfaReminderDismissal(username) {
    const key = getMfaReminderDismissalKey(username)
    if (!key) {
      return
    }

    try {
      sessionStorage.removeItem(key)
    } catch (error) {
      // Ignore storage errors in restricted browser contexts.
    }
  }

  function setMfaSetupMessage(message = '', tone = 'neutral') {
    if (!ui.mfaSetupMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.mfaSetupMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.mfaSetupMessage.dataset.tone = tone
    } else {
      delete ui.mfaSetupMessage.dataset.tone
    }
  }

  function setMfaSetupBusy(isBusy, label = 'Verify code') {
    if (ui.mfaSetupSubmit) {
      ui.mfaSetupSubmit.disabled = Boolean(isBusy)
      ui.mfaSetupSubmit.textContent =
        label === 'Loading...'
          ? 'Loading...'
          : label === 'Sending...'
            ? 'Sending...'
            : 'Verify code'
    }
    if (ui.mfaSetupDownload) {
      ui.mfaSetupDownload.disabled = Boolean(isBusy) || !state.authMfaSetupRecoveryCodes.length
    }
    if (ui.mfaSetupFinish) {
      ui.mfaSetupFinish.disabled = Boolean(isBusy)
    }
    if (ui.mfaSetupForm) {
      const fields = ui.mfaSetupForm.querySelectorAll('input, select, textarea, button')
      fields.forEach((field) => {
        if (field === ui.mfaSetupSubmit || field === ui.mfaSetupDownload || field === ui.mfaSetupFinish || field === ui.mfaSetupClose) {
          return
        }
        field.disabled = Boolean(isBusy)
      })
    }
  }

  function resetMfaSetupState() {
    state.authMfaSetupOpen = false
    state.authMfaSetupLoaded = false
    state.authMfaSetupStep = 'idle'
    state.authMfaSetupUsername = ''
    state.authMfaSetupData = null
    state.authMfaSetupRecoveryCodes = []
    state.authMfaSetupError = ''
    setMfaSetupMessage('')
    if (ui.mfaSetupQr) {
      ui.mfaSetupQr.removeAttribute('src')
    }
    if (ui.mfaSetupSecret) {
      ui.mfaSetupSecret.textContent = ''
    }
    if (ui.mfaSetupUri) {
      ui.mfaSetupUri.textContent = ''
    }
    if (ui.mfaSetupRecoveryList) {
      ui.mfaSetupRecoveryList.innerHTML = ''
    }
    if (ui.mfaSetupCode) {
      ui.mfaSetupCode.value = ''
    }
    if (ui.mfaSetupForm) {
      ui.mfaSetupForm.reset()
    }
    if (ui.mfaSetupSubmit) {
      ui.mfaSetupSubmit.disabled = false
    }
    if (ui.mfaSetupDownload) {
      ui.mfaSetupDownload.disabled = true
    }
    if (ui.mfaSetupFinish) {
      ui.mfaSetupFinish.disabled = true
    }
  }

  function renderMfaSetupRecoveryCodes(codes) {
    if (!ui.mfaSetupRecoveryList) {
      return
    }

    const normalizedCodes = Array.isArray(codes)
      ? codes.map((code) => String(code || '').trim()).filter(Boolean)
      : []
    state.authMfaSetupRecoveryCodes = normalizedCodes

    if (!normalizedCodes.length) {
      ui.mfaSetupRecoveryList.innerHTML = '<div class="panel-empty">No recovery codes available.</div>'
      return
    }

    ui.mfaSetupRecoveryList.innerHTML = `
      <div class="recovery-code-grid">
        ${normalizedCodes
          .map((code) => `<code class="recovery-code">${escapeHtml(code)}</code>`)
          .join('')}
      </div>
    `
  }

  function downloadMfaSetupRecoveryCodes() {
    if (!Array.isArray(state.authMfaSetupRecoveryCodes) || !state.authMfaSetupRecoveryCodes.length) {
      return
    }

    downloadTextFile(
      'mfa-recovery-codes.txt',
      state.authMfaSetupRecoveryCodes.map((code) => code.trim()).filter(Boolean).join('\n')
    )
  }

  function closeMfaReminderModal(preserveUsername = false) {
    state.authMfaReminderOpen = false
    if (!preserveUsername) {
      state.authMfaReminderUsername = ''
      state.authMfaReminderDeferred = false
    }
    if (ui.mfaReminderModal) {
      ui.mfaReminderModal.hidden = true
    }
  }

  function openMfaReminderModal(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return false
    }

    state.authMfaReminderOpen = true
    state.authMfaReminderDeferred = false
    state.authMfaReminderUsername = normalizedUsername
    if (ui.mfaReminderModal) {
      ui.mfaReminderModal.hidden = false
    }
    if (ui.mfaReminderSetup) {
      window.requestAnimationFrame(() => {
        ui.mfaReminderSetup.focus()
      })
    }
    return true
  }

  function closeMfaSetupModal() {
    const shouldRestoreReminder =
      Boolean(state.authMfaReminderDeferred) &&
      !state.authMfaEnabled &&
      !state.viewerInitialized &&
      Boolean(state.authMfaReminderUsername)
    state.authMfaSetupOpen = false
    if (ui.mfaSetupModal) {
      ui.mfaSetupModal.hidden = true
    }
    setMfaSetupBusy(false)
    setMfaSetupMessage('')
    closeSettingsMenu()
    if (shouldRestoreReminder) {
      openMfaReminderModal(state.authMfaReminderUsername)
    }
  }

  function showMfaSetupEnrollment(data) {
    state.authMfaSetupStep = 'setup'
    state.authMfaSetupData = data || null
    state.authMfaSetupLoaded = true
    if (ui.mfaSetupQr && data && data.qrCodeDataUrl) {
      ui.mfaSetupQr.src = data.qrCodeDataUrl
    }
    if (ui.mfaSetupSecret) {
      ui.mfaSetupSecret.textContent = data && data.secret ? data.secret : ''
    }
    if (ui.mfaSetupUri) {
      ui.mfaSetupUri.textContent = data && data.otpauthUri ? data.otpauthUri : ''
    }
    if (ui.mfaSetupCode) {
      ui.mfaSetupCode.value = ''
      window.requestAnimationFrame(() => {
        ui.mfaSetupCode.focus()
      })
    }
    if (ui.mfaSetupComplete) {
      ui.mfaSetupComplete.hidden = true
    }
    if (ui.mfaSetupForm) {
      ui.mfaSetupForm.hidden = false
    }
    if (ui.mfaSetupFinish) {
      ui.mfaSetupFinish.disabled = true
    }
    setMfaSetupMessage('Scan the QR code or enter the manual setup key.', 'success')
    if (ui.mfaSetupDownload) {
      ui.mfaSetupDownload.disabled = true
    }
  }

  function showMfaSetupComplete(codes) {
    state.authMfaSetupStep = 'complete'
    if (ui.mfaSetupForm) {
      ui.mfaSetupForm.hidden = true
    }
    if (ui.mfaSetupComplete) {
      ui.mfaSetupComplete.hidden = false
    }
    renderMfaSetupRecoveryCodes(codes)
    setMfaSetupMessage('MFA is now enabled.', 'success')
    if (ui.mfaSetupDownload) {
      ui.mfaSetupDownload.disabled = !state.authMfaSetupRecoveryCodes.length
    }
    if (ui.mfaSetupFinish) {
      ui.mfaSetupFinish.disabled = false
    }
  }

  async function startSelfServiceMfaSetup(options = {}) {
    const suppressReminder = Boolean(options.suppressReminder)
    if (!state.authenticated) {
      setStatus('Sign in first.', 'error')
      return false
    }

    if (!state.authMfaEnabled && !suppressReminder && state.authUser) {
      dismissMfaReminder(state.authUser)
    }

    if (state.authMfaReminderUsername) {
      state.authMfaReminderDeferred = true
      closeMfaReminderModal(true)
    } else {
      closeMfaReminderModal()
    }
    closeSettingsMenu()
    if (!ui.mfaSetupModal) {
      return false
    }

    resetMfaSetupState()
    state.authMfaSetupUsername = state.authUser
    ui.mfaSetupModal.hidden = false
    state.authMfaSetupOpen = true
    setMfaSetupMessage('Loading MFA setup...')
    setMfaSetupBusy(true, 'Loading...')
    try {
      const response = await fetchJson('/api/auth/mfa/enrollment/start', {
        method: 'POST'
      })
      showMfaSetupEnrollment(response)
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      setMfaSetupMessage(`Unable to start MFA setup: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      setMfaSetupBusy(false)
    }
  }

  async function completeSelfServiceMfaSetup() {
    if (!state.authMfaSetupOpen) {
      return false
    }

    const code = String(ui.mfaSetupCode?.value || '').trim()
    if (!code) {
      setMfaSetupMessage('Enter a verification code.', 'error')
      return false
    }

    setMfaSetupBusy(true, 'Sending...')
    try {
      const response = await fetchJson('/api/auth/mfa/enrollment/complete', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ code })
      })

      state.authMfaEnabled = Boolean(response.user && response.user.mfaEnabled)
      state.authMfaReminderDeferred = false
      if (response.user && response.user.username) {
        state.authUser = response.user.username
      }
      if (Array.isArray(response.recoveryCodes)) {
        state.authMfaSetupRecoveryCodes = response.recoveryCodes
      }
      showMfaSetupComplete(response.recoveryCodes || [])
      updateUserManagementVisibility()
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      setMfaSetupMessage(`Unable to finish MFA setup: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      setMfaSetupBusy(false)
    }
  }

  async function continueAfterMfaSetup() {
    state.authMfaReminderDeferred = false
    closeMfaSetupModal()
    if (!state.viewerInitialized) {
      try {
        await initializeViewer()
      } catch (error) {
        if (!(error && typeof error === 'object' && Number(error.statusCode) === 401)) {
          setStatus(`Signed in, but unable to load the viewer: ${getErrorMessage(error)}`, 'error')
        }
      }
    }
  }

  function showLoginScreen() {
    state.authView = 'login'
    state.authMfaChallengeUsername = ''
    setAuthView('login')
    setAuthMessage('')
    if (ui.authMfaCode) {
      ui.authMfaCode.value = ''
    }
    if (ui.authPassword) {
      ui.authPassword.value = ''
    }
    if (ui.authUsername) {
      window.requestAnimationFrame(() => {
        ui.authUsername.focus()
        if (typeof ui.authUsername.select === 'function') {
          ui.authUsername.select()
        }
      })
    }
  }

  function showMfaChallengeScreen(username) {
    state.authView = 'mfa'
    state.authMfaChallengeUsername = String(username || '').trim()
    setAuthView('mfa')
    setAuthMessage('Verification required. Enter a TOTP or recovery code.')
    if (ui.authMfaDescription) {
      ui.authMfaDescription.textContent = state.authMfaChallengeUsername
        ? `Use an authenticator code or a recovery code to sign in as ${state.authMfaChallengeUsername}.`
        : 'Use an authenticator code or a recovery code to finish signing in.'
    }
    if (ui.authMfaCode) {
      ui.authMfaCode.value = ''
      window.requestAnimationFrame(() => {
        ui.authMfaCode.focus()
      })
    }
  }

  function showInviteScreen(invite) {
    state.authView = 'invite'
    state.authMfaChallengeUsername = ''
    state.authInvite = invite || null
    state.authInviteLoaded = Boolean(invite)
    state.authInviteUsername = String(invite && invite.username ? invite.username : '').trim()
    setAuthView('invite')
    setAuthMessage('')
    renderInviteDetails(invite)
    if (ui.inviteSummary) {
      ui.inviteSummary.textContent = invite && invite.username
        ? `You have been invited as ${invite.username}. Choose a password to continue.`
        : 'Set your password to continue.'
    }
    if (ui.inviteMfaPrompt) {
      ui.inviteMfaPrompt.hidden = true
    }
    if (ui.inviteMfaSetup) {
      ui.inviteMfaSetup.hidden = true
    }
    if (ui.inviteMfaComplete) {
      ui.inviteMfaComplete.hidden = true
    }
    if (ui.invitePassword) {
      ui.invitePassword.value = ''
    }
    if (ui.inviteConfirmPassword) {
      ui.inviteConfirmPassword.value = ''
    }
    if (ui.invitePassword) {
      window.requestAnimationFrame(() => {
        ui.invitePassword.focus()
      })
    }
  }

  function showInviteMfaPrompt() {
    state.authInviteStep = 'prompt'
    if (ui.inviteMfaPrompt) {
      ui.inviteMfaPrompt.hidden = false
    }
    if (ui.inviteMfaSetup) {
      ui.inviteMfaSetup.hidden = true
    }
    if (ui.inviteMfaComplete) {
      ui.inviteMfaComplete.hidden = true
    }
  }

  function showInviteMfaSetup(data) {
    state.authInviteStep = 'setup'
    state.authInviteSetup = data || null
    if (ui.inviteMfaPrompt) {
      ui.inviteMfaPrompt.hidden = true
    }
    if (ui.inviteMfaSetup) {
      ui.inviteMfaSetup.hidden = false
    }
    if (ui.inviteMfaComplete) {
      ui.inviteMfaComplete.hidden = true
    }
    if (ui.inviteMfaQr && data && data.qrCodeDataUrl) {
      ui.inviteMfaQr.src = data.qrCodeDataUrl
    }
    if (ui.inviteMfaSecret) {
      ui.inviteMfaSecret.textContent = data && data.secret ? data.secret : ''
    }
    if (ui.inviteMfaUri) {
      ui.inviteMfaUri.textContent = data && data.otpauthUri ? data.otpauthUri : ''
    }
    if (ui.inviteMfaCode) {
      ui.inviteMfaCode.value = ''
      window.requestAnimationFrame(() => {
        ui.inviteMfaCode.focus()
      })
    }
  }

  function showInviteMfaComplete(codes) {
    state.authInviteStep = 'complete'
    if (ui.inviteMfaPrompt) {
      ui.inviteMfaPrompt.hidden = true
    }
    if (ui.inviteMfaSetup) {
      ui.inviteMfaSetup.hidden = true
    }
    if (ui.inviteMfaComplete) {
      ui.inviteMfaComplete.hidden = false
    }
    renderInviteRecoveryCodes(codes)
  }

  async function submitInvitePassword() {
    if (!state.authInviteToken) {
      setAuthError('Invite token is missing.')
      return false
    }

    const password = String(ui.invitePassword?.value || '')
    const confirmPassword = String(ui.inviteConfirmPassword?.value || '')
    if (!password || !confirmPassword) {
      setAuthError('Enter and confirm a password.')
      return false
    }
    if (password !== confirmPassword) {
      setAuthError('Passwords do not match.')
      return false
    }

    setAuthBusy(true)
    try {
      const response = await fetchJson(`/api/auth/invites/${encodeURIComponent(state.authInviteToken)}/accept`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          password,
          confirmPassword
        })
      })

      const username = (response.user && response.user.username) || state.authInviteUsername || ''
      if (!username) {
        throw new Error('Unable to complete invite.')
      }

      state.authInviteUsername = username
      state.authInviteMfaAvailable = Boolean(response.mfaAvailable)
      if (ui.invitePassword) {
        ui.invitePassword.value = ''
      }
      if (ui.inviteConfirmPassword) {
        ui.inviteConfirmPassword.value = ''
      }
      try {
        window.history.replaceState({}, '', '/')
      } catch (error) {
        // Ignore history updates in restrictive browser contexts.
      }
      setAuthMessage('Password saved.', 'success')

      if (response.mfaAvailable) {
        showInviteMfaPrompt()
      } else {
        await finalizeInviteOnboarding(username, false, true, {
          suppressReminder: true
        })
      }

      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 400) {
        setAuthError(getErrorMessage(error))
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 404) {
        setAuthError('Invite not found.')
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 410) {
        setAuthError(getErrorMessage(error))
        return false
      }
      setAuthError(`Unable to accept invite: ${getErrorMessage(error)}`)
      return false
    } finally {
      setAuthBusy(false)
    }
  }

  async function startInviteMfaSetup() {
    if (!state.authInviteUsername) {
      setAuthError('Set your password first.')
      return
    }

    setAuthBusy(true)
    try {
      const response = await fetchJson('/api/auth/mfa/enrollment/start', {
        method: 'POST'
      })
      state.authInviteSetup = response
      showInviteMfaSetup(response)
      setAuthMessage('Scan the QR code or enter the manual setup key.', 'success')
    } catch (error) {
      setAuthError(`Unable to start MFA setup: ${getErrorMessage(error)}`)
    } finally {
      setAuthBusy(false)
    }
  }

  async function completeInviteMfaSetup() {
    if (!state.authInviteUsername) {
      setAuthError('Set your password first.')
      return
    }

    const code = String(ui.inviteMfaCode?.value || '').trim()
    if (!code) {
      setAuthError('Enter a verification code.')
      return
    }

    setAuthBusy(true)
    try {
      const response = await fetchJson('/api/auth/mfa/enrollment/complete', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ code })
      })
      state.authMfaEnabled = Boolean(response.user && response.user.mfaEnabled)
      state.authInviteRecoveryCodes = Array.isArray(response.recoveryCodes) ? response.recoveryCodes : []
      showInviteMfaComplete(state.authInviteRecoveryCodes)
      setAuthMessage('MFA is now enabled.', 'success')
    } catch (error) {
      setAuthError(`Unable to finish MFA setup: ${getErrorMessage(error)}`)
    } finally {
      setAuthBusy(false)
    }
  }

  async function beginAuthenticatedWorkspace(username, canManageUsers = false, mfaEnabled = false, options = {}) {
    const normalizedUsername = String(username || '').trim()
    try {
      window.history.replaceState({}, '', '/')
    } catch (error) {
      // Ignore history updates in restrictive browser contexts.
    }

    showViewerShell(normalizedUsername, canManageUsers, mfaEnabled)
    applyTheme(readThemePreference())
    if (!state.authenticated) {
      return false
    }

    const suppressReminder = Boolean(options && options.suppressReminder)
    const shouldPauseForReminder =
      Boolean(state.authEnabled) &&
      Boolean(normalizedUsername) &&
      !state.authMfaEnabled &&
      !suppressReminder &&
      !hasDismissedMfaReminder(normalizedUsername)

    if (shouldPauseForReminder) {
      openMfaReminderModal(normalizedUsername)
      return false
    }

    try {
      await initializeViewer()
      return true
    } catch (error) {
      if (!(error && typeof error === 'object' && Number(error.statusCode) === 401)) {
        setStatus(`Signed in, but unable to load the viewer: ${getErrorMessage(error)}`, 'error')
      }
      return false
    }
  }

  async function finalizeInviteOnboarding(username, canManageUsers = false, mfaEnabled = false, options = {}) {
    return beginAuthenticatedWorkspace(username, canManageUsers, mfaEnabled, options)
  }

  function resetUserManagementState() {
    state.selectedUserActivityUsername = ''
    state.userActivityEntries = []
    state.userActivityLoaded = false
    setUserActivityBusy(false)
    setUserActivityMessage('')
    updateUserActivitySelectionLabel()
    renderUserActivityEntries([])
  }

  function showAuthScreen() {
    const previousUser = String(state.authUser || '').trim()
    if (previousUser) {
      clearMfaReminderDismissal(previousUser)
    }
    state.authenticated = false
    state.authUser = ''
    state.authCanManageUsers = false
    state.authMfaEnabled = false
    state.authUsers = []
    state.authUsersLoaded = false
    state.viewerInitialized = false
    state.authMfaChallengeUsername = ''
    state.authMfaChallengeError = ''
    state.authMfaReminderOpen = false
    state.authMfaReminderUsername = ''
    state.authMfaReminderDeferred = false
    state.authMfaSetupOpen = false
    state.authMfaSetupLoaded = false
    state.authMfaSetupStep = 'idle'
    state.authMfaSetupUsername = ''
    state.authMfaSetupData = null
    state.authMfaSetupRecoveryCodes = []
    state.authMfaSetupError = ''
    resetUserManagementState()
    resetInviteState()
    state.activityLogEntries = []
    state.activityLogLoaded = false
    clearMailboxSessionState({ suppressSave: true })
    applyTheme('dark')
    closeMfaReminderModal()
    closeMfaSetupModal()
    closeSettingsMenu()
    closeUserManagementModal()
    closeSmtpSettingsModal()
    closeActivityLogModal()
    updateUserManagementVisibility()
    setAuthBusy(false)
    setUserManagementBusy(false)
    setUserManagementMessage('')
    setActivityLogBusy(false)
    setActivityLogMessage('')
    setBodyBusy(false)
    setAuthView('login')
    if (ui.appShell) {
      ui.appShell.hidden = true
    }
    if (ui.authScreen) {
      ui.authScreen.hidden = false
    }
    showLoginScreen()
  }

  function showViewerShell(username, canManageUsers = false, mfaEnabled = false) {
    state.authenticated = true
    state.authUser = String(username || state.authUser || '').trim()
    state.authCanManageUsers = Boolean(canManageUsers)
    state.authMfaEnabled = Boolean(mfaEnabled)
    state.authMfaReminderOpen = false
    state.authMfaReminderUsername = ''
    state.authMfaReminderDeferred = false
    state.authMfaSetupOpen = false
    state.authMfaSetupLoaded = false
    state.authMfaSetupStep = 'idle'
    state.authMfaSetupUsername = ''
    state.authMfaSetupData = null
    state.authMfaSetupRecoveryCodes = []
    state.authMfaSetupError = ''
    state.hiddenFiltersOpen = false
    clearMailboxSessionState({ suppressSave: true })
    resetUserManagementState()
    resetInviteState()
    state.activityLogEntries = []
    state.activityLogLoaded = false
    if (ui.authScreen) {
      ui.authScreen.hidden = true
    }
    if (ui.appShell) {
      ui.appShell.hidden = false
    }
    closeMfaReminderModal()
    closeMfaSetupModal()
    if (ui.authUser) {
      ui.authUser.textContent = state.authUser || 'Signed in'
    }
    closeSettingsMenu()
    closeUserManagementModal()
    closeSmtpSettingsModal()
    closeActivityLogModal()
    updateUserManagementVisibility()
    loadWorkspaceState()
    setAuthView('login')
    setAuthError('')
    setAuthBusy(false)
    setUserManagementBusy(false)
    setUserManagementMessage('')
    setActivityLogBusy(false)
    setActivityLogMessage('')
    setBodyBusy(false)
  }

  function handleAuthFailure(message = 'Session expired. Sign in again.') {
    showAuthScreen()
    setStatus(message, 'error')
  }

  function updateUserManagementVisibility() {
    const canUseSettings = Boolean(state.authEnabled) && Boolean(state.authenticated)
    const canManageUsers = canUseSettings && Boolean(state.authCanManageUsers)
    const canSetUpMfa = canUseSettings && !state.authMfaEnabled

    if (ui.settingsButton) {
      ui.settingsButton.hidden = !canUseSettings
      ui.settingsButton.setAttribute('aria-expanded', 'false')
    }
    if (ui.setupMfaButton) {
      ui.setupMfaButton.hidden = !canSetUpMfa
    }
    if (ui.activityLogButton) {
      ui.activityLogButton.hidden = !canManageUsers
    }
    if (ui.smtpSettingsButton) {
      ui.smtpSettingsButton.hidden = !canManageUsers
    }
  }

  function closeSettingsMenu() {
    state.settingsMenuOpen = false
    if (ui.settingsMenu) {
      ui.settingsMenu.hidden = true
    }
    if (ui.settingsButton) {
      ui.settingsButton.setAttribute('aria-expanded', 'false')
    }
  }

  function openSettingsMenu() {
    if (!state.authEnabled || !state.authenticated) {
      return
    }
    if (!ui.settingsMenu || !ui.settingsButton) {
      return
    }

    ui.settingsMenu.hidden = false
    state.settingsMenuOpen = true
    ui.settingsButton.setAttribute('aria-expanded', 'true')
  }

  function toggleSettingsMenu() {
    if (state.settingsMenuOpen) {
      closeSettingsMenu()
      return
    }

    openSettingsMenu()
  }

  function focusUserManagementUsername() {
    if (!ui.userManagementUsername) {
      return
    }

    window.requestAnimationFrame(() => {
      ui.userManagementUsername.focus()
      if (typeof ui.userManagementUsername.select === 'function') {
        ui.userManagementUsername.select()
      }
    })
  }

  function closeUserManagementModal() {
    state.userManagementOpen = false
    if (ui.userManagementModal) {
      ui.userManagementModal.hidden = true
    }
    setUserManagementBusy(false)
    setUserManagementMessage('')
    closeSettingsMenu()
  }

  function renderPaperPlaneIcon() {
    return `
      <svg class="button-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M4.5 11.5 19.5 4.5 14.75 19.5l-3.35-6.1-6.9-1.9Z" />
        <path d="m14.75 19.5-4.45-4.1" />
      </svg>
    `
  }

  function renderRotateIcon() {
    return `
      <svg class="button-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="M7.5 7.5H4.75V4.75" />
        <path d="M4.75 7.5A8.25 8.25 0 1 1 12 20.25a8.25 8.25 0 0 1-7.25-4.25" />
      </svg>
    `
  }

  function renderKeyIcon() {
    return `
      <svg class="button-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <circle cx="9" cy="9" r="3.75" />
        <path d="M12 12l6 6" />
        <path d="M16.5 15.5H19v2.5h-2.5v2.5H14" />
      </svg>
    `
  }

  function setSmtpSettingsMessage(message = '', tone = 'neutral') {
    if (!ui.smtpSettingsMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.smtpSettingsMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.smtpSettingsMessage.dataset.tone = tone
    } else {
      delete ui.smtpSettingsMessage.dataset.tone
    }
  }

  function setSmtpSettingsBusy(isBusy, label = 'Save settings') {
    state.smtpSettingsLoading = Boolean(isBusy) && label === 'Loading...'
    state.smtpSettingsSaving = Boolean(isBusy) && label === 'Saving...'
    state.smtpSettingsTesting = Boolean(isBusy) && label === 'Sending...'

    const isBusyState = Boolean(isBusy)
    if (ui.smtpSettingsSubmit) {
      ui.smtpSettingsSubmit.disabled = isBusyState
      ui.smtpSettingsSubmit.textContent = label === 'Saving...' ? 'Saving...' : 'Save settings'
    }
    if (ui.smtpTestSend) {
      ui.smtpTestSend.disabled = isBusyState
      ui.smtpTestSend.textContent = label === 'Sending...' ? 'Sending...' : 'Send test email'
    }

    if (ui.smtpSettingsForm) {
      const fields = ui.smtpSettingsForm.querySelectorAll('input, select, textarea, button')
      fields.forEach((field) => {
        if (field === ui.smtpSettingsSubmit || field === ui.smtpTestSend || field === ui.smtpSettingsClose) {
          return
        }
        field.disabled = isBusyState
      })
    }
  }

  function renderSmtpSettings(settings) {
    const normalized = {
      enabled: Boolean(settings && settings.enabled),
      host: String(settings && settings.host ? settings.host : ''),
      port: Number(settings && settings.port ? settings.port : 587) || 587,
      secure: Boolean(settings && settings.secure),
      username: String(settings && settings.username ? settings.username : ''),
      hasPassword: Boolean(settings && settings.hasPassword),
      fromName: String(settings && settings.fromName ? settings.fromName : ''),
      fromAddress: String(settings && settings.fromAddress ? settings.fromAddress : ''),
      replyTo: String(settings && settings.replyTo ? settings.replyTo : '')
    }

    state.smtpSettingsLoaded = true
    state.smtpSettingsHasPassword = normalized.hasPassword

    if (ui.smtpSettingsEnabled) {
      ui.smtpSettingsEnabled.checked = normalized.enabled
    }
    if (ui.smtpSettingsHost) {
      ui.smtpSettingsHost.value = normalized.host
    }
    if (ui.smtpSettingsPort) {
      ui.smtpSettingsPort.value = String(normalized.port || 587)
    }
    if (ui.smtpSettingsSecure) {
      ui.smtpSettingsSecure.checked = normalized.secure
    }
    if (ui.smtpSettingsUsername) {
      ui.smtpSettingsUsername.value = normalized.username
    }
    if (ui.smtpSettingsPassword) {
      ui.smtpSettingsPassword.value = ''
      ui.smtpSettingsPassword.placeholder = normalized.hasPassword
        ? 'Leave blank to keep the current password'
        : 'Enter the SMTP password'
    }
    if (ui.smtpPasswordHint) {
      ui.smtpPasswordHint.textContent = normalized.hasPassword
        ? 'Leave blank to keep the current password.'
        : 'Enter the SMTP password before saving.'
    }
    if (ui.smtpSettingsFromName) {
      ui.smtpSettingsFromName.value = normalized.fromName
    }
    if (ui.smtpSettingsFromAddress) {
      ui.smtpSettingsFromAddress.value = normalized.fromAddress
    }
    if (ui.smtpSettingsReplyTo) {
      ui.smtpSettingsReplyTo.value = normalized.replyTo
    }
    if (ui.smtpTestRecipient) {
      ui.smtpTestRecipient.value = ''
    }
    if (ui.smtpSettingsSubmit) {
      ui.smtpSettingsSubmit.disabled = false
    }
    if (ui.smtpTestSend) {
      ui.smtpTestSend.disabled = false
    }
  }

  function collectSmtpSettingsFormValues() {
    const portValue = String(ui.smtpSettingsPort?.value || '').trim()
    return {
      enabled: Boolean(ui.smtpSettingsEnabled && ui.smtpSettingsEnabled.checked),
      host: String(ui.smtpSettingsHost?.value || '').trim(),
      port: portValue ? Number.parseInt(portValue, 10) : undefined,
      secure: Boolean(ui.smtpSettingsSecure && ui.smtpSettingsSecure.checked),
      username: String(ui.smtpSettingsUsername?.value || '').trim(),
      password: String(ui.smtpSettingsPassword?.value || ''),
      fromName: String(ui.smtpSettingsFromName?.value || '').trim(),
      fromAddress: String(ui.smtpSettingsFromAddress?.value || '').trim(),
      replyTo: String(ui.smtpSettingsReplyTo?.value || '').trim()
    }
  }

  function focusSmtpSettingsHost() {
    if (!ui.smtpSettingsHost) {
      return
    }

    window.requestAnimationFrame(() => {
      ui.smtpSettingsHost.focus()
      if (typeof ui.smtpSettingsHost.select === 'function') {
        ui.smtpSettingsHost.select()
      }
    })
  }

  function closeSmtpSettingsModal() {
    state.smtpSettingsOpen = false
    if (ui.smtpSettingsModal) {
      ui.smtpSettingsModal.hidden = true
    }
    setSmtpSettingsBusy(false)
    setSmtpSettingsMessage('')
    closeSettingsMenu()
  }

  async function loadSmtpSettings(options = {}) {
    if (!state.authEnabled || !state.authCanManageUsers) {
      renderSmtpSettings({
        enabled: false,
        host: '',
        port: 587,
        secure: false,
        username: '',
        hasPassword: false,
        fromName: '',
        fromAddress: '',
        replyTo: ''
      })
      return true
    }

    const showBusy = options.showBusy !== false
    const silent = Boolean(options.silent)
    if (showBusy) {
      setSmtpSettingsBusy(true, 'Loading...')
    }
    if (!silent) {
      setSmtpSettingsMessage('Loading SMTP settings...')
    }

    try {
      const response = await fetchJson('/api/settings/smtp', {
        cache: 'no-store'
      })
      renderSmtpSettings(response && response.settings ? response.settings : {})
      if (!silent) {
        setSmtpSettingsMessage('')
      }
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setSmtpSettingsMessage('Admin access required.', 'error')
        return false
      }
      setSmtpSettingsMessage(`Unable to load SMTP settings: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      if (showBusy) {
        setSmtpSettingsBusy(false)
      }
    }
  }

  function openSmtpSettingsModal() {
    if (!state.authEnabled || !state.authenticated || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    if (!ui.smtpSettingsModal) {
      return
    }

    closeActivityLogModal()
    closeUserManagementModal()
    closeSettingsMenu()
    ui.smtpSettingsModal.hidden = false
    state.smtpSettingsOpen = true
    setSmtpSettingsMessage('')
    void (async () => {
      const loaded = await loadSmtpSettings({
        showBusy: true,
        silent: false
      })

      if (loaded && state.smtpSettingsOpen) {
        focusSmtpSettingsHost()
      }
    })()
  }

  async function saveSmtpSettings() {
    if (!state.authEnabled || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    const payload = collectSmtpSettingsFormValues()
    setSmtpSettingsMessage('')
    setSmtpSettingsBusy(true, 'Saving...')
    try {
      const response = await fetchJson('/api/settings/smtp', {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      })

      if (!response.settings) {
        throw new Error('Unable to save SMTP settings.')
      }

      renderSmtpSettings(response.settings)
      setSmtpSettingsMessage('SMTP settings saved.', 'success')
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setSmtpSettingsMessage('Admin access required.', 'error')
        return
      }
      setSmtpSettingsMessage(`Unable to save SMTP settings: ${getErrorMessage(error)}`, 'error')
    } finally {
      setSmtpSettingsBusy(false)
    }
  }

  async function sendSmtpTestEmail() {
    if (!state.authEnabled || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    const payload = collectSmtpSettingsFormValues()
    const recipient = String(ui.smtpTestRecipient?.value || '').trim()
    if (!recipient) {
      setSmtpSettingsMessage('Enter a test recipient.', 'error')
      return
    }

    setSmtpSettingsMessage('')
    setSmtpSettingsBusy(true, 'Sending...')
    try {
      const response = await fetchJson('/api/settings/smtp/test', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          ...payload,
          recipient
        })
      })

      if (!response.success) {
        throw new Error('Unable to send test email.')
      }

      setSmtpSettingsMessage(`Test email sent to ${recipient}.`, 'success')
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setSmtpSettingsMessage('Admin access required.', 'error')
        return
      }
      setSmtpSettingsMessage(`Unable to send test email: ${getErrorMessage(error)}`, 'error')
    } finally {
      setSmtpSettingsBusy(false)
    }
  }

  function closeActivityLogModal() {
    state.activityLogOpen = false
    if (ui.activityLogModal) {
      ui.activityLogModal.hidden = true
    }
    setActivityLogBusy(false)
    setActivityLogMessage('')
    closeSettingsMenu()
  }

  function isCurrentManagedUser(username) {
    return Boolean(state.authUser) && normalizeAuthUserKey(username) === normalizeAuthUserKey(state.authUser)
  }

  function resolveDefaultUserActivityUsername() {
    if (!Array.isArray(state.authUsers) || !state.authUsers.length) {
      return ''
    }

    const currentUser = String(state.authUser || '').trim()
    if (currentUser) {
      const match = state.authUsers.find((user) => isCurrentManagedUser(user.username))
      if (match) {
        return match.username
      }
    }

    return state.authUsers[0]?.username || ''
  }

  function updateUserActivitySelectionLabel() {
    if (!ui.userActivitySelected) {
      return
    }

    const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
    if (!selectedUsername) {
      ui.userActivitySelected.textContent = 'Select a user to view activity.'
      ui.userActivitySelected.dataset.empty = 'true'
      if (ui.userActivityRefresh) {
        ui.userActivityRefresh.disabled = true
      }
      return
    }

    const isVisible = state.authUsers.some(
      (user) => normalizeAuthUserKey(user.username) === normalizeAuthUserKey(selectedUsername)
    )
    ui.userActivitySelected.textContent = isVisible
      ? `Viewing activity for ${selectedUsername}.`
      : `${selectedUsername} no longer appears in the user list.`
    ui.userActivitySelected.dataset.empty = 'false'
    if (ui.userActivityRefresh) {
      ui.userActivityRefresh.disabled = state.userActivityLoading
    }
    if (ui.userActivityExport) {
      ui.userActivityExport.disabled = state.userActivityLoading || !selectedUsername
    }
  }

  function setUserActivityBusy(isBusy, label = 'Refresh') {
    state.userActivityLoading = isBusy
    if (ui.userActivityRefresh) {
      ui.userActivityRefresh.disabled = isBusy || !String(state.selectedUserActivityUsername || '').trim()
      ui.userActivityRefresh.textContent = isBusy ? label : 'Refresh'
    }
    if (ui.userActivityExport) {
      ui.userActivityExport.disabled =
        isBusy || !String(state.selectedUserActivityUsername || '').trim()
    }
  }

  function setUserActivityMessage(message = '', tone = 'neutral') {
    if (!ui.userActivityMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.userActivityMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.userActivityMessage.dataset.tone = tone
    } else {
      delete ui.userActivityMessage.dataset.tone
    }
  }

  function renderUserActivityEntries(entries) {
    const normalizedEntries = Array.isArray(entries)
      ? entries
          .map((entry) => ({
            timestamp: String(entry && entry.timestamp ? entry.timestamp : ''),
            actor: {
              username: String(entry && entry.actor && entry.actor.username ? entry.actor.username : 'anonymous'),
              authenticated: Boolean(entry && entry.actor && entry.actor.authenticated),
              admin: Boolean(entry && entry.actor && entry.actor.admin)
            },
            action: String(entry && entry.action ? entry.action : ''),
            target: String(entry && entry.target ? entry.target : ''),
            outcome: String(entry && entry.outcome ? entry.outcome : 'success'),
            request: {
              method: String(entry && entry.request && entry.request.method ? entry.request.method : ''),
              path: String(entry && entry.request && entry.request.path ? entry.request.path : ''),
              origin: String(entry && entry.request && entry.request.origin ? entry.request.origin : ''),
              ip: String(entry && entry.request && entry.request.ip ? entry.request.ip : '')
            },
            metadata:
              entry && entry.metadata && typeof entry.metadata === 'object' && !Array.isArray(entry.metadata)
                ? entry.metadata
                : {}
          }))
          .filter((entry) => Boolean(entry.timestamp) && Boolean(entry.action))
      : []

    state.userActivityEntries = normalizedEntries
    state.userActivityLoaded = true

    if (ui.userActivityCountBadge) {
      ui.userActivityCountBadge.textContent = String(normalizedEntries.length)
    }

    updateUserActivitySelectionLabel()

    if (!ui.userActivityList) {
      return
    }

    const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
    if (!selectedUsername) {
      ui.userActivityList.innerHTML =
        '<div class="activity-log-empty">Select a user to view activity.</div>'
      return
    }

    if (!normalizedEntries.length) {
      ui.userActivityList.innerHTML =
        `<div class="activity-log-empty">No activity has been recorded for ${escapeHtml(selectedUsername)} yet.</div>`
      return
    }

    ui.userActivityList.innerHTML = normalizedEntries
      .map((entry) => {
        const outcome = String(entry.outcome || 'success').toLowerCase()
        const outcomeLabel = outcome === 'denied' ? 'Denied' : outcome === 'failure' ? 'Failure' : 'Success'
        const requestParts = [
          entry.request.method,
          entry.request.path,
          entry.request.ip ? `IP ${entry.request.ip}` : '',
          entry.request.origin ? `Origin ${entry.request.origin}` : ''
        ].filter(Boolean)
        const metadataText = formatActivityLogMetadata(entry.metadata)

        return `
          <article class="activity-log-item" data-outcome="${escapeAttr(outcome)}">
            <div class="activity-log-item-head">
              <span class="activity-log-item-time">${escapeHtml(formatActivityLogTimestamp(entry.timestamp))}</span>
              <span class="activity-log-item-outcome">${escapeHtml(outcomeLabel)}</span>
            </div>
            <div class="activity-log-item-main">
              <strong class="activity-log-item-user">${escapeHtml(entry.actor.username || 'anonymous')}</strong>
              <span class="activity-log-item-action">${escapeHtml(entry.action)}</span>
              <span class="activity-log-item-target">${escapeHtml(entry.target || 'No target')}</span>
            </div>
            <div class="activity-log-item-meta">${escapeHtml(requestParts.join(' · ') || 'No request details')}</div>
            ${
              metadataText
                ? `<div class="activity-log-item-extra">${escapeHtml(metadataText)}</div>`
                : ''
            }
          </article>
        `
      })
      .join('')
  }

  async function loadUserActivityLog(username, options = {}) {
    if (!state.authEnabled || !state.authCanManageUsers) {
      renderUserActivityEntries([])
      return true
    }

    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      state.selectedUserActivityUsername = ''
      renderUserActivityEntries([])
      return true
    }

    const showBusy = options.showBusy !== false
    const silent = Boolean(options.silent)
    if (showBusy) {
      setUserActivityBusy(true, 'Loading...')
    }
    if (!silent) {
      setUserActivityMessage(`Loading activity for ${normalizedUsername}...`)
    }

    try {
      const response = await fetchJson(
        `/api/activity-log?limit=100&username=${encodeURIComponent(normalizedUsername)}`,
        {
          cache: 'no-store'
        }
      )
      renderUserActivityEntries((response && response.entries) || [])
      if (!silent) {
        setUserActivityMessage('')
      }
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserActivityMessage('Admin access required.', 'error')
        return false
      }
      setUserActivityMessage(`Unable to load activity: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      if (showBusy) {
        setUserActivityBusy(false)
      }
    }
  }

  function selectUserActivity(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return
    }

    state.selectedUserActivityUsername = normalizedUsername
    renderAuthUsers(state.authUsers)
    void loadUserActivityLog(normalizedUsername, {
      showBusy: true,
      silent: false
    })
  }

  async function deleteUser(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return
    }

    if (isCurrentManagedUser(normalizedUsername)) {
      setUserManagementMessage('The admin account cannot be deleted.', 'error')
      return
    }

    if (!window.confirm(`Delete ${normalizedUsername}? Active sessions will be revoked.`)) {
      return
    }

    setUserManagementMessage('')
    setUserManagementBusy(true, 'Deleting...')
    try {
      const response = await fetchJson(`/api/auth/users/${encodeURIComponent(normalizedUsername)}`, {
        method: 'DELETE'
      })
      if (!response.user || !response.user.username) {
        throw new Error('Unable to delete user.')
      }

      setUserManagementMessage(`Deleted ${response.user.username}.`, 'success')
      await loadAuthUsers({
        showBusy: false,
        silent: true
      })
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 404) {
        setUserManagementMessage('User not found.', 'error')
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 400) {
        setUserManagementMessage(getErrorMessage(error), 'error')
        return
      }
      setUserManagementMessage(`Unable to delete user: ${getErrorMessage(error)}`, 'error')
    } finally {
      setUserManagementBusy(false)
    }
  }

  async function inviteUser() {
    const username = String(ui.userManagementUsername?.value || '').trim()
    const recipientEmail = String(ui.userManagementEmail?.value || '').trim()
    if (!username || !recipientEmail) {
      setUserManagementMessage('Enter both a username and email address.', 'error')
      return
    }

    setUserManagementMessage('')
    setUserManagementBusy(true, 'Inviting...')
    try {
      const response = await fetchJson('/api/auth/users', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          username,
          recipientEmail
        })
      })

      if (!response.user || !response.user.username) {
        throw new Error('Unable to create invite.')
      }

      const inviteParts = [`Invited ${response.user.username}.`]
      if (response.emailSent) {
        inviteParts.push('The invite email was sent.')
      }
      if (response.inviteUrl) {
        inviteParts.push(`Invite link: ${response.inviteUrl}`)
      }
      setUserManagementMessage(inviteParts.join(' '), 'success')
      if (ui.userManagementForm) {
        ui.userManagementForm.reset()
      }
      await loadAuthUsers({
        showBusy: false,
        silent: true
      })
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 409) {
        setUserManagementMessage(getErrorMessage(error), 'error')
        return
      }
      setUserManagementMessage(`Unable to invite user: ${getErrorMessage(error)}`, 'error')
    } finally {
      setUserManagementBusy(false)
    }
  }

  async function resendInvite(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return
    }

    setUserManagementMessage('')
    setUserManagementBusy(true, 'Sending...')
    try {
      const response = await fetchJson(`/api/auth/users/${encodeURIComponent(normalizedUsername)}/invite/resend`, {
        method: 'POST'
      })
      if (!response.user || !response.user.username) {
        throw new Error('Unable to resend invite.')
      }

      const parts = [`Resent invite for ${response.user.username}.`]
      if (response.emailSent) {
        parts.push('The invite email was sent.')
      }
      if (response.inviteUrl) {
        parts.push(`Invite link: ${response.inviteUrl}`)
      }
      setUserManagementMessage(parts.join(' '), 'success')
      await loadAuthUsers({
        showBusy: false,
        silent: true
      })
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return
      }
      setUserManagementMessage(`Unable to resend invite: ${getErrorMessage(error)}`, 'error')
    } finally {
      setUserManagementBusy(false)
    }
  }

  async function revokeInvite(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return
    }

    if (!window.confirm(`Revoke the invite for ${normalizedUsername}?`)) {
      return
    }

    setUserManagementMessage('')
    setUserManagementBusy(true, 'Revoking...')
    try {
      const response = await fetchJson(`/api/auth/users/${encodeURIComponent(normalizedUsername)}/invite`, {
        method: 'DELETE'
      })
      if (!response.user || !response.user.username) {
        throw new Error('Unable to revoke invite.')
      }

      setUserManagementMessage(`Revoked invite for ${response.user.username}.`, 'success')
      await loadAuthUsers({
        showBusy: false,
        silent: true
      })
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return
      }
      setUserManagementMessage(`Unable to revoke invite: ${getErrorMessage(error)}`, 'error')
    } finally {
      setUserManagementBusy(false)
    }
  }

  async function resetUserMfa(username) {
    const normalizedUsername = String(username || '').trim()
    if (!normalizedUsername) {
      return
    }

    if (!window.confirm(`Reset MFA for ${normalizedUsername}? Recovery codes will be invalidated.`)) {
      return
    }

    setUserManagementMessage('')
    setUserManagementBusy(true, 'Resetting...')
    try {
      const response = await fetchJson(`/api/auth/users/${encodeURIComponent(normalizedUsername)}/mfa/reset`, {
        method: 'POST'
      })
      if (!response.user || !response.user.username) {
        throw new Error('Unable to reset MFA.')
      }

      setUserManagementMessage(`Reset MFA for ${response.user.username}.`, 'success')
      await loadAuthUsers({
        showBusy: false,
        silent: true
      })
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return
      }
      setUserManagementMessage(`Unable to reset MFA: ${getErrorMessage(error)}`, 'error')
    } finally {
      setUserManagementBusy(false)
    }
  }

  function openActivityLogModal() {
    if (!state.authEnabled || !state.authenticated || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    if (!ui.activityLogModal) {
      return
    }

    closeUserManagementModal()
    closeSettingsMenu()
    ui.activityLogModal.hidden = false
    state.activityLogOpen = true
    setActivityLogMessage('')
    window.requestAnimationFrame(() => {
      if (ui.activityLogClose && typeof ui.activityLogClose.focus === 'function') {
        ui.activityLogClose.focus()
      }
    })
    void (async () => {
      const loaded = await loadActivityLog({
        showBusy: true,
        silent: false
      })

      if (loaded && state.activityLogOpen) {
        window.requestAnimationFrame(() => {
          if (ui.activityLogClose && typeof ui.activityLogClose.focus === 'function') {
            ui.activityLogClose.focus()
          }
        })
      }
    })()
  }

  function openUserManagementModal() {
    if (!state.authEnabled || !state.authenticated || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    if (!ui.userManagementModal) {
      return
    }

    closeActivityLogModal()
    closeSettingsMenu()
    ui.userManagementModal.hidden = false
    state.userManagementOpen = true
    setUserManagementMessage('')
    void (async () => {
      const loaded = await loadAuthUsers({
        showBusy: true,
        silent: false
      })

      if (loaded && state.userManagementOpen) {
        if (!String(state.selectedUserActivityUsername || '').trim()) {
          state.selectedUserActivityUsername = resolveDefaultUserActivityUsername()
        }
        renderAuthUsers(state.authUsers)
        const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
        if (selectedUsername) {
          void loadUserActivityLog(selectedUsername, {
            showBusy: true,
            silent: false
          })
        }
        focusUserManagementUsername()
      }
    })()
  }

  function setUserManagementBusy(isBusy, label = 'Invite user') {
    state.authUsersLoading = isBusy
    if (ui.userManagementSubmit) {
      ui.userManagementSubmit.disabled = isBusy
      ui.userManagementSubmit.textContent = isBusy ? label : 'Invite user'
    }
    if (ui.userManagementUsername) {
      ui.userManagementUsername.disabled = isBusy
    }
    if (ui.userManagementEmail) {
      ui.userManagementEmail.disabled = isBusy
    }
    if (ui.userList) {
      const actionButtons = ui.userList.querySelectorAll(
        'button[data-action="select-user-activity"], button[data-action="delete-user"], button[data-action="resend-invite"], button[data-action="revoke-invite"], button[data-action="reset-mfa"]'
      )
      actionButtons.forEach((button) => {
        const username = String(button.getAttribute('data-username') || '').trim()
        if (!username) {
          button.disabled = isBusy
          return
        }
        if (button.getAttribute('data-action') === 'delete-user' && isCurrentManagedUser(username)) {
          button.disabled = true
          return
        }
        if (button.getAttribute('data-action') === 'revoke-invite' && isCurrentManagedUser(username)) {
          button.disabled = true
          return
        }
        button.disabled = isBusy
      })
    }
  }

  function setUserManagementMessage(message = '', tone = 'neutral') {
    if (!ui.userManagementMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.userManagementMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.userManagementMessage.dataset.tone = tone
    } else {
      delete ui.userManagementMessage.dataset.tone
    }
  }

  function renderAuthUsers(users) {
    const normalizedUsers = Array.isArray(users)
      ? users
          .map((user) => ({
            username: String(user && user.username ? user.username : '').trim(),
            createdAt: String(user && user.createdAt ? user.createdAt : ''),
            recipientEmail: String(user && user.recipientEmail ? user.recipientEmail : ''),
            inviteStatus: String(user && user.inviteStatus ? user.inviteStatus : 'active'),
            inviteSentAt: String(user && user.inviteSentAt ? user.inviteSentAt : ''),
            inviteExpiresAt: String(user && user.inviteExpiresAt ? user.inviteExpiresAt : ''),
            inviteAcceptedAt: String(user && user.inviteAcceptedAt ? user.inviteAcceptedAt : ''),
            inviteRevokedAt: String(user && user.inviteRevokedAt ? user.inviteRevokedAt : ''),
            mfaEnabled: Boolean(user && user.mfaEnabled),
            mfaEnrolledAt: String(user && user.mfaEnrolledAt ? user.mfaEnrolledAt : '')
          }))
          .filter((user) => Boolean(user.username))
      : []

    state.authUsers = normalizedUsers
    state.authUsersLoaded = true

    if (ui.userCountBadge) {
      ui.userCountBadge.textContent = String(normalizedUsers.length)
    }

    if (!ui.userList) {
      return
    }

    if (!normalizedUsers.length) {
      ui.userList.innerHTML =
        '<div class="user-list-empty">No invite-based users are configured.</div>'
      updateUserActivitySelectionLabel()
      return
    }

    const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
    ui.userList.innerHTML = normalizedUsers
      .map((user) => {
        const isCurrent = isCurrentManagedUser(user.username)
        const isSelected =
          Boolean(selectedUsername) &&
          normalizeAuthUserKey(user.username) === normalizeAuthUserKey(selectedUsername)
        const inviteStatus = String(user.inviteStatus || 'active').toLowerCase()
        const statusLabel =
          inviteStatus === 'pending'
            ? 'Pending invite'
            : inviteStatus === 'revoked'
              ? 'Invite revoked'
              : inviteStatus === 'expired'
                ? 'Invite expired'
                : 'Active'
        const statusClass =
          inviteStatus === 'pending'
            ? 'accent'
            : inviteStatus === 'revoked' || inviteStatus === 'expired'
              ? 'danger'
              : 'green'
        const statusChips = [
          `<span class="chip ${statusClass}">${escapeHtml(statusLabel)}</span>`,
          user.mfaEnabled ? '<span class="chip green">MFA enabled</span>' : '',
          user.recipientEmail ? `<span class="chip">${escapeHtml(user.recipientEmail)}</span>` : ''
        ].filter(Boolean)
        const userMeta = [
          user.inviteSentAt ? `Invited ${formatActivityLogTimestamp(user.inviteSentAt)}` : '',
          user.inviteAcceptedAt ? `Accepted ${formatActivityLogTimestamp(user.inviteAcceptedAt)}` : '',
          user.inviteRevokedAt ? `Revoked ${formatActivityLogTimestamp(user.inviteRevokedAt)}` : '',
          user.mfaEnrolledAt ? `MFA ${formatActivityLogTimestamp(user.mfaEnrolledAt)}` : ''
        ]
          .filter(Boolean)
          .join(' · ')
        return `
          <div class="user-item-row${isSelected ? ' selected' : ''}">
            <button
              class="user-item user-item-select${isSelected ? ' selected' : ''}"
              type="button"
              data-action="select-user-activity"
              data-username="${escapeAttr(user.username)}"
              aria-pressed="${isSelected ? 'true' : 'false'}"
            >
              <span class="user-item-copy">
                <span class="user-item-name">${escapeHtml(user.username)}</span>
                <span class="user-item-badges">${statusChips.join('')}</span>
                <span class="user-item-meta">${
                  isCurrent ? 'Current admin' : userMeta || 'View activity'
                }</span>
              </span>
              <span class="user-item-hint">${isSelected ? 'Selected' : 'Open'}</span>
            </button>
            <div class="user-item-actions">
              ${
                inviteStatus === 'pending' || inviteStatus === 'expired'
                  ? `
                    <button
                      class="ghost-button small icon-button user-resend-button"
                      type="button"
                      data-action="resend-invite"
                      data-username="${escapeAttr(user.username)}"
                      aria-label="Resend invite for ${escapeAttr(user.username)}"
                      title="Resend invite"
                    >
                      ${renderPaperPlaneIcon()}
                    </button>
                    <button
                      class="ghost-button small icon-button user-revoke-button"
                      type="button"
                      data-action="revoke-invite"
                      data-username="${escapeAttr(user.username)}"
                      aria-label="Revoke invite for ${escapeAttr(user.username)}"
                      title="Revoke invite"
                    >
                      ${renderRotateIcon()}
                    </button>
                  `
                  : ''
              }
              ${
                user.mfaEnabled
                  ? `
                    <button
                      class="ghost-button small icon-button user-mfa-reset-button"
                      type="button"
                      data-action="reset-mfa"
                      data-username="${escapeAttr(user.username)}"
                      aria-label="Reset MFA for ${escapeAttr(user.username)}"
                      title="Reset MFA"
                    >
                      ${renderKeyIcon()}
                    </button>
                  `
                  : ''
              }
              <button
                class="ghost-button small icon-button user-delete-button"
                type="button"
                data-action="delete-user"
                data-username="${escapeAttr(user.username)}"
                aria-label="Delete ${escapeAttr(user.username)}"
                title="Delete user"
                ${isCurrent ? 'disabled' : ''}
              >
                ${renderTrashIcon()}
              </button>
            </div>
          </div>
        `
      })
      .join('')

    updateUserActivitySelectionLabel()
  }

  async function loadAuthUsers(options = {}) {
    if (!state.authEnabled || !state.authCanManageUsers) {
      renderAuthUsers([])
      updateUserManagementVisibility()
      return true
    }

    const showBusy = options.showBusy !== false
    const silent = Boolean(options.silent)
    if (showBusy) {
      setUserManagementBusy(true, 'Loading...')
    }
    if (!silent) {
      setUserManagementMessage('Loading users...')
    }

    try {
      const response = await fetchJson('/api/auth/users', {
        cache: 'no-store'
      })
      renderAuthUsers((response && response.users) || [])
      if (!silent) {
        setUserManagementMessage('')
      }
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setUserManagementMessage('Admin access required.', 'error')
        return false
      }
      setUserManagementMessage(`Unable to load users: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      if (showBusy) {
        setUserManagementBusy(false)
      }
    }
  }

  function formatActivityLogTimestamp(value) {
    const date = new Date(String(value || ''))
    if (Number.isNaN(date.getTime())) {
      return String(value || '')
    }
    return date.toLocaleString()
  }

  function formatActivityLogValue(value) {
    if (value == null || value === '') {
      return ''
    }
    if (Array.isArray(value)) {
      return value.map((item) => formatActivityLogValue(item)).filter(Boolean).join(', ')
    }
    if (typeof value === 'object') {
      try {
        return JSON.stringify(value)
      } catch {
        return '[object]'
      }
    }
    return String(value)
  }

  function formatActivityLogMetadata(metadata) {
    if (!metadata || typeof metadata !== 'object') {
      return ''
    }

    return Object.entries(metadata)
      .map(([key, value]) => {
        const text = formatActivityLogValue(value)
        return text ? `${key}: ${text}` : ''
      })
      .filter(Boolean)
      .join(' · ')
  }

  function setActivityLogBusy(isBusy, label = 'Refresh') {
    state.activityLogLoading = isBusy
    if (ui.activityLogRefresh) {
      ui.activityLogRefresh.disabled = isBusy
      ui.activityLogRefresh.textContent = isBusy ? label : 'Refresh'
    }
    if (ui.activityLogExport) {
      ui.activityLogExport.disabled = isBusy
    }
  }

  function setActivityLogMessage(message = '', tone = 'neutral') {
    if (!ui.activityLogMessage) {
      return
    }

    const normalizedMessage = String(message || '')
    ui.activityLogMessage.textContent = normalizedMessage
    if (normalizedMessage) {
      ui.activityLogMessage.dataset.tone = tone
    } else {
      delete ui.activityLogMessage.dataset.tone
    }
  }

  function renderActivityLogEntries(entries) {
    const normalizedEntries = Array.isArray(entries)
      ? entries
          .map((entry) => ({
            timestamp: String(entry && entry.timestamp ? entry.timestamp : ''),
            actor: {
              username: String(entry && entry.actor && entry.actor.username ? entry.actor.username : 'anonymous'),
              authenticated: Boolean(entry && entry.actor && entry.actor.authenticated),
              admin: Boolean(entry && entry.actor && entry.actor.admin)
            },
            action: String(entry && entry.action ? entry.action : ''),
            target: String(entry && entry.target ? entry.target : ''),
            outcome: String(entry && entry.outcome ? entry.outcome : 'success'),
            request: {
              method: String(entry && entry.request && entry.request.method ? entry.request.method : ''),
              path: String(entry && entry.request && entry.request.path ? entry.request.path : ''),
              origin: String(entry && entry.request && entry.request.origin ? entry.request.origin : ''),
              ip: String(entry && entry.request && entry.request.ip ? entry.request.ip : '')
            },
            metadata:
              entry && entry.metadata && typeof entry.metadata === 'object' && !Array.isArray(entry.metadata)
                ? entry.metadata
                : {}
          }))
          .filter((entry) => Boolean(entry.timestamp) && Boolean(entry.action))
      : []

    state.activityLogEntries = normalizedEntries
    state.activityLogLoaded = true

    if (ui.activityLogCountBadge) {
      ui.activityLogCountBadge.textContent = String(normalizedEntries.length)
    }

    if (!ui.activityLogList) {
      return
    }

    if (!normalizedEntries.length) {
      ui.activityLogList.innerHTML =
        '<div class="activity-log-empty">No activity has been recorded yet.</div>'
      return
    }

    ui.activityLogList.innerHTML = normalizedEntries
      .map((entry) => {
        const outcome = String(entry.outcome || 'success').toLowerCase()
        const outcomeLabel = outcome === 'denied' ? 'Denied' : outcome === 'failure' ? 'Failure' : 'Success'
        const requestParts = [
          entry.request.method,
          entry.request.path,
          entry.request.ip ? `IP ${entry.request.ip}` : '',
          entry.request.origin ? `Origin ${entry.request.origin}` : ''
        ].filter(Boolean)
        const metadataText = formatActivityLogMetadata(entry.metadata)

        return `
          <article class="activity-log-item" data-outcome="${escapeAttr(outcome)}">
            <div class="activity-log-item-head">
              <span class="activity-log-item-time">${escapeHtml(formatActivityLogTimestamp(entry.timestamp))}</span>
              <span class="activity-log-item-outcome">${escapeHtml(outcomeLabel)}</span>
            </div>
            <div class="activity-log-item-main">
              <strong class="activity-log-item-user">${escapeHtml(entry.actor.username || 'anonymous')}</strong>
              <span class="activity-log-item-action">${escapeHtml(entry.action)}</span>
              <span class="activity-log-item-target">${escapeHtml(entry.target || 'No target')}</span>
            </div>
            <div class="activity-log-item-meta">${escapeHtml(requestParts.join(' · ') || 'No request details')}</div>
            ${
              metadataText
                ? `<div class="activity-log-item-extra">${escapeHtml(metadataText)}</div>`
                : ''
            }
          </article>
        `
      })
      .join('')
  }

  async function loadActivityLog(options = {}) {
    if (!state.authEnabled || !state.authCanManageUsers) {
      renderActivityLogEntries([])
      updateUserManagementVisibility()
      return true
    }

    const showBusy = options.showBusy !== false
    const silent = Boolean(options.silent)
    if (showBusy) {
      setActivityLogBusy(true, 'Loading...')
    }
    if (!silent) {
      setActivityLogMessage('Loading activity log...')
    }

    try {
      const response = await fetchJson('/api/activity-log?limit=100', {
        cache: 'no-store'
      })
      renderActivityLogEntries((response && response.entries) || [])
      if (!silent) {
        setActivityLogMessage('')
      }
      return true
    } catch (error) {
      if (error && typeof error === 'object' && Number(error.statusCode) === 401) {
        handleAuthFailure()
        return false
      }
      if (error && typeof error === 'object' && Number(error.statusCode) === 403) {
        setActivityLogMessage('Admin access required.', 'error')
        return false
      }
      setActivityLogMessage(`Unable to load activity log: ${getErrorMessage(error)}`, 'error')
      return false
    } finally {
      if (showBusy) {
        setActivityLogBusy(false)
      }
    }
  }

  function buildActivityLogCsvUrl(username = '') {
    const params = new URLSearchParams()
    const normalizedUsername = String(username || '').trim()
    if (normalizedUsername) {
      params.set('username', normalizedUsername)
    }
    const query = params.toString()
    return query ? `/api/activity-log.csv?${query}` : '/api/activity-log.csv'
  }

  function downloadActivityLogCsv(username = '') {
    if (!state.authEnabled || !state.authenticated || !state.authCanManageUsers) {
      setStatus('Admin access required.', 'error')
      return
    }

    const normalizedUsername = String(username || '').trim()
    if (normalizedUsername) {
      window.location.assign(buildActivityLogCsvUrl(normalizedUsername))
      return
    }

    window.location.assign(buildActivityLogCsvUrl())
  }

  function getErrorMessage(error) {
    if (!error) {
      return 'Request failed'
    }
    if (error && typeof error === 'object' && 'payload' in error) {
      const payload = error.payload
      if (payload && typeof payload === 'object' && typeof payload.error === 'string') {
        return payload.error
      }
    }
    return error.message || String(error)
  }

  function readStorageBool(key, fallback) {
    const raw = localStorage.getItem(key)
    if (raw == null) {
      return fallback
    }
    return !['0', 'false', 'no', 'off'].includes(raw.toLowerCase())
  }

  function normalizeStorageNamespace(value) {
    const text = String(value || '').trim().toLowerCase()
    if (!text) {
      return 'local'
    }

    return text.replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '') || 'local'
  }

  function getWorkspaceStorageNamespace() {
    if (!state.authEnabled || !state.authenticated) {
      return 'local'
    }

    return normalizeStorageNamespace(state.authUser)
  }

  function getWorkspaceStorageKey(key) {
    return `${STORAGE_KEYS[key]}::${getWorkspaceStorageNamespace()}`
  }

  function readWorkspaceStorageItem(key, fallback = '') {
    const scopedKey = getWorkspaceStorageKey(key)
    const stored = localStorage.getItem(scopedKey)
    if (stored != null) {
      return stored
    }

    const legacyKey = STORAGE_KEYS[key]
    const legacy = localStorage.getItem(legacyKey)
    if (legacy != null) {
      try {
        localStorage.setItem(scopedKey, legacy)
      } catch (error) {
        // Ignore storage migration errors in restricted browser contexts.
      }
      return legacy
    }

    return fallback
  }

  function readWorkspaceStorageBool(key, fallback) {
    const raw = readWorkspaceStorageItem(key, null)
    if (raw == null) {
      return fallback
    }
    return !['0', 'false', 'no', 'off'].includes(String(raw).toLowerCase())
  }

  function writeWorkspaceStorageItem(key, value) {
    const scopedKey = getWorkspaceStorageKey(key)
    if (value == null || value === '') {
      localStorage.removeItem(scopedKey)
      return
    }

    localStorage.setItem(scopedKey, String(value))
  }

  function removeWorkspaceStorageItem(key) {
    localStorage.removeItem(getWorkspaceStorageKey(key))
  }

  function normalizeTheme(value) {
    return String(value || '').toLowerCase() === 'dark' ? 'dark' : 'light'
  }

  function getSystemThemePreference() {
    try {
      if (
        typeof window.matchMedia === 'function' &&
        window.matchMedia('(prefers-color-scheme: dark)').matches
      ) {
        return 'dark'
      }
    } catch (error) {
      // Ignore media query errors and fall back to light mode.
    }

    return 'light'
  }

  function readThemePreference() {
    const stored = localStorage.getItem(STORAGE_KEYS.theme)
    if (stored === 'dark' || stored === 'light') {
      return stored
    }

    return getSystemThemePreference()
  }

  function persistThemePreference(theme) {
    try {
      localStorage.setItem(STORAGE_KEYS.theme, normalizeTheme(theme))
    } catch (error) {
      // Ignore storage failures in private browsing or restricted contexts.
    }
  }

  function loadWorkspaceState() {
    const searchScope = readWorkspaceStorageItem('searchScope', 'pst')
    const mailboxScopeView = readWorkspaceStorageItem('mailboxScopeView', 'search')
    const catalogMode = readWorkspaceStorageItem('catalogMode', 'active')
    const flaggedBundleScope = readWorkspaceStorageItem('flaggedBundleScope', 'all')

    state.selectedCasePath = normalizeScopePath(readWorkspaceStorageItem('casePath', ''))
    state.selectedScopePath = normalizeScopePath(readWorkspaceStorageItem('scopePath', ''))
    state.selectedScopeLabel = getScopeLabel(state.selectedScopePath)
    state.selectedPstFileName = readWorkspaceStorageItem('pstFileName', '') || null
    state.currentFolderId = readWorkspaceStorageItem('folderId', '') || null
    state.selectedMessageId = readWorkspaceStorageItem('messageId', '') || null
    state.query = readWorkspaceStorageItem('query', '')
    state.searchScope = ['pst', 'search', 'all'].includes(searchScope) ? searchScope : 'pst'
    state.mailOnly = readWorkspaceStorageBool('mailOnly', true)
    state.sort = readWorkspaceStorageItem('sort', 'date-desc') || 'date-desc'
    state.reviewFlaggedOnly = readWorkspaceStorageBool('reviewFlaggedOnly', false)
    state.reviewTaggedOnly = readWorkspaceStorageBool('reviewTaggedOnly', false)
    state.mailboxScopeView = mailboxScopeView === 'all' ? 'all' : 'search'
    state.mailboxFilter = readWorkspaceStorageItem('mailboxFilter', '')
    state.catalogMode = catalogMode === 'removed' ? 'removed' : 'active'
    state.bundleScope = ['all', 'search', 'pst'].includes(flaggedBundleScope)
      ? flaggedBundleScope
      : 'all'
    applyStateToControls()
  }

  function updateThemeToggleButtons(theme) {
    const normalized = normalizeTheme(theme)
    const nextTheme = normalized === 'dark' ? 'light' : 'dark'
    const label = nextTheme === 'dark' ? 'Dark mode' : 'Light mode'
    const ariaLabel = `Switch to ${nextTheme} mode`

    for (const button of Array.isArray(ui.themeToggleButtons) ? ui.themeToggleButtons : []) {
      button.textContent = label
      button.setAttribute('aria-pressed', String(normalized === 'dark'))
      button.setAttribute('aria-label', ariaLabel)
      button.title = ariaLabel
    }
  }

  function applyTheme(theme) {
    const normalized = normalizeTheme(theme)
    state.theme = normalized
    document.documentElement.dataset.theme = normalized
    document.documentElement.style.colorScheme = normalized
    updateThemeToggleButtons(normalized)

    if (state.currentMessageDetail) {
      renderMessageDetail()
    }
  }

  function toggleTheme() {
    const nextTheme = state.theme === 'dark' ? 'light' : 'dark'
    applyTheme(nextTheme)
    persistThemePreference(nextTheme)
  }

  function saveState() {
    if (state.selectedCasePath) {
      writeWorkspaceStorageItem('casePath', state.selectedCasePath)
    } else {
      removeWorkspaceStorageItem('casePath')
    }

    if (state.selectedScopePath) {
      writeWorkspaceStorageItem('scopePath', state.selectedScopePath)
    } else {
      removeWorkspaceStorageItem('scopePath')
    }

    if (state.selectedPstFileName) {
      writeWorkspaceStorageItem('pstFileName', state.selectedPstFileName)
    } else {
      removeWorkspaceStorageItem('pstFileName')
    }

    if (state.currentFolderId) {
      writeWorkspaceStorageItem('folderId', state.currentFolderId)
    } else {
      removeWorkspaceStorageItem('folderId')
    }

    if (state.selectedMessageId) {
      writeWorkspaceStorageItem('messageId', state.selectedMessageId)
    } else {
      removeWorkspaceStorageItem('messageId')
    }

    writeWorkspaceStorageItem('query', state.query)
    writeWorkspaceStorageItem('mailOnly', state.mailOnly ? '1' : '0')
    writeWorkspaceStorageItem('sort', state.sort)
    writeWorkspaceStorageItem(
      'reviewFlaggedOnly',
      state.reviewFlaggedOnly ? '1' : '0'
    )
    writeWorkspaceStorageItem(
      'reviewTaggedOnly',
      state.reviewTaggedOnly ? '1' : '0'
    )
    writeWorkspaceStorageItem('searchScope', state.searchScope)
    writeWorkspaceStorageItem('mailboxScopeView', state.mailboxScopeView)
    writeWorkspaceStorageItem('mailboxFilter', state.mailboxFilter)
    writeWorkspaceStorageItem('catalogMode', state.catalogMode)
    writeWorkspaceStorageItem('flaggedBundleScope', state.bundleScope)
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
    ui.hiddenFiltersPanel.setAttribute('aria-hidden', isOpen ? 'false' : 'true')
    ui.hiddenFiltersPanel.style.display = isOpen ? 'flex' : 'none'

    if (!isOpen) {
      return
    }

    const closeButtons = ui.hiddenFiltersPanel.querySelectorAll('[data-action="close-hidden-filters"]')
    closeButtons.forEach((button) => {
      button.onclick = (event) => {
        event.preventDefault()
        event.stopPropagation()
        closeHiddenFiltersDropdown()
      }
    })

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

  function renderPstRow(file) {
    const modifiedAt = file.modifiedAt ? formatDate(file.modifiedAt) : 'Unknown date'
    const size = Number.isFinite(file.size) ? formatBytes(file.size) : 'Unknown size'
    const pathLine =
      state.mailboxScopeView === 'all' && file.displayPath
        ? `<span class="pst-item-path">${escapeHtml(file.displayPath)}</span>`
        : ''
    const rowTitle = `
      <span class="pst-item-name">${escapeHtml(file.fileName)}</span>
      ${state.catalogMode === 'removed'
        ? `<button
            class="ghost-button small pst-item-action pst-item-action--compact pst-item-action--inline"
            type="button"
            data-action="restore-pst"
            data-pst-file-name="${escapeAttr(file.fileName)}"
            data-scope-path="${escapeAttr(file.scopePath || '')}"
            aria-label="Restore PST to active catalog"
            title="Restore PST to active catalog"
          >
            +
          </button>`
        : `<button
            class="ghost-button small pst-item-action pst-item-action--compact pst-item-action--inline"
            type="button"
            data-action="remove-pst"
            data-pst-file-name="${escapeAttr(file.fileName)}"
            data-scope-path="${escapeAttr(file.scopePath || '')}"
            aria-label="Remove PST from platform"
            title="Remove PST from platform"
          >
            -
          </button>`}
    `
    const rowClasses =
      state.catalogMode === 'removed'
        ? 'pst-item pst-item-static pst-item-inline'
        : `pst-item pst-item-clickable${file.fileName === state.selectedPstFileName &&
          normalizeScopePath(file.scopePath) === normalizeScopePath(state.selectedScopePath)
            ? ' active'
            : ''}`

    if (state.catalogMode === 'removed') {
      return `
        <div class="pst-item-row">
          <div
            class="${rowClasses}"
            title="${escapeAttr(file.fileName)}"
            data-pst-file-name="${escapeAttr(file.fileName)}"
            data-scope-path="${escapeAttr(file.scopePath || '')}"
          >
            <div class="pst-item-title-line">
              ${rowTitle}
            </div>
            ${pathLine}
            <span class="pst-item-meta">${escapeHtml(size)} · ${escapeHtml(modifiedAt)}</span>
          </div>
        </div>
      `
    }

    return `
      <div class="pst-item-row">
        <div
          class="${rowClasses}"
          role="button"
          tabindex="0"
          title="${escapeAttr(file.fileName)}"
          data-pst-file-name="${escapeAttr(file.fileName)}"
          data-scope-path="${escapeAttr(file.scopePath || '')}"
        >
          <div class="pst-item-title-line">
            ${rowTitle}
          </div>
          ${pathLine}
          <span class="pst-item-meta">${escapeHtml(size)} · ${escapeHtml(modifiedAt)}</span>
        </div>
      </div>
    `
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

    ui.mailboxesTitle.textContent = getCatalogModeLabel(state.catalogMode)
    ui.catalogModeToggle.textContent =
      state.catalogMode === 'removed' ? 'Active' : 'Removed'
    ui.catalogModeToggle.setAttribute('aria-pressed', state.catalogMode === 'removed' ? 'true' : 'false')

    const allMailboxCount = selectedCase
      ? getMailboxEntriesForCase(selectedCase.casePath, 'all').length
      : 0
    const visibleMailboxes = selectedCase
      ? getMailboxEntriesForDisplay(selectedCase.casePath, state.mailboxScopeView)
      : []
    const selectedMailboxVisible =
      state.catalogMode === 'active' &&
      visibleMailboxes.some(
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
        `<strong>Loading ${
          state.catalogMode === 'removed' ? 'removed PST files' : 'PST files'
        }...</strong> Scanning the project <code>PST/</code> folder.`
      ui.pstList.innerHTML = ''
      return
    }

    if (!searches.length) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML = escapeHtml(
        state.catalogMessage ||
          (state.catalogMode === 'removed'
            ? 'No removed PSTs are available yet.'
            : 'No searches available.')
      )
      ui.pstList.innerHTML = ''
      return
    }

    if (!visibleMailboxes.length) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML = escapeHtml(
        state.mailboxFilter
          ? 'No PST files match the current mailbox filter.'
          : state.catalogMode === 'removed'
            ? state.mailboxScopeView === 'all'
              ? 'No removed PST files were found for this case.'
              : state.catalogMessage || 'No removed PST files found.'
            : state.mailboxScopeView === 'all'
              ? 'No PST files were found for this case.'
              : state.catalogMessage || 'No PST files found.'
      )
      ui.pstList.innerHTML = ''
      return
    }

    ui.pstEmpty.classList.add('hidden')
    ui.pstList.innerHTML = visibleMailboxes.map((file) => renderPstRow(file)).join('')
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

    const visibleItems = filterVisibleMessageItems(page.items)
    const selectedVisible =
      state.selectedMessageId &&
      visibleItems.some((item) => resolveMessageId(item) === state.selectedMessageId)
    if (state.selectedMessageId && !selectedVisible) {
      state.selectedMessageId = null
      state.currentMessageDetail = null
      ui.messageDetail.innerHTML =
        '<div class="panel-empty">Select a message to inspect it.</div>'
      saveState()
    }

    ui.messageCountBadge.textContent = String(state.summary?.stats?.messageCount ?? 0)
    ui.messageList.innerHTML = visibleItems.length
      ? visibleItems.map((item) => renderMessageRow(item)).join('')
      : '<div class="panel-empty">No messages match the current filters.</div>'

    const folderName = page.folder?.displayName || page.scopeLabel || 'folder'
    const queryLabel = page.query ? ` filtered by "${page.query}"` : ''
    ui.messageResultCount.textContent = visibleItems.length
      ? `Showing ${visibleItems.length} of ${page.total} messages in ${folderName}${queryLabel}.`
      : `No messages found in ${folderName}${queryLabel}.`
    ui.pageInfo.textContent = `Page ${page.page} of ${page.totalPages}`
    updatePagingButtons()
  }

  function buildHtmlFrame(html) {
    const isDark = state.theme === 'dark'
    const frameText = isDark ? '#e6edf7' : '#1f2a37'
    const frameBackground = isDark ? '#111a29' : '#ffffff'
    const frameLink = isDark ? '#7ab8ff' : '#2b6cb0'
    const frameQuoteBorder = isDark ? '#3a4760' : '#c2ccd9'
    const frameQuoteText = isDark ? '#a8b3c4' : '#526072'

    const srcdoc = `<!doctype html>
      <html>
        <head>
          <meta charset="utf-8" />
          <base target="_blank" />
          <style>
            :root {
              color-scheme: ${isDark ? 'dark' : 'light'};
            }
            body {
              margin: 0;
              padding: 8px 4px 12px;
              font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
              font-size: 13px;
              line-height: 1.5;
              color: ${frameText};
              background: ${frameBackground};
            }
            img { max-width: 100%; height: auto; }
            table { border-collapse: collapse; }
            a { color: ${frameLink}; }
            blockquote {
              border-left: 3px solid ${frameQuoteBorder};
              margin-left: 0;
              padding-left: 1rem;
              color: ${frameQuoteText};
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

    const isDownloadable = attachment.isDownloadable !== false && Boolean(attachment.downloadUrl)
    const downloadAction = isDownloadable
      ? `<a class="attachment-link" href="${escapeAttr(attachment.downloadUrl)}" download="${escapeAttr(
          name
        )}">Download</a>`
      : `<span class="attachment-unavailable" aria-disabled="true" title="Attachment bytes are not stored in this PST.">Unavailable</span>`

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
        ${downloadAction}
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
    const requestUrl = String(url || '')
    const { headers: initHeaders, ...requestInit } = init
    const response = await fetch(requestUrl, {
      ...requestInit,
      credentials: 'include',
      headers: {
        Accept: 'application/json',
        ...(initHeaders || {})
      }
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
      const requestPath = requestUrl.split('?')[0]
      if (
        response.status === 401 &&
        !['/api/auth/login', '/api/auth/me', '/api/auth/logout'].includes(requestPath) &&
        state.authEnabled !== false
      ) {
        handleAuthFailure('Session expired. Sign in again.')
      }
      const error = new Error(message)
      error.statusCode = response.status
      error.payload = payload
      throw error
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
      const restoreFolderId = options.restoreFolderId || readWorkspaceStorageItem('folderId', '')
      const restoreMessageId =
        options.restoreMessageId || readWorkspaceStorageItem('messageId', '')
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
    const shouldAutoOpen = state.catalogMode === 'active' && options.refreshOnly !== true
    if (showBusy) {
      setBodyBusy(true)
    }
    setStatus('Loading PST catalog...')
    try {
      const casePath = normalizeScopePath(
        options.casePath ||
          state.selectedCasePath ||
          readWorkspaceStorageItem('casePath', '') ||
          ''
      )
      const scopePath = normalizeScopePath(
        options.scopePath ||
          state.selectedScopePath ||
          readWorkspaceStorageItem('scopePath', '') ||
          ''
      )
      const response = await fetchJson(
        `${getCatalogEndpoint()}${scopePath ? `?scopePath=${encodeURIComponent(scopePath)}` : ''}`
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
        readWorkspaceStorageItem('pstFileName', '') ||
        ''
      const preferredScopePath = normalizeScopePath(
        options.preferredScopePath ||
          state.selectedScopePath ||
          readWorkspaceStorageItem('scopePath', '') ||
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

      if (!shouldAutoOpen) {
        renderPstCatalog()
        setStatus(response.message || 'Catalog loaded.', 'success')
        return
      }

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
        restoreFolderId: readWorkspaceStorageItem('folderId', '') || undefined,
        restoreMessageId: readWorkspaceStorageItem('messageId', '') || undefined
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
        options.preferredMessageId || readWorkspaceStorageItem('messageId', '') || null
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
        options.preferredMessageId || readWorkspaceStorageItem('messageId', '') || null
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

  async function initializeViewer() {
    if (state.viewerInitialized) {
      refreshControls()
      return
    }

    state.viewerInitialized = true
    try {
      refreshControls()
      wireEvents()
      renderPstCatalog()
      await loadMailboxCatalog({
        showBusy: true,
        refreshOnly: state.catalogMode === 'removed',
        casePath: state.selectedCasePath || undefined,
        scopePath: state.selectedScopePath || undefined,
        preferredFileName: state.selectedPstFileName || undefined
      })
      await loadHiddenFilters()
    } catch (error) {
      state.viewerInitialized = false
      throw error
    }
  }

  async function loadAuthState() {
    try {
      const response = await fetchJson('/api/auth/me', {
        cache: 'no-store'
      })
      state.authEnabled = Boolean(response.enabled)
      state.authChecked = true
      if (!response.enabled) {
        state.authUser = 'Local access'
        await beginAuthenticatedWorkspace(state.authUser, false, false, {
          suppressReminder: true
        })
        return true
      }

      if (response.authenticated) {
        state.authUser = (response.user && response.user.username) || state.authUser || ''
        await beginAuthenticatedWorkspace(
          state.authUser,
          Boolean(response.canManageUsers),
          Boolean(response.mfaEnabled)
        )
        return true
      }

      if (response.mfaRequired && response.user && response.user.username) {
        state.authUser = ''
        showMfaChallengeScreen(response.user.username)
        return false
      }

      state.authUser = ''
      if (state.authInviteLoaded && state.authInvite) {
        showInviteScreen(state.authInvite)
        return false
      }

      showLoginScreen()
      return false
    } catch (error) {
      state.authChecked = true
      state.authEnabled = true
      if (error && typeof error === 'object' && error.statusCode === 401) {
        if (state.authInviteLoaded && state.authInvite) {
          showInviteScreen(state.authInvite)
        } else {
          showLoginScreen()
        }
        return false
      }
      if (state.authInviteLoaded && state.authInvite) {
        showInviteScreen(state.authInvite)
      } else {
        showLoginScreen()
      }
      return false
    }
  }

  async function loadInviteDetails(inviteToken) {
    const token = String(inviteToken || '').trim()
    if (!token) {
      return false
    }

    state.authInviteToken = token
    setAuthBusy(true)
    setAuthMessage('Loading invite...')
    setAuthView('invite')
    try {
      const response = await fetchJson(`/api/auth/invites/${encodeURIComponent(token)}`, {
        cache: 'no-store'
      })
      const invite = response && response.invite ? response.invite : null
      if (!invite || !invite.username) {
        throw new Error('Invite not found.')
      }

      showInviteScreen(invite)
      setAuthMessage('Invite validated. Choose a password to continue.', 'success')
      return true
    } catch (error) {
      showLoginScreen()
      setAuthError(getErrorMessage(error))
      return false
    } finally {
      setAuthBusy(false)
    }
  }

  function wireAuthEvents() {
    if (ui.authForm) {
      ui.authForm.addEventListener('submit', (event) => {
        void (async () => {
          event.preventDefault()
          setAuthError('')
          const username = String(ui.authUsername?.value || '').trim()
          const password = String(ui.authPassword?.value || '')
          if (!username || !password) {
            setAuthError('Enter both a username and password.')
            return
          }

          setAuthBusy(true)
          try {
            const response = await fetchJson('/api/auth/login', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({
                username,
                password
              })
            })

            if (response.mfaRequired) {
              const challengeUsername =
                (response.user && response.user.username) || username || state.authMfaChallengeUsername
              if (!challengeUsername) {
                throw new Error('Verification required.')
              }
              showMfaChallengeScreen(challengeUsername)
              setAuthMessage('Enter your verification code to finish signing in.', 'success')
              return
            }

            if (!response.authenticated) {
              throw new Error('Unable to sign in.')
            }

            await beginAuthenticatedWorkspace(
              (response.user && response.user.username) || username,
              Boolean(response.canManageUsers),
              Boolean(response.mfaEnabled)
            )
          } catch (error) {
            setAuthError(getErrorMessage(error))
          } finally {
            setAuthBusy(false)
          }
        })()
      })
    }

    if (ui.authMfaForm) {
      ui.authMfaForm.addEventListener('submit', (event) => {
        void (async () => {
          event.preventDefault()
          setAuthError('')
          const code = String(ui.authMfaCode?.value || '').trim()
          if (!code) {
            setAuthError('Enter a verification code.')
            return
          }

          setAuthBusy(true)
          try {
            const response = await fetchJson('/api/auth/mfa/challenge', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({ code })
            })

            if (!response.authenticated) {
              throw new Error('Unable to verify the code.')
            }

            await beginAuthenticatedWorkspace(
              (response.user && response.user.username) || state.authMfaChallengeUsername || state.authUser,
              Boolean(response.canManageUsers),
              Boolean(response.mfaEnabled)
            )
          } catch (error) {
            setAuthError(getErrorMessage(error))
          } finally {
            setAuthBusy(false)
          }
        })()
      })
    }

    if (ui.inviteForm) {
      ui.inviteForm.addEventListener('submit', (event) => {
        void (async () => {
          event.preventDefault()
          await submitInvitePassword()
        })()
      })
    }

    if (ui.inviteMfaStart) {
      ui.inviteMfaStart.addEventListener('click', () => {
        void startInviteMfaSetup()
      })
    }

    if (ui.inviteMfaSkip) {
      ui.inviteMfaSkip.addEventListener('click', () => {
        void (async () => {
          if (!state.authInviteUsername) {
            return
          }
          dismissMfaReminder(state.authInviteUsername)
          await finalizeInviteOnboarding(state.authInviteUsername, false, false, {
            suppressReminder: true
          })
        })()
      })
    }

    if (ui.inviteMfaForm) {
      ui.inviteMfaForm.addEventListener('submit', (event) => {
        void (async () => {
          event.preventDefault()
          await completeInviteMfaSetup()
        })()
      })
    }

    if (ui.inviteMfaDownload) {
      ui.inviteMfaDownload.addEventListener('click', () => {
        downloadInviteRecoveryCodes()
      })
    }

    if (ui.inviteFinish) {
      ui.inviteFinish.addEventListener('click', () => {
        void finalizeInviteOnboarding(state.authInviteUsername, false, true, {
          suppressReminder: true
        })
      })
    }

    if (ui.authLogout) {
      ui.authLogout.addEventListener('click', () => {
        void (async () => {
          setAuthBusy(true)
          try {
            await fetchJson('/api/auth/logout', {
              method: 'POST'
            })
            clearMfaReminderDismissal(state.authUser)
            window.location.reload()
          } catch (error) {
            setAuthError(getErrorMessage(error))
          } finally {
            setAuthBusy(false)
          }
        })()
      })
    }
  }

  function wireUserManagementEvents() {
    if (ui.settingsButton) {
      ui.settingsButton.addEventListener('click', (event) => {
        event.preventDefault()
        toggleSettingsMenu()
      })
    }

    if (ui.setupMfaButton) {
      ui.setupMfaButton.addEventListener('click', (event) => {
        event.preventDefault()
        void startSelfServiceMfaSetup()
      })
    }

    if (ui.manageUsersButton) {
      ui.manageUsersButton.addEventListener('click', (event) => {
        event.preventDefault()
        openUserManagementModal()
      })
    }

    if (ui.smtpSettingsButton) {
      ui.smtpSettingsButton.addEventListener('click', (event) => {
        event.preventDefault()
        openSmtpSettingsModal()
      })
    }

    if (ui.activityLogButton) {
      ui.activityLogButton.addEventListener('click', (event) => {
        event.preventDefault()
        openActivityLogModal()
      })
    }

    if (ui.mfaReminderSetup) {
      ui.mfaReminderSetup.addEventListener('click', () => {
        void startSelfServiceMfaSetup({
          suppressReminder: true
        })
      })
    }

    if (ui.mfaReminderSkip) {
      ui.mfaReminderSkip.addEventListener('click', () => {
        void (async () => {
          const username = String(state.authMfaReminderUsername || state.authUser || '').trim()
          if (username) {
            dismissMfaReminder(username)
          }
          state.authMfaReminderDeferred = false
          closeMfaReminderModal()
          if (!state.viewerInitialized) {
            try {
              await initializeViewer()
            } catch (error) {
              if (!(error && typeof error === 'object' && Number(error.statusCode) === 401)) {
                setStatus(`Signed in, but unable to load the viewer: ${getErrorMessage(error)}`, 'error')
              }
            }
          }
        })()
      })
    }

    if (ui.mfaSetupClose) {
      ui.mfaSetupClose.addEventListener('click', () => {
        closeMfaSetupModal()
      })
    }

    if (ui.mfaSetupModal) {
      ui.mfaSetupModal.addEventListener('click', (event) => {
        const target = event.target
        if (!(target instanceof Element)) {
          return
        }
        if (target === ui.mfaSetupModal || target.getAttribute('data-action') === 'close-mfa-setup') {
          closeMfaSetupModal()
        }
      })
    }

    if (ui.mfaSetupForm) {
      ui.mfaSetupForm.addEventListener('submit', (event) => {
        event.preventDefault()
        void completeSelfServiceMfaSetup()
      })
    }

    if (ui.mfaSetupDownload) {
      ui.mfaSetupDownload.addEventListener('click', () => {
        downloadMfaSetupRecoveryCodes()
      })
    }

    if (ui.mfaSetupFinish) {
      ui.mfaSetupFinish.addEventListener('click', () => {
        void continueAfterMfaSetup()
      })
    }

    if (ui.userManagementClose) {
      ui.userManagementClose.addEventListener('click', () => {
        closeUserManagementModal()
      })
    }

    if (ui.userManagementModal) {
      ui.userManagementModal.addEventListener('click', (event) => {
        const target = event.target
        if (!(target instanceof Element)) {
          return
        }
        if (
          target === ui.userManagementModal ||
          target.getAttribute('data-action') === 'close-user-management'
        ) {
          closeUserManagementModal()
        }
      })
    }

    if (ui.smtpSettingsClose) {
      ui.smtpSettingsClose.addEventListener('click', () => {
        closeSmtpSettingsModal()
      })
    }

    if (ui.smtpSettingsModal) {
      ui.smtpSettingsModal.addEventListener('click', (event) => {
        const target = event.target
        if (!(target instanceof Element)) {
          return
        }
        if (target === ui.smtpSettingsModal || target.getAttribute('data-action') === 'close-smtp-settings') {
          closeSmtpSettingsModal()
        }
      })
    }

    if (ui.smtpSettingsForm) {
      ui.smtpSettingsForm.addEventListener('submit', (event) => {
        event.preventDefault()
        void saveSmtpSettings()
      })
    }

    if (ui.smtpTestSend) {
      ui.smtpTestSend.addEventListener('click', (event) => {
        event.preventDefault()
        void sendSmtpTestEmail()
      })
    }

    if (ui.userList) {
      ui.userList.addEventListener('click', (event) => {
        const target = event.target
        if (!(target instanceof Element)) {
          return
        }

        const actionButton = target.closest(
          '[data-action="select-user-activity"], [data-action="delete-user"], [data-action="resend-invite"], [data-action="revoke-invite"], [data-action="reset-mfa"]'
        )
        if (!actionButton || !ui.userList.contains(actionButton)) {
          return
        }

        const action = actionButton.getAttribute('data-action')
        const username = String(actionButton.getAttribute('data-username') || '').trim()
        if (!username) {
          return
        }

        event.preventDefault()
        if (action === 'select-user-activity') {
          selectUserActivity(username)
          return
        }
        if (action === 'delete-user') {
          void deleteUser(username)
          return
        }
        if (action === 'resend-invite') {
          void resendInvite(username)
          return
        }
        if (action === 'revoke-invite') {
          void revokeInvite(username)
          return
        }
        if (action === 'reset-mfa') {
          void resetUserMfa(username)
        }
      })
    }

    if (ui.userManagementForm) {
      ui.userManagementForm.addEventListener('submit', (event) => {
        event.preventDefault()
        void inviteUser()
      })
    }

    if (ui.userActivityRefresh) {
      ui.userActivityRefresh.addEventListener('click', () => {
        const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
        if (!selectedUsername) {
          return
        }
        void loadUserActivityLog(selectedUsername, {
          showBusy: true,
          silent: false
        })
      })
    }

    if (ui.userActivityExport) {
      ui.userActivityExport.addEventListener('click', () => {
        const selectedUsername = String(state.selectedUserActivityUsername || '').trim()
        if (!selectedUsername) {
          setUserActivityMessage('Select a user before exporting.', 'error')
          return
        }
        downloadActivityLogCsv(selectedUsername)
      })
    }

    if (ui.activityLogClose) {
      ui.activityLogClose.addEventListener('click', () => {
        closeActivityLogModal()
      })
    }

    if (ui.activityLogExport) {
      ui.activityLogExport.addEventListener('click', () => {
        downloadActivityLogCsv()
      })
    }

    if (ui.activityLogRefresh) {
      ui.activityLogRefresh.addEventListener('click', () => {
        void loadActivityLog({
          showBusy: true,
          silent: false
        })
      })
    }

    if (ui.activityLogModal) {
      ui.activityLogModal.addEventListener('click', (event) => {
        const target = event.target
        if (!(target instanceof Element)) {
          return
        }
        if (
          target === ui.activityLogModal ||
          target.getAttribute('data-action') === 'close-activity-log'
        ) {
          closeActivityLogModal()
        }
      })
    }

    document.addEventListener('click', (event) => {
      if (!state.settingsMenuOpen) {
        return
      }

      const target = event.target
      if (!(target instanceof Node)) {
        return
      }

      if (ui.settingsMenu && ui.settingsMenu.contains(target)) {
        return
      }

      if (ui.settingsButton && ui.settingsButton.contains(target)) {
        return
      }

      closeSettingsMenu()
    })

    document.addEventListener('keydown', (event) => {
      if (event.key !== 'Escape') {
        return
      }

      if (ui.userManagementModal && !ui.userManagementModal.hidden) {
        closeUserManagementModal()
        return
      }

      if (ui.smtpSettingsModal && !ui.smtpSettingsModal.hidden) {
        closeSmtpSettingsModal()
        return
      }

      if (ui.activityLogModal && !ui.activityLogModal.hidden) {
        closeActivityLogModal()
        return
      }

      if (ui.mfaSetupModal && !ui.mfaSetupModal.hidden) {
        closeMfaSetupModal()
        return
      }

      if (state.settingsMenuOpen) {
        closeSettingsMenu()
      }
    })

  }

  function wireThemeEvents() {
    for (const button of Array.isArray(ui.themeToggleButtons) ? ui.themeToggleButtons : []) {
      button.addEventListener('click', () => {
        toggleTheme()
      })
    }
  }

  function wireEvents() {
    ui.pstCountBadge.addEventListener('click', () => {
      void (async () => {
        setStatus('Refreshing mailbox catalog and search cache...')
        try {
          await refreshSearchIndex()
        } catch (error) {
          setStatus(`Unable to refresh search cache: ${error.message}`, 'error')
        }
        await loadMailboxCatalog({
          showBusy: true,
          refreshOnly: state.catalogMode === 'removed',
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
        refreshOnly: state.catalogMode === 'removed',
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
        refreshOnly: state.catalogMode === 'removed',
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
      const actionButton = event.target.closest('[data-action]')
      if (actionButton && actionButton.dataset.action) {
        const action = actionButton.dataset.action
        const fileName = actionButton.dataset.pstFileName || ''
        const scopePath = actionButton.dataset.scopePath || ''
        if (action === 'remove-pst' && fileName) {
          void removeMailboxFromPlatform(fileName, scopePath)
        } else if (action === 'restore-pst' && fileName) {
          void restoreMailboxToPlatform(fileName, scopePath)
        }
        return
      }

      const button = event.target.closest('[data-pst-file-name]')
      if (!button) {
        return
      }
      if (state.catalogMode === 'removed') {
        return
      }
      const fileName = button.dataset.pstFileName
      const scopePath = normalizeScopePath(button.dataset.scopePath || '')
      if (!fileName) {
        return
      }
      void openMailbox(fileName, { showBusy: true, scopePath })
    })

    ui.pstList.addEventListener('keydown', (event) => {
      const target = event.target.closest('[data-pst-file-name]')
      if (!target || state.catalogMode === 'removed') {
        return
      }
      if (event.key !== 'Enter' && event.key !== ' ') {
        return
      }
      if (event.target.closest('[data-action]')) {
        return
      }
      event.preventDefault()
      const fileName = target.dataset.pstFileName
      const scopePath = normalizeScopePath(target.dataset.scopePath || '')
      if (!fileName) {
        return
      }
      void openMailbox(fileName, { showBusy: true, scopePath })
    })

    if (ui.catalogModeToggle) {
      ui.catalogModeToggle.addEventListener('click', () => {
        state.catalogMode = state.catalogMode === 'removed' ? 'active' : 'removed'
        saveState()
        void loadMailboxCatalog({
          showBusy: true,
          refreshOnly: true,
          casePath: state.selectedCasePath || undefined,
          scopePath: state.selectedScopePath || undefined,
          preferredFileName: state.selectedPstFileName || undefined,
          preferredScopePath: state.selectedScopePath || undefined
        })
      })
    }

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
      saveState()
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
    ui.authScreen = getElement('auth-screen')
    ui.appShell = getElement('app-shell')
    ui.authMessage = getElement('auth-message')
    ui.authLoginView = getElement('auth-login-view')
    ui.authMfaView = getElement('auth-mfa-view')
    ui.authForm = getElement('auth-form')
    ui.authUsername = getElement('auth-username')
    ui.authPassword = getElement('auth-password')
    ui.authSubmit = getElement('auth-submit')
    ui.authMfaForm = getElement('auth-mfa-form')
    ui.authMfaCode = getElement('auth-mfa-code')
    ui.authMfaSubmit = getElement('auth-mfa-submit')
    ui.authMfaDescription = getElement('auth-mfa-description')
    ui.inviteView = getElement('auth-invite-view')
    ui.inviteDetails = getElement('invite-details')
    ui.inviteSummary = getElement('invite-summary')
    ui.inviteForm = getElement('invite-form')
    ui.invitePassword = getElement('invite-password')
    ui.inviteConfirmPassword = getElement('invite-confirm-password')
    ui.inviteSubmit = getElement('invite-submit')
    ui.inviteMfaPrompt = getElement('invite-mfa-prompt')
    ui.inviteMfaStart = getElement('invite-mfa-start')
    ui.inviteMfaSkip = getElement('invite-mfa-skip')
    ui.inviteMfaSetup = getElement('invite-mfa-setup')
    ui.inviteMfaQr = getElement('invite-mfa-qr')
    ui.inviteMfaSecret = getElement('invite-mfa-secret')
    ui.inviteMfaUri = getElement('invite-mfa-uri')
    ui.inviteMfaForm = getElement('invite-mfa-form')
    ui.inviteMfaCode = getElement('invite-mfa-code')
    ui.inviteMfaSubmit = getElement('invite-mfa-submit')
    ui.inviteMfaComplete = getElement('invite-mfa-complete')
    ui.inviteMfaRecoveryList = getElement('invite-mfa-recovery-list')
    ui.inviteMfaDownload = getElement('invite-mfa-download')
    ui.inviteFinish = getElement('invite-finish')
    ui.authLogout = getElement('auth-logout')
    ui.authUser = getElement('auth-user')
    ui.settingsButton = getElement('settings-button')
    ui.settingsMenu = getElement('settings-menu')
    ui.setupMfaButton = getElement('setup-mfa-button')
    ui.manageUsersButton = getElement('manage-users-button')
    ui.smtpSettingsButton = getElement('smtp-settings-button')
    ui.activityLogButton = getElement('activity-log-button')
    ui.mfaReminderModal = getElement('mfa-reminder-modal')
    ui.mfaReminderSetup = getElement('mfa-reminder-setup')
    ui.mfaReminderSkip = getElement('mfa-reminder-skip')
    ui.mfaSetupModal = getElement('mfa-setup-modal')
    ui.mfaSetupClose = getElement('mfa-setup-close')
    ui.mfaSetupMessage = getElement('mfa-setup-message')
    ui.mfaSetupQr = getElement('mfa-setup-qr')
    ui.mfaSetupSecret = getElement('mfa-setup-secret')
    ui.mfaSetupUri = getElement('mfa-setup-uri')
    ui.mfaSetupForm = getElement('mfa-setup-form')
    ui.mfaSetupCode = getElement('mfa-setup-code')
    ui.mfaSetupSubmit = getElement('mfa-setup-submit')
    ui.mfaSetupComplete = getElement('mfa-setup-complete')
    ui.mfaSetupRecoveryList = getElement('mfa-setup-recovery-list')
    ui.mfaSetupDownload = getElement('mfa-setup-download')
    ui.mfaSetupFinish = getElement('mfa-setup-finish')
    ui.userManagementModal = getElement('user-management-modal')
    ui.userManagementClose = getElement('user-management-close')
    ui.userManagementBackdrop = ui.userManagementModal.querySelector('[data-action="close-user-management"]')
    ui.userManagementForm = getElement('user-management-form')
    ui.userManagementUsername = getElement('user-management-username')
    ui.userManagementEmail = getElement('user-management-email')
    ui.userManagementSubmit = getElement('user-management-submit')
    ui.userManagementMessage = getElement('user-management-message')
    ui.userList = getElement('user-list')
    ui.userCountBadge = getElement('user-count-badge')
    ui.userActivitySelected = getElement('user-activity-selected')
    ui.userActivityMessage = getElement('user-activity-message')
    ui.userActivityList = getElement('user-activity-list')
    ui.userActivityCountBadge = getElement('user-activity-count-badge')
    ui.userActivityExport = getElement('user-activity-export')
    ui.userActivityRefresh = getElement('user-activity-refresh')
    ui.userActivityTitle = getElement('user-activity-title')
    ui.smtpSettingsModal = getElement('smtp-settings-modal')
    ui.smtpSettingsClose = getElement('smtp-settings-close')
    ui.smtpSettingsBackdrop = ui.smtpSettingsModal.querySelector('[data-action="close-smtp-settings"]')
    ui.smtpSettingsForm = getElement('smtp-settings-form')
    ui.smtpSettingsMessage = getElement('smtp-settings-message')
    ui.smtpSettingsEnabled = getElement('smtp-settings-enabled')
    ui.smtpSettingsHost = getElement('smtp-settings-host')
    ui.smtpSettingsPort = getElement('smtp-settings-port')
    ui.smtpSettingsSecure = getElement('smtp-settings-secure')
    ui.smtpSettingsUsername = getElement('smtp-settings-username')
    ui.smtpSettingsPassword = getElement('smtp-settings-password')
    ui.smtpPasswordHint = getElement('smtp-password-hint')
    ui.smtpSettingsFromName = getElement('smtp-settings-from-name')
    ui.smtpSettingsFromAddress = getElement('smtp-settings-from-address')
    ui.smtpSettingsReplyTo = getElement('smtp-settings-reply-to')
    ui.smtpTestRecipient = getElement('smtp-test-recipient')
    ui.smtpTestSend = getElement('smtp-test-send')
    ui.smtpSettingsSubmit = getElement('smtp-settings-submit')
    ui.activityLogModal = getElement('activity-log-modal')
    ui.activityLogClose = getElement('activity-log-close')
    ui.activityLogBackdrop = ui.activityLogModal.querySelector('[data-action="close-activity-log"]')
    ui.activityLogExport = getElement('activity-log-export')
    ui.activityLogRefresh = getElement('activity-log-refresh')
    ui.activityLogMessage = getElement('activity-log-message')
    ui.activityLogList = getElement('activity-log-list')
    ui.activityLogCountBadge = getElement('activity-log-count-badge')
    ui.sessionSummary = getElement('session-summary')
    ui.pstCountBadge = getElement('pst-count-badge')
    ui.catalogModeToggle = getElement('catalog-mode-toggle')
    ui.mailboxesTitle = getElement('mailboxes-section-title')
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
    ui.themeToggleButtons = Array.from(document.querySelectorAll('[data-theme-toggle]'))

    applyTheme('dark')
    state.hiddenFiltersOpen = false
    state.mailboxFilter = ''
    wireThemeEvents()
    wireAuthEvents()
    wireUserManagementEvents()
    const inviteToken = getInviteTokenFromLocation()
    if (inviteToken) {
      await loadInviteDetails(inviteToken)
      return
    }

    const authenticated = await loadAuthState()
    if (!authenticated) {
      return
    }
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
