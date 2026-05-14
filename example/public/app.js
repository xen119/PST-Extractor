(function () {
  const STORAGE_KEYS = {
    pstFileName: 'pst-mail-explorer.pstFileName',
    folderId: 'pst-mail-explorer.folderId',
    messageId: 'pst-mail-explorer.messageId',
    query: 'pst-mail-explorer.query',
    mailOnly: 'pst-mail-explorer.mailOnly',
    sort: 'pst-mail-explorer.sort',
    reviewFlaggedOnly: 'pst-mail-explorer.reviewFlaggedOnly',
    reviewTaggedOnly: 'pst-mail-explorer.reviewTaggedOnly'
  }

  const state = {
    sessionId: null,
    catalog: [],
    catalogLoaded: false,
    catalogMessage: '',
    selectedPstFileName: null,
    summary: null,
    tree: null,
    folderMap: new Map(),
    currentFolderId: null,
    currentFolderPage: null,
    currentMessageDetail: null,
    selectedMessageId: null,
    query: '',
    mailOnly: true,
    sort: 'date-desc',
    reviewFlaggedOnly: false,
    reviewTaggedOnly: false,
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
    if (!detail.isMailLike) {
      return `
        <section class="review-panel disabled">
          <div class="review-note">Review controls are available for mail items only.</div>
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
        </div>
        <div class="review-tags">${tagChips}</div>
        <form class="review-tag-form" data-action="add-review-tag">
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
          item.id === messageId
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
    if (!state.currentMessageDetail || !state.currentMessageDetail.isMailLike) {
      return
    }

    const review = normalizeReviewState(state.currentMessageDetail.review)
    await saveReviewState(state.currentMessageDetail.id, {
      flagged: !review.flagged
    })
  }

  async function clearReview() {
    if (!state.currentMessageDetail || !state.currentMessageDetail.isMailLike) {
      return
    }
    await saveReviewState(state.currentMessageDetail.id, {}, { deleteReview: true })
  }

  async function addReviewTag(tag) {
    if (!state.currentMessageDetail || !state.currentMessageDetail.isMailLike) {
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
    if (!state.currentMessageDetail || !state.currentMessageDetail.isMailLike) {
      return
    }

    const review = normalizeReviewState(state.currentMessageDetail.review)
    const nextTags = review.tags.filter((value) => value.toLowerCase() !== String(tag).toLowerCase())
    await saveReviewState(state.currentMessageDetail.id, {
      tags: nextTags
    })
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
  }

  function applyStateToControls() {
    ui.searchInput.value = state.query
    ui.mailOnlyToggle.checked = state.mailOnly
    ui.sortSelect.value = state.sort
    ui.reviewFlaggedToggle.checked = state.reviewFlaggedOnly
    ui.reviewTaggedToggle.checked = state.reviewTaggedOnly
  }

  function resetSessionState(message = 'Select a PST file from the list on the left.') {
    state.sessionId = null
    state.summary = null
    state.tree = null
    state.folderMap = new Map()
    state.currentFolderId = null
    state.currentFolderPage = null
    state.currentMessageDetail = null
    state.selectedMessageId = null
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

  function renderPstCatalog() {
    const files = Array.isArray(state.catalog) ? state.catalog : []
    ui.pstCountBadge.textContent = String(files.length)

    if (!state.catalogLoaded) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML =
        '<strong>Loading PST files...</strong> Scanning the project <code>PST/</code> folder.'
      ui.pstList.innerHTML = ''
      return
    }

    if (!files.length) {
      ui.pstEmpty.classList.remove('hidden')
      ui.pstEmpty.innerHTML = escapeHtml(state.catalogMessage || 'No PST files found.')
      ui.pstList.innerHTML = ''
      return
    }

    ui.pstEmpty.classList.add('hidden')
    ui.pstList.innerHTML = files
      .map((file) => {
        const isActive = file.fileName === state.selectedPstFileName ? ' active' : ''
        const modifiedAt = file.modifiedAt ? formatDate(file.modifiedAt) : 'Unknown date'
        const size = Number.isFinite(file.size) ? formatBytes(file.size) : 'Unknown size'
        return `
          <button
            class="pst-item${isActive}"
            data-pst-file-name="${escapeAttr(file.fileName)}"
            title="${escapeAttr(file.fileName)}"
          >
            <span class="pst-item-name">${escapeHtml(file.fileName)}</span>
            <span class="pst-item-meta">${escapeHtml(size)} · ${escapeHtml(modifiedAt)}</span>
          </button>
        `
      })
      .join('')
  }

  function renderFolderNode(node) {
    const isActive = node.id === state.currentFolderId ? ' active' : ''
    const hasChildren = Array.isArray(node.children) && node.children.length > 0
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
          data-folder-id="${escapeAttr(node.id)}"
          title="${escapeAttr(node.path || node.displayName)}"
        >
          <span class="folder-name">${escapeHtml(node.displayName || '(untitled)')}</span>
          <span class="folder-badges">${badges}</span>
        </button>
        ${
          hasChildren
            ? `<ul class="folder-children">${node.children.map((child) => renderFolderNode(child)).join('')}</ul>`
            : ''
        }
      </li>
    `
  }

  function renderFolderTree() {
    if (!state.tree) {
      ui.folderTree.innerHTML =
        '<div class="panel-empty tree-empty">No mailbox loaded.</div>'
      return
    }

    ui.folderTree.innerHTML = `<ul class="folder-list">${renderFolderNode(state.tree)}</ul>`
  }

  function renderMessageRow(item) {
    const isActive = item.id === state.selectedMessageId ? ' active' : ''
    const sender = item.senderName || item.senderEmailAddress || '(unknown sender)'
    const time = item.sortDate || item.creationTime || item.clientSubmitTime
    const review = normalizeReviewState(item.review)
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

    return `
      <button class="message-row${isActive}" data-message-id="${escapeAttr(item.id)}">
        <div class="message-row-top">
          <div class="message-subject">${escapeHtml(item.subject || '(no subject)')}</div>
          <div class="message-date">${escapeHtml(formatDate(time))}</div>
        </div>
        <div class="message-sender">${escapeHtml(sender)}</div>
        <div class="message-row-meta">
          ${chips}
          <span>${escapeHtml(item.recipientText || item.displayTo || 'No recipients')}</span>
        </div>
      </button>
    `
  }

  function updatePagingButtons() {
    const page = state.currentFolderPage
    const hasPage = Boolean(page)
    ui.pagePrev.disabled = !hasPage || page.page <= 1
    ui.pageNext.disabled = !hasPage || page.page >= page.totalPages

    const items = page ? page.items || [] : []
    const currentIndex = state.selectedMessageId
      ? items.findIndex((item) => item.id === state.selectedMessageId)
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
    const page = state.currentFolderPage
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

    const folderName = page.folder?.displayName || 'folder'
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
    const recipientLines = [
      detail.displayTo ? `To ${detail.displayTo}` : '',
      detail.displayCC ? `Cc ${detail.displayCC}` : '',
      detail.displayBCC ? `Bcc ${detail.displayBCC}` : ''
    ]
      .filter(Boolean)
      .join(' · ')
    const sentTime = formatDate(detail.sortDate || detail.clientSubmitTime || detail.creationTime)
    const reviewPanel = renderReviewPanel(detail)

    return `
      <article class="detail-card" data-message-id="${escapeAttr(detail.id)}">
        <header class="outlook-header">
          <h3 class="detail-title">${escapeHtml(detail.subject || '(no subject)')}</h3>
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
                  ? `<span class="sender-address">&lt;${escapeHtml(senderEmail)}&gt;</span>`
                  : ''
              }
            </div>
            ${recipientLines ? `<div class="recipient-line">${escapeHtml(recipientLines)}</div>` : ''}
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
    setStatus(`Opening ${fileName}...`)
    try {
      const response = await fetchJson('/api/psts/open', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ fileName })
      })

      state.sessionId = response.sessionId
      state.selectedPstFileName = response.fileName || fileName
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
      const response = await fetchJson('/api/psts')
      state.catalogLoaded = true
      state.catalogMessage = response.message || ''
      state.catalog = Array.isArray(response.files) ? response.files : []
      renderPstCatalog()

      const preferredFileName =
        options.preferredFileName ||
        state.selectedPstFileName ||
        localStorage.getItem(STORAGE_KEYS.pstFileName) ||
        ''
      const preferred = preferredFileName
        ? state.catalog.find((item) => item.fileName === preferredFileName)
        : null
      const nextMailbox = preferred || state.catalog[0] || null

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
    if (state.query.trim()) {
      params.set('q', state.query.trim())
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

  async function navigateFolderPage(delta) {
    if (!state.currentFolderPage) {
      return
    }
    const nextPage = state.currentFolderPage.page + delta
    if (nextPage < 1 || nextPage > state.currentFolderPage.totalPages) {
      return
    }
    await loadFolderPage(nextPage, { selectPreferred: true })
  }

  async function navigateMessage(delta) {
    if (!state.currentFolderPage || !state.selectedMessageId) {
      return
    }

    const items = state.currentFolderPage.items || []
    const index = items.findIndex((item) => item.id === state.selectedMessageId)
    if (index >= 0) {
      const targetIndex = index + delta
      if (targetIndex >= 0 && targetIndex < items.length) {
        await selectMessage(items[targetIndex].id, { refresh: true })
        return
      }
    }

    const nextPage = state.currentFolderPage.page + (delta > 0 ? 1 : -1)
    if (nextPage < 1 || nextPage > state.currentFolderPage.totalPages) {
      return
    }

    const token = (state.folderLoadToken = (state.folderLoadToken || 0) + 1)
    const folderId = state.currentFolderId
    const params = new URLSearchParams()
    params.set('page', String(nextPage))
    params.set('pageSize', String(state.pageSize))
    params.set('mailOnly', state.mailOnly ? '1' : '0')
    params.set('sort', state.sort)
    if (state.reviewFlaggedOnly) {
      params.set('reviewFlagged', '1')
    }
    if (state.reviewTaggedOnly) {
      params.set('reviewTagged', '1')
    }
    if (state.query.trim()) {
      params.set('q', state.query.trim())
    }

    try {
      const response = await fetchJson(
        `/api/sessions/${encodeURIComponent(state.sessionId)}/folders/${encodeURIComponent(
          folderId
        )}/messages?${params.toString()}`
      )
      if (token !== state.folderLoadToken) {
        return
      }
      state.currentFolderPage = response.page
      renderMessageList()
      const pageItems = response.page?.items || []
      const target = delta > 0 ? pageItems[0] : pageItems[pageItems.length - 1]
      if (target) {
        await selectMessage(target.id, { refresh: true })
      }
    } catch (error) {
      setStatus(`Unable to navigate messages: ${error.message}`, 'error')
    }
  }

  async function chooseFolder(folderId) {
    if (!folderId || folderId === state.currentFolderId) {
      return
    }
    state.currentFolderId = folderId
    state.currentFolderPage = null
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
      void loadMailboxCatalog({ showBusy: true })
    })

    ui.pstList.addEventListener('click', (event) => {
      const button = event.target.closest('[data-pst-file-name]')
      if (!button) {
        return
      }
      const fileName = button.dataset.pstFileName
      if (!fileName) {
        return
      }
      void openMailbox(fileName, { showBusy: true })
    })

    ui.folderTree.addEventListener('click', (event) => {
      const button = event.target.closest('[data-folder-id]')
      if (!button) {
        return
      }
      void chooseFolder(button.dataset.folderId)
    })

    ui.messageList.addEventListener('click', (event) => {
      const button = event.target.closest('[data-message-id]')
      if (!button) {
        return
      }
      void selectMessage(button.dataset.messageId, { refresh: true })
    })

    ui.messageDetail.addEventListener('click', (event) => {
      const actionButton = event.target.closest('[data-action]')
      if (!actionButton) {
        return
      }

      const action = actionButton.dataset.action
      if (action === 'download-json') {
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

    ui.searchInput.addEventListener(
      'input',
      debounce(() => {
        state.query = ui.searchInput.value.trim()
        saveState()
        if (state.sessionId && state.currentFolderId) {
          void loadFolderPage(1, { selectPreferred: true })
        }
      }, 250)
    )

    ui.mailOnlyToggle.addEventListener('change', () => {
      state.mailOnly = ui.mailOnlyToggle.checked
      saveState()
      if (state.sessionId && state.currentFolderId) {
        void loadFolderPage(1, { selectPreferred: true })
      }
    })

    ui.reviewFlaggedToggle.addEventListener('change', () => {
      state.reviewFlaggedOnly = ui.reviewFlaggedToggle.checked
      saveState()
      if (state.sessionId && state.currentFolderId) {
        void loadFolderPage(1, { selectPreferred: true })
      }
    })

    ui.reviewTaggedToggle.addEventListener('change', () => {
      state.reviewTaggedOnly = ui.reviewTaggedToggle.checked
      saveState()
      if (state.sessionId && state.currentFolderId) {
        void loadFolderPage(1, { selectPreferred: true })
      }
    })

    ui.sortSelect.addEventListener('change', () => {
      state.sort = ui.sortSelect.value
      saveState()
      if (state.sessionId && state.currentFolderId) {
        void loadFolderPage(1, { selectPreferred: true })
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
    ui.searchInput = getElement('message-search')
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
    state.mailOnly = readStorageBool(STORAGE_KEYS.mailOnly, true)
    state.sort = localStorage.getItem(STORAGE_KEYS.sort) || 'date-desc'
    state.reviewFlaggedOnly = readStorageBool(STORAGE_KEYS.reviewFlaggedOnly, false)
    state.reviewTaggedOnly = readStorageBool(STORAGE_KEYS.reviewTaggedOnly, false)
    state.selectedPstFileName = localStorage.getItem(STORAGE_KEYS.pstFileName) || null
    refreshControls()
    wireEvents()
    renderPstCatalog()
    await loadMailboxCatalog({
      showBusy: true,
      preferredFileName: state.selectedPstFileName || undefined
    })
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
