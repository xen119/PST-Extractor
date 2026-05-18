(function (globalScope, factory) {
  const api = factory()

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api
  }

  globalScope.PstExplorerSearch = api
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  function normalizeMessageId(item) {
    if (!item || typeof item !== 'object') {
      return ''
    }

    return String(item.id || item.messageId || '').trim()
  }

  function normalizeSearchResultItem(item) {
    if (!item || typeof item !== 'object') {
      return item
    }

    const id = normalizeMessageId(item)
    if (!id) {
      return { ...item }
    }

    return {
      ...item,
      id,
      messageId: String(item.messageId || item.id || id).trim()
    }
  }

  function normalizeSearchResultsPage(page) {
    if (!page || typeof page !== 'object') {
      return page
    }

    const items = Array.isArray(page.items) ? page.items : []
    return {
      ...page,
      items: items.map((item) => normalizeSearchResultItem(item))
    }
  }

  return {
    normalizeMessageId,
    normalizeSearchResultItem,
    normalizeSearchResultsPage
  }
})
