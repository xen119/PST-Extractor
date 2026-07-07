import { normalizeText } from './utils'

export interface SearchTerm {
  text: string
  phrase: boolean
}

export interface ParsedSearchQuery {
  mode: 'and' | 'or'
  terms: SearchTerm[]
}

export interface ResolvedSelectionScope {
  scope: 'all' | 'search'
  scopePath: string
  casePath?: string
}

export function normalizeSearchResultItem<T extends Record<string, unknown>>(item: T): T {
  if (!item || typeof item !== 'object') {
    return item
  }

  const id = normalizeText((item as Record<string, unknown>).id || (item as Record<string, unknown>).messageId)
  if (!id) {
    return { ...item }
  }

  return {
    ...item,
    id,
    messageId: normalizeText((item as Record<string, unknown>).messageId || id)
  }
}

export function normalizeSearchResultsPage<T extends { items?: unknown[] }>(page: T): T {
  if (!page || typeof page !== 'object') {
    return page
  }

  const items = Array.isArray(page.items) ? page.items : []
  return {
    ...page,
    items: items.map((item) => normalizeSearchResultItem(item as Record<string, unknown>))
  }
}

export function resolveSelectionScope(selectedCasePath: string, selectedScopePath: string): ResolvedSelectionScope {
  const normalizedCasePath = normalizeText(selectedCasePath)
  const normalizedScopePath = normalizeText(selectedScopePath)

  if (!normalizedCasePath) {
    return {
      scope: 'all',
      scopePath: ''
    }
  }

  if (normalizedScopePath) {
    return {
      scope: 'search',
      scopePath: normalizedScopePath,
      casePath: normalizedCasePath || undefined
    }
  }

  if (normalizedCasePath) {
    return {
      scope: 'all',
      scopePath: '',
      casePath: normalizedCasePath
    }
  }

  return {
    scope: 'all',
    scopePath: ''
  }
}

export function deriveSearchMode(query: string, fallback: 'and' | 'or' = 'and'): 'and' | 'or' {
  const { mode } = parseSearchQuery(query, fallback)
  return mode
}

export function parseSearchQuery(query: string, fallbackMode: 'and' | 'or' = 'and'): ParsedSearchQuery {
  const text = normalizeText(query)
  if (!text) {
    return { mode: fallbackMode, terms: [] }
  }

  const terms: SearchTerm[] = []
  let mode: 'and' | 'or' = fallbackMode
  const pattern = /"([^"]+)"|(\S+)/g
  let match: RegExpExecArray | null
  while ((match = pattern.exec(text))) {
    const raw = match[1] || match[2] || ''
    const normalized = normalizeText(raw)
    if (!normalized) {
      continue
    }
    if (!match[1]) {
      const lowered = normalized.toLowerCase()
      if (lowered === '|' || lowered.startsWith('|')) {
        mode = 'or'
        const remainder = normalized.length > 1 ? normalizeText(normalized.slice(1)) : ''
        if (remainder) {
          terms.push({ text: remainder, phrase: false })
        }
        continue
      }
      if (lowered === '+' || lowered.startsWith('+')) {
        mode = 'and'
        const remainder = normalized.length > 1 ? normalizeText(normalized.slice(1)) : ''
        if (remainder) {
          terms.push({ text: remainder, phrase: false })
        }
        continue
      }
    }
    terms.push({
      text: normalized,
      phrase: Boolean(match[1])
    })
  }

  return { mode, terms }
}
