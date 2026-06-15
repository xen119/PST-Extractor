import { normalizeText } from './utils'

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
  theme: 'pst-mail-explorer.theme',
  hiddenFiltersOpen: 'pst-mail-explorer.hiddenFiltersOpen',
  activityFilterUser: 'pst-mail-explorer.activityFilterUser'
} as const

export type WorkspaceStorageKey = keyof typeof STORAGE_KEYS

export function normalizeStorageNamespace(value: unknown): string {
  const text = normalizeText(value).toLowerCase()
  if (!text) {
    return 'local'
  }

  return text.replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '') || 'local'
}

export function getWorkspaceStorageNamespace(authenticated: boolean, authUser: string): string {
  if (!authenticated) {
    return 'local'
  }

  return normalizeStorageNamespace(authUser)
}

export function getWorkspaceStorageKey(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string
): string {
  return `${STORAGE_KEYS[key]}::${getWorkspaceStorageNamespace(authenticated, authUser)}`
}

export function readWorkspaceStorageItem(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string,
  fallback = ''
): string {
  const scopedKey = getWorkspaceStorageKey(key, authenticated, authUser)
  const stored = localStorage.getItem(scopedKey)
  if (stored != null) {
    return stored
  }

  const legacyKey = STORAGE_KEYS[key]
  const legacy = localStorage.getItem(legacyKey)
  if (legacy != null) {
    try {
      localStorage.setItem(scopedKey, legacy)
    } catch {
      // Ignore storage migration errors in restricted browser contexts.
    }
    return legacy
  }

  return fallback
}

export function readWorkspaceStorageBool(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string,
  fallback: boolean
): boolean {
  const raw = readWorkspaceStorageItem(key, authenticated, authUser, '')
  if (!raw) {
    return fallback
  }

  return !['0', 'false', 'no', 'off'].includes(raw.toLowerCase())
}

export function writeWorkspaceStorageItem(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string,
  value: string | number | boolean | null | undefined
): void {
  const scopedKey = getWorkspaceStorageKey(key, authenticated, authUser)
  if (value == null || value === '') {
    localStorage.removeItem(scopedKey)
    return
  }

  localStorage.setItem(scopedKey, String(value))
}

export function removeWorkspaceStorageItem(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string
): void {
  localStorage.removeItem(getWorkspaceStorageKey(key, authenticated, authUser))
}
