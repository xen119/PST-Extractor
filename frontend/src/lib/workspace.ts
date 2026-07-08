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
  sourceType: 'pst-mail-explorer.sourceType',
  mailOnly: 'pst-mail-explorer.mailOnly',
  sort: 'pst-mail-explorer.sort',
  reviewFlaggedOnly: 'pst-mail-explorer.reviewFlaggedOnly',
  reviewTaggedOnly: 'pst-mail-explorer.reviewTaggedOnly',
  flaggedBundleSizePreset: 'pst-mail-explorer.flaggedBundleSizePreset',
  flaggedBundleCustomSizeMb: 'pst-mail-explorer.flaggedBundleCustomSizeMb',
  flaggedBundleScope: 'pst-mail-explorer.flaggedBundleScope',
  theme: 'pst-mail-explorer.theme',
  hiddenFiltersOpen: 'pst-mail-explorer.hiddenFiltersOpen',
  activityFilterUser: 'pst-mail-explorer.activityFilterUser'
} as const

export type WorkspaceStorageKey = keyof typeof STORAGE_KEYS

function getWorkspaceStorage(): Storage | null {
  try {
    if (typeof window === 'undefined' || typeof window.localStorage === 'undefined') {
      return null
    }
    return window.localStorage
  } catch {
    return null
  }
}

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
  const storage = getWorkspaceStorage()
  if (!storage) {
    return fallback
  }

  const stored = storage.getItem(scopedKey)
  if (stored != null) {
    return stored
  }

  const legacyKey = STORAGE_KEYS[key]
  const legacy = storage.getItem(legacyKey)
  if (legacy != null) {
    try {
      storage.setItem(scopedKey, legacy)
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
  const storage = getWorkspaceStorage()
  if (!storage) {
    return
  }

  const scopedKey = getWorkspaceStorageKey(key, authenticated, authUser)
  if (value == null || value === '') {
    storage.removeItem(scopedKey)
    return
  }

  storage.setItem(scopedKey, String(value))
}

export function removeWorkspaceStorageItem(
  key: WorkspaceStorageKey,
  authenticated: boolean,
  authUser: string
): void {
  const storage = getWorkspaceStorage()
  if (!storage) {
    return
  }

  storage.removeItem(getWorkspaceStorageKey(key, authenticated, authUser))
}
