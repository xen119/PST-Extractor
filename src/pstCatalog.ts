import * as fs from 'fs'
import * as path from 'path'
import { createViewerSession, type ViewerSessionIndex } from './viewer'

export interface PstCatalogEntry {
  fileName: string
  size: number
  modifiedAt: string | null
}

export interface PstCatalogScopeEntry {
  scopePath: string
  scopeLabel: string
  fileCount: number
  files: PstCatalogEntry[]
}

export interface PstCatalogResult {
  rootPath: string
  rootExists: boolean
  scopes: PstCatalogScopeEntry[]
  scopePath: string
  scopeLabel: string
  files: PstCatalogEntry[]
  message: string
}

const SUPPORTED_EXTENSIONS = new Set(['.pst', '.ost'])

function normalizeScopePath(scopePath: unknown): string {
  const text = String(scopePath ?? '')
    .trim()
    .replace(/\\/g, '/')

  if (!text || text === '.') {
    return ''
  }

  if (path.isAbsolute(text)) {
    throw new Error('Scope path must stay within the PST folder')
  }

  const segments = text
    .split('/')
    .map((segment) => segment.trim())
    .filter(Boolean)

  if (!segments.length) {
    return ''
  }

  if (segments.some((segment) => segment === '..')) {
    throw new Error('Scope path must stay within the PST folder')
  }

  return segments.join('/')
}

function resolvePstScopeDirectory(rootPath: string, scopePath: unknown): string {
  const resolvedRoot = path.resolve(rootPath)
  const normalizedScopePath = normalizeScopePath(scopePath)
  const resolvedScopePath = normalizedScopePath
    ? path.resolve(resolvedRoot, ...normalizedScopePath.split('/'))
    : resolvedRoot
  const relative = path.relative(resolvedRoot, resolvedScopePath)

  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('Scope path must stay within the PST folder')
  }

  return resolvedScopePath
}

function formatScopeLabel(scopePath: string): string {
  return scopePath ? scopePath.split('/').join(' / ') : 'PST root'
}

function readMailboxEntries(directoryPath: string): PstCatalogEntry[] {
  return fs
    .readdirSync(directoryPath, { withFileTypes: true })
    .filter((entry) => entry.isFile() && isSupportedMailboxFile(entry.name))
    .map((entry) => {
      const filePath = path.join(directoryPath, entry.name)
      const stats = fs.statSync(filePath)
      return {
        fileName: entry.name,
        size: stats.size,
        modifiedAt: stats.mtime.toISOString()
      }
    })
    .sort((left, right) =>
      left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
    )
}

function discoverMailboxScopes(rootPath: string): PstCatalogScopeEntry[] {
  const scopes: PstCatalogScopeEntry[] = []

  function visit(directoryPath: string, scopePath: string): void {
    const entries = fs.readdirSync(directoryPath, { withFileTypes: true })
    const files = readMailboxEntries(directoryPath)

    if (files.length > 0) {
      const normalizedScopePath = normalizeScopePath(scopePath)
      scopes.push({
        scopePath: normalizedScopePath,
        scopeLabel: formatScopeLabel(normalizedScopePath),
        fileCount: files.length,
        files
      })
    }

    const childDirectories = entries
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name)
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: 'base' }))

    for (const childName of childDirectories) {
      const childDirectoryPath = path.join(directoryPath, childName)
      const childScopePath = scopePath ? `${scopePath}/${childName}` : childName
      visit(childDirectoryPath, childScopePath)
    }
  }

  visit(rootPath, '')

  scopes.sort((left, right) => {
    if (left.scopePath === right.scopePath) {
      return 0
    }
    if (left.scopePath === '') {
      return -1
    }
    if (right.scopePath === '') {
      return 1
    }
    return left.scopeLabel.localeCompare(right.scopeLabel, undefined, { sensitivity: 'base' })
  })

  return scopes
}

function chooseScopeEntry(
  scopes: PstCatalogScopeEntry[],
  requestedScopePath: unknown
): PstCatalogScopeEntry | null {
  const normalizedRequestedScopePath = normalizeScopePath(requestedScopePath)
  if (normalizedRequestedScopePath) {
    const requestedScope = scopes.find(
      (scope) => scope.scopePath === normalizedRequestedScopePath
    )
    if (requestedScope) {
      return requestedScope
    }
  }

  const rootScope = scopes.find((scope) => scope.scopePath === '')
  return rootScope || scopes[0] || null
}

export function getDefaultPstRootDirectory(): string {
  return path.resolve(__dirname, '..', 'PST')
}

export function isSupportedMailboxFile(fileName: string): boolean {
  return SUPPORTED_EXTENSIONS.has(path.extname(fileName).toLowerCase())
}

export function resolvePstMailboxPath(rootPath: string, fileName: string): string
export function resolvePstMailboxPath(
  rootPath: string,
  scopePath: string,
  fileName: string
): string
export function resolvePstMailboxPath(
  rootPath: string,
  scopePathOrFileName: string,
  fileNameMaybe?: string
): string {
  const hasScopePath = fileNameMaybe !== undefined
  const scopePath = hasScopePath ? scopePathOrFileName : ''
  const candidateFileName = hasScopePath ? fileNameMaybe : scopePathOrFileName
  const candidate = String(candidateFileName || '').trim()

  if (!candidate) {
    throw new Error('Mailbox file name is required')
  }
  if (candidate !== path.basename(candidate)) {
    throw new Error('Mailbox file name must not include a path')
  }
  if (!isSupportedMailboxFile(candidate)) {
    throw new Error('Only .pst and .ost files are supported')
  }

  const resolvedRoot = path.resolve(rootPath)
  const resolvedScopeDirectory = resolvePstScopeDirectory(resolvedRoot, scopePath)
  const resolvedPath = path.resolve(resolvedScopeDirectory, candidate)
  const relative = path.relative(resolvedRoot, resolvedPath)

  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('Mailbox file must stay within the PST folder')
  }

  return resolvedPath
}

export function listPstMailboxFiles(
  rootPath = getDefaultPstRootDirectory(),
  scopePath?: string
): PstCatalogResult {
  const resolvedRoot = path.resolve(rootPath)
  if (!fs.existsSync(resolvedRoot) || !fs.statSync(resolvedRoot).isDirectory()) {
    return {
      rootPath: resolvedRoot,
      rootExists: false,
      scopes: [],
      scopePath: '',
      scopeLabel: '',
      files: [],
      message: 'Create a PST folder at the project root and place .pst or .ost files in it.'
    }
  }

  const scopes = discoverMailboxScopes(resolvedRoot)
  const selectedScope = chooseScopeEntry(scopes, scopePath)
  const files = selectedScope?.files || []

  return {
    rootPath: resolvedRoot,
    rootExists: true,
    scopes,
    scopePath: selectedScope?.scopePath || '',
    scopeLabel: selectedScope?.scopeLabel || '',
    files,
    message:
      scopes.length > 0
        ? `Found ${selectedScope?.fileCount || 0} mailbox file${
            selectedScope && selectedScope.fileCount === 1 ? '' : 's'
          } in ${selectedScope?.scopeLabel || 'PST root'}.`
        : 'No PST or OST files were found anywhere under the PST folder.'
  }
}

export function openPstMailbox(rootPath: string, fileName: string): ViewerSessionIndex
export function openPstMailbox(
  rootPath: string,
  scopePath: string,
  fileName: string
): ViewerSessionIndex
export function openPstMailbox(
  rootPath: string,
  scopePathOrFileName: string,
  fileNameMaybe?: string
): ViewerSessionIndex {
  const resolvedPath =
    fileNameMaybe === undefined
      ? resolvePstMailboxPath(rootPath, scopePathOrFileName)
      : resolvePstMailboxPath(rootPath, scopePathOrFileName, fileNameMaybe)
  const resolvedFileName = String(
    fileNameMaybe === undefined ? path.basename(scopePathOrFileName) : fileNameMaybe
  )

  if (!fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) {
    throw new Error(`Mailbox not found: ${resolvedFileName}`)
  }

  return createViewerSession(resolvedPath, path.basename(resolvedFileName))
}
