import * as fs from 'fs'
import * as path from 'path'
import { createViewerSession, type ViewerSessionIndex } from './viewer'

export interface PstCatalogEntry {
  fileName: string
  size: number
  modifiedAt: string | null
}

export interface PstCatalogResult {
  rootPath: string
  rootExists: boolean
  files: PstCatalogEntry[]
  message: string
}

const SUPPORTED_EXTENSIONS = new Set(['.pst', '.ost'])

export function getDefaultPstRootDirectory(): string {
  return path.resolve(__dirname, '..', 'PST')
}

export function isSupportedMailboxFile(fileName: string): boolean {
  return SUPPORTED_EXTENSIONS.has(path.extname(fileName).toLowerCase())
}

export function resolvePstMailboxPath(rootPath: string, fileName: string): string {
  const candidate = String(fileName || '').trim()
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
  const resolvedPath = path.resolve(resolvedRoot, candidate)
  const relative = path.relative(resolvedRoot, resolvedPath)
  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error('Mailbox file must stay within the PST folder')
  }
  return resolvedPath
}

export function listPstMailboxFiles(rootPath = getDefaultPstRootDirectory()): PstCatalogResult {
  const resolvedRoot = path.resolve(rootPath)
  if (!fs.existsSync(resolvedRoot) || !fs.statSync(resolvedRoot).isDirectory()) {
    return {
      rootPath: resolvedRoot,
      rootExists: false,
      files: [],
      message: 'Create a PST folder at the project root and place .pst or .ost files in it.'
    }
  }

  const files = fs
    .readdirSync(resolvedRoot, { withFileTypes: true })
    .filter((entry) => entry.isFile() && isSupportedMailboxFile(entry.name))
    .map((entry) => {
      const filePath = path.join(resolvedRoot, entry.name)
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

  return {
    rootPath: resolvedRoot,
    rootExists: true,
    files,
    message:
      files.length > 0
        ? `Found ${files.length} mailbox file${files.length === 1 ? '' : 's'} in PST/.`
        : 'No PST or OST files were found in the PST folder.'
  }
}

export function openPstMailbox(
  rootPath: string,
  fileName: string
): ViewerSessionIndex {
  const resolvedPath = resolvePstMailboxPath(rootPath, fileName)
  if (!fs.existsSync(resolvedPath) || !fs.statSync(resolvedPath).isFile()) {
    throw new Error(`Mailbox not found: ${fileName}`)
  }
  return createViewerSession(resolvedPath, path.basename(fileName))
}
