import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { createHash, randomBytes } from 'crypto'
import { spawn } from 'child_process'
import { EventEmitter } from 'events'
import { Readable } from 'stream'
// ts-node in the example app only compiles the entrypoint and imported TS files.
// Use require here so the runtime does not depend on an ambient module declaration.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const AdmZip = require('adm-zip') as any
// eslint-disable-next-line @typescript-eslint/no-var-requires
const yauzl = require('yauzl') as any
import type { PstCatalogEntry, PstCatalogResult, PstCatalogScopeEntry } from './pstCatalog'

export type ArchiveBundleSourceType = 'teams' | 'sharepoint'
export type ArchivePreviewKind = 'text' | 'html' | 'binary'

export interface ArchiveBundleItem {
  bundlePath: string
  bundleFileName: string
  scopePath: string
  scopeLabel: string
  sourceType: ArchiveBundleSourceType
  entryChain: string[]
  entryPath: string
  entryName: string
  entrySize: number
  modifiedAt: string | null
  contentType: string
  previewKind: ArchivePreviewKind
  previewText: string
  previewHtml: string
  downloadFilename: string
  searchText: string
  archiveItemId: string
}

const SUPPORTED_BUNDLE_EXTENSION = '.zip'
const MAX_BUNDLE_DEPTH = 4
const MAX_ENTRY_BYTES = 25 * 1024 * 1024
const MAX_PREVIEW_BYTES = 256 * 1024
const MAX_TEXT_LENGTH = 120_000

interface ArchiveEntryLike {
  fileName: string
  entryName?: string
  isDirectory: boolean
  uncompressedSize: number
  compressedSize?: number
  getLastModDate?: () => Date | null
}

class TarArchiveSource extends EventEmitter {
  private readonly entries: ArchiveEntryLike[]
  private readonly tempFilePath: string | null
  private index = 0
  private closed = false

  constructor(
    private readonly archivePath: string,
    entries: ArchiveEntryLike[],
    tempFilePath: string | null = null
  ) {
    super()
    this.entries = entries
    this.tempFilePath = tempFilePath
  }

  readEntry(): void {
    if (this.closed) {
      return
    }

    const entry = this.entries[this.index++]
    if (!entry) {
      process.nextTick(() => this.emit('end'))
      return
    }

    process.nextTick(() => this.emit('entry', entry))
  }

  openReadStream(entry: ArchiveEntryLike, callback: (error: unknown, stream?: NodeJS.ReadableStream | null) => void): void {
    const entryName = normalizeEntryPath(entry.fileName || entry.entryName || '')
    readTarEntryBuffer(this.archivePath, entryName, MAX_ENTRY_BYTES)
      .then((buffer) => {
        if (!buffer) {
          callback(new Error(`Unable to read archive entry: ${entryName}`), null)
          return
        }
        callback(null, Readable.from(buffer))
      })
      .catch((error) => {
        callback(error, null)
      })
  }

  close(): void {
    if (this.closed) {
      return
    }
    this.closed = true
    if (!this.tempFilePath) {
      return
    }
    try {
      fs.rmSync(this.tempFilePath, { force: true })
    } catch {
      // ignore
    }
  }
}

function parseTarDate(month: string, day: string, yearOrTime: string): Date | null {
  const currentYear = new Date().getFullYear()
  const timestamp =
    /^\d{4}$/.test(yearOrTime)
      ? Date.parse(`${month} ${day} ${yearOrTime} 00:00:00 GMT`)
      : Date.parse(`${month} ${day} ${currentYear} ${yearOrTime}:00 GMT`)
  if (Number.isNaN(timestamp)) {
    return null
  }
  return new Date(timestamp)
}

function parseTarListingLine(line: string): ArchiveEntryLike | null {
  const normalizedLine = line.trim()
  if (!normalizedLine) {
    return null
  }

  const match = normalizedLine.match(/^(\S+)\s+(.*?)\s+(\d+)\s+([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4}|\d{2}:\d{2})\s+(.+)$/)
  if (!match) {
    return null
  }

  const [, mode, , sizeText, month, day, yearOrTime, rawName] = match
  const name = normalizeEntryPath(rawName || '')
  if (!name) {
    return null
  }

  const size = Number.parseInt(sizeText || '0', 10)
  const modifiedAt = parseTarDate(month || '', day || '', yearOrTime || '')
  return {
    fileName: name,
    entryName: name,
    isDirectory: (mode || '').startsWith('d') || name.endsWith('/'),
    uncompressedSize: Number.isFinite(size) ? size : 0,
    compressedSize: Number.isFinite(size) ? size : 0,
    getLastModDate: () => modifiedAt
  }
}

async function listTarEntries(archivePath: string): Promise<ArchiveEntryLike[]> {
  return new Promise((resolve, reject) => {
    const child = spawn('tar', ['-tvf', archivePath], {
      windowsHide: true
    })
    const stdoutChunks: Buffer[] = []
    const stderrChunks: Buffer[] = []

    child.stdout.on('data', (chunk: Buffer) => {
      stdoutChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
    })
    child.stderr.on('data', (chunk: Buffer) => {
      stderrChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
    })
    child.on('error', (error) => {
      reject(error)
    })
    child.on('close', (code) => {
      if (code !== 0) {
        const stderr = Buffer.concat(stderrChunks).toString('utf8').trim()
        reject(new Error(stderr || `tar exited with code ${code}`))
        return
      }

      const lines = Buffer.concat(stdoutChunks).toString('utf8').split(/\r?\n/)
      const entries = lines.map(parseTarListingLine).filter((entry): entry is ArchiveEntryLike => Boolean(entry))
      if (!entries.length && lines.some((line) => line.trim())) {
        reject(new Error('Unable to parse archive listing'))
        return
      }
      resolve(entries)
    })
  })
}

async function readTarEntryBuffer(archivePath: string, entryPath: string, maxBytes = MAX_ENTRY_BYTES): Promise<Buffer | null> {
  return new Promise((resolve, reject) => {
    const child = spawn('tar', ['-xOf', archivePath, entryPath], {
      windowsHide: true
    })
    const chunks: Buffer[] = []
    let total = 0
    let settled = false

    const settle = (value: Buffer | null): void => {
      if (settled) {
        return
      }
      settled = true
      resolve(value)
    }

    child.stdout.on('data', (chunk: Buffer) => {
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
      total += buffer.length
      if (total > maxBytes) {
        child.kill()
        settle(null)
        return
      }
      chunks.push(buffer)
    })
    child.stderr.on('data', () => {
      // ignore tar progress noise and missing-entry warnings here
    })
    child.on('error', (error) => {
      if (!settled) {
        settled = true
        reject(error)
      }
    })
    child.on('close', (code) => {
      if (settled) {
        return
      }
      if (code !== 0) {
        settle(null)
        return
      }
      settle(Buffer.concat(chunks))
    })
  })
}

async function createTarArchiveSourceFromPath(archivePath: string, tempFilePath: string | null = null): Promise<TarArchiveSource> {
  const entries = (await listTarEntries(archivePath)).sort((left, right) =>
    left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
  )
  return new TarArchiveSource(archivePath, entries, tempFilePath)
}

async function openZipFileWithYauzl(zipPath: string): Promise<any> {
  return new Promise((resolve, reject) => {
    yauzl.open(
      zipPath,
      {
        lazyEntries: true,
        decodeStrings: true
      },
      (error: unknown, zipFile: unknown) => {
        if (error || !zipFile) {
          reject(error || new Error(`Unable to open archive: ${zipPath}`))
          return
        }
        resolve(zipFile)
      }
    )
  })
}

async function openZipBufferWithYauzl(buffer: Buffer): Promise<any> {
  return new Promise((resolve, reject) => {
    yauzl.fromBuffer(
      buffer,
      {
        lazyEntries: true,
        decodeStrings: true
      },
      (error: unknown, zipFile: unknown) => {
        if (error || !zipFile) {
          reject(error || new Error('Unable to open nested archive'))
          return
        }
        resolve(zipFile)
      }
    )
  })
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

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

function formatScopeLabel(scopePath: string): string {
  return scopePath ? scopePath.split('/').join(' / ') : 'PST root'
}

function isTopLevelBundleFile(fileName: string): boolean {
  return /^Items.*\.zip$/i.test(fileName)
}

function isBundleFile(fileName: string): boolean {
  return path.extname(fileName).toLowerCase() === SUPPORTED_BUNDLE_EXTENSION
}

async function openZipFile(zipPath: string): Promise<any> {
  const resolvedPath = path.resolve(zipPath)

  try {
    return await createTarArchiveSourceFromPath(resolvedPath)
  } catch {
    return openZipFileWithYauzl(resolvedPath)
  }
}

async function openZipBuffer(buffer: Buffer): Promise<any> {
  try {
    return await openZipBufferWithYauzl(buffer)
  } catch {
    const tempFilePath = path.join(
      os.tmpdir(),
      `pst-archive-${Date.now()}-${randomBytes(6).toString('hex')}.zip`
    )
    await fs.promises.writeFile(tempFilePath, buffer)
    try {
      return await createTarArchiveSourceFromPath(tempFilePath, tempFilePath)
    } catch (error) {
      try {
        fs.rmSync(tempFilePath, { force: true })
      } catch {
        // ignore
      }
      throw error
    }
  }
}

function collectZipEntries(zipFile: any): Promise<any[]> {
  return new Promise((resolve, reject) => {
    const entries: any[] = []
    let settled = false

    const settleResolve = (): void => {
      if (settled) {
        return
      }
      settled = true
      resolve(entries)
    }

    const settleReject = (error: unknown): void => {
      if (settled) {
        return
      }
      settled = true
      try {
        zipFile.close()
      } catch {
        // ignore
      }
      reject(error)
    }

    zipFile.on('entry', (entry: any) => {
      entries.push(entry)
      zipFile.readEntry()
    })
    zipFile.on('end', settleResolve)
    zipFile.on('error', settleReject)
    zipFile.readEntry()
  })
}

function readZipEntryBufferAsync(zipFile: any, entry: any): Promise<Buffer | null> {
  return new Promise((resolve) => {
    if (!entry || entry.uncompressedSize > MAX_ENTRY_BYTES) {
      resolve(null)
      return
    }

    zipFile.openReadStream(entry, (error: unknown, readStream: unknown) => {
      if (error || !readStream || typeof (readStream as NodeJS.ReadableStream).on !== 'function') {
        resolve(null)
        return
      }

      const chunks: Buffer[] = []
      const stream = readStream as NodeJS.ReadableStream
      stream.on('data', (chunk: Buffer) => {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
      })
      stream.on('error', () => resolve(null))
      stream.on('end', () => resolve(Buffer.concat(chunks)))
    })
  })
}

function buildScopeEntries(rootPath: string): PstCatalogScopeEntry[] {
  const scopes: PstCatalogScopeEntry[] = []

  function visit(directoryPath: string, scopePath: string): void {
    const entries = fs.readdirSync(directoryPath, { withFileTypes: true })
    const files: PstCatalogEntry[] = entries
      .filter((entry) => entry.isFile() && isTopLevelBundleFile(entry.name))
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
      .filter((childName) => childName.toLowerCase() !== '_removed')
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
    const requestedScope = scopes.find((scope) => scope.scopePath === normalizedRequestedScopePath)
    if (requestedScope) {
      return requestedScope
    }
  }

  const rootScope = scopes.find((scope) => scope.scopePath === '')
  return rootScope || scopes[0] || null
}

function normalizeEntryPath(entryName: string): string {
  return entryName.replace(/\\/g, '/').replace(/^\/+/, '')
}

function getEntryName(entryPath: string): string {
  const normalized = normalizeEntryPath(entryPath)
  return normalized.split('/').pop() || normalized
}

function inferContentType(entryPath: string): string {
  const ext = path.extname(entryPath).toLowerCase()
  switch (ext) {
    case '.html':
    case '.htm':
      return 'text/html; charset=utf-8'
    case '.txt':
    case '.csv':
    case '.xml':
    case '.ics':
      return 'text/plain; charset=utf-8'
    case '.json':
      return 'application/json; charset=utf-8'
    case '.pdf':
      return 'application/pdf'
    case '.png':
      return 'image/png'
    case '.bmp':
      return 'image/bmp'
    case '.jpg':
    case '.jpeg':
      return 'image/jpeg'
    case '.gif':
      return 'image/gif'
    case '.webp':
      return 'image/webp'
    case '.svg':
      return 'image/svg+xml'
    case '.docx':
      return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    case '.pptx':
      return 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    case '.xlsx':
      return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    case '.doc':
      return 'application/msword'
    case '.ppt':
      return 'application/vnd.ms-powerpoint'
    case '.xls':
      return 'application/vnd.ms-excel'
    default:
      return 'application/octet-stream'
  }
}

function classifySourceType(entryPath: string): ArchiveBundleSourceType {
  const normalized = normalizeEntryPath(entryPath).toLowerCase()
  if (normalized === 'sharepoint' || normalized.startsWith('sharepoint/') || normalized.includes('/sharepoint/')) {
    return 'sharepoint'
  }
  return 'teams'
}

function stripHtml(value: string): string {
  return value
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim()
}

function extractXmlText(xml: string): string {
  return stripHtml(xml)
}

function extractOfficeText(zip: any, entryPath: string): string {
  const lower = entryPath.toLowerCase()
  const candidates =
    lower.endsWith('.docx')
      ? ['word/document.xml']
      : lower.endsWith('.pptx')
        ? zip
            .getEntries()
            .map((entry: any) => entry.entryName)
            .filter((name: string) => name.toLowerCase().startsWith('ppt/slides/') && name.toLowerCase().endsWith('.xml'))
        : lower.endsWith('.xlsx')
          ? zip
              .getEntries()
              .map((entry: any) => entry.entryName)
              .filter((name: string) => {
                const normalized = name.toLowerCase()
                return (
                  normalized === 'xl/sharedstrings.xml' ||
                  normalized.startsWith('xl/worksheets/') ||
                  normalized.startsWith('xl/chartsheets/')
                )
              })
          : []

  const texts: string[] = []
  for (const candidate of candidates) {
    const entry = zip.getEntry(candidate)
    if (!entry) {
      continue
    }
    const buffer = zip.readFile(entry) || Buffer.alloc(0)
    const text = extractXmlText(buffer.toString('utf8'))
    if (text) {
      texts.push(text)
    }
  }
  return truncateText(texts.join(' '))
}

function truncateText(value: string): string {
  const normalized = normalizeText(value)
  if (normalized.length <= MAX_TEXT_LENGTH) {
    return normalized
  }
  return `${normalized.slice(0, MAX_TEXT_LENGTH).trim()}`
}

function prettyPrintJson(value: string): string {
  const normalized = value.trim()
  if (!normalized) {
    return ''
  }

  try {
    return JSON.stringify(JSON.parse(normalized), null, 2)
  } catch {
    return truncateText(value)
  }
}

function buildArchiveItemId(bundlePath: string, entryChain: string[]): string {
  return createHash('sha1')
    .update(normalizeText(bundlePath))
    .update('\u0000')
    .update(entryChain.map((segment) => normalizeText(segment)).join('\u0000'))
    .digest('hex')
}

function buildPreviewFromBuffer(entryPath: string, buffer: Buffer): {
  previewKind: ArchivePreviewKind
  previewText: string
  previewHtml: string
  searchText: string
} {
  const contentType = inferContentType(entryPath)
  const ext = path.extname(entryPath).toLowerCase()
  if (ext === '.docx' || ext === '.pptx' || ext === '.xlsx' || ext === '.doc' || ext === '.ppt' || ext === '.xls') {
    try {
      const zip = new AdmZip(buffer)
      const text = extractOfficeText(zip, entryPath)
      return {
        previewKind: text ? 'text' : 'binary',
        previewText: text,
        previewHtml: '',
        searchText: text
      }
    } catch {
      const text = truncateText(buffer.toString('utf8').slice(0, MAX_PREVIEW_BYTES))
      return {
        previewKind: text ? 'text' : 'binary',
        previewText: text,
        previewHtml: '',
        searchText: text
      }
    }
  }

  if (contentType.startsWith('text/html') || ext === '.html' || ext === '.htm') {
    const html = buffer.toString('utf8').slice(0, MAX_PREVIEW_BYTES)
    return {
      previewKind: 'html',
      previewText: stripHtml(html),
      previewHtml: html,
      searchText: stripHtml(html)
    }
  }

  if (
    contentType.startsWith('text/') ||
    ext === '.txt' ||
    ext === '.csv' ||
    ext === '.json' ||
    ext === '.xml' ||
    ext === '.ics'
  ) {
    const rawText = buffer.toString('utf8').slice(0, MAX_PREVIEW_BYTES)
    const text = ext === '.json' ? prettyPrintJson(rawText) : truncateText(rawText)
    return {
      previewKind: 'text',
      previewText: text,
      previewHtml: '',
      searchText: text
    }
  }

  return {
    previewKind: 'binary',
    previewText: '',
    previewHtml: '',
    searchText: ''
  }
}

async function walkArchiveEntries(
  bundlePath: string,
  bundleFileName: string,
  scopePath: string,
  scopeLabel: string,
  zipSource: any,
  parentChain: string[] = [],
  depth = 0,
  results: ArchiveBundleItem[] = []
): Promise<ArchiveBundleItem[]> {
  if (depth > MAX_BUNDLE_DEPTH) {
    return results
  }

  try {
    const entries = (await collectZipEntries(zipSource)).sort((left: any, right: any) =>
      left.fileName.localeCompare(right.fileName, undefined, { sensitivity: 'base' })
    )

    for (const entry of entries as any[]) {
      if (entry.isDirectory) {
        continue
      }

      const entryName = normalizeEntryPath(entry.fileName || entry.entryName || '')
      const chain = [...parentChain, entryName]
      const entryLeafName = getEntryName(entryName)
      const contentType = inferContentType(entryName)
      const modifiedAt = entry.getLastModDate ? entry.getLastModDate() : null
      const modifiedAtText = modifiedAt instanceof Date && !Number.isNaN(modifiedAt.valueOf()) ? modifiedAt.toISOString() : null

      if (isBundleFile(entryLeafName)) {
        const nestedBuffer = await readZipEntryBufferAsync(zipSource, entry)
        if (!nestedBuffer || nestedBuffer.length > MAX_ENTRY_BYTES) {
          continue
        }
        try {
          const nestedZip = await openZipBuffer(nestedBuffer)
          try {
            await walkArchiveEntries(
              bundlePath,
              bundleFileName,
              scopePath,
              scopeLabel,
              nestedZip,
              chain,
              depth + 1,
              results
            )
          } finally {
            try {
              nestedZip.close()
            } catch {
              // ignore
            }
          }
        } catch {
          continue
        }
        continue
      }

      if (!entry.fileName || entry.uncompressedSize > MAX_ENTRY_BYTES) {
        continue
      }

      const sourceType = classifySourceType(chain[0] || entryName)
      const buffer = await readZipEntryBufferAsync(zipSource, entry)
      if (!buffer) {
        continue
      }
      const { previewKind, previewText, previewHtml, searchText } = buildPreviewFromBuffer(entryName, buffer)
      const archiveItemId = buildArchiveItemId(bundlePath, chain)

      results.push({
        bundlePath,
        bundleFileName,
        scopePath,
        scopeLabel,
        sourceType,
        entryChain: chain,
        entryPath: chain.join('/'),
        entryName: entryLeafName,
        entrySize: entry.uncompressedSize || entry.compressedSize || buffer.length,
        modifiedAt: modifiedAtText,
        contentType,
        previewKind,
        previewText,
        previewHtml,
        downloadFilename: entryLeafName,
        searchText,
        archiveItemId
      })
    }
  } finally {
    try {
      zipSource.close()
    } catch {
      // ignore
    }
  }

  return results
}

export function listArchiveBundleFiles(
  rootPath: string,
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
      message: `Create a PST folder at ${resolvedRoot} and place Items*.zip bundles in it.`
    }
  }

  const scopes = buildScopeEntries(resolvedRoot)
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
        ? `Found ${selectedScope?.fileCount || 0} Items bundle${
            selectedScope && selectedScope.fileCount === 1 ? '' : 's'
          } in ${selectedScope?.scopeLabel || 'PST root'}.`
        : `No Items*.zip bundles were found anywhere under ${resolvedRoot}.`
  }
}

export async function extractArchiveBundleItems(
  bundlePath: string,
  scopePath: string,
  bundleFileName?: string
): Promise<ArchiveBundleItem[]> {
  const resolvedBundlePath = path.resolve(bundlePath)
  if (!fs.existsSync(resolvedBundlePath) || !fs.statSync(resolvedBundlePath).isFile()) {
    throw new Error(`Bundle not found: ${resolvedBundlePath}`)
  }

  const zip = await openZipFile(resolvedBundlePath)
  const scopeLabel = formatScopeLabel(scopePath)
  return walkArchiveEntries(
    resolvedBundlePath,
    bundleFileName || path.basename(resolvedBundlePath),
    normalizeScopePath(scopePath),
    scopeLabel,
    zip
  )
}

export async function readArchiveBundleItemContent(bundlePath: string, entryChain: string[]): Promise<{
  buffer: Buffer
  contentType: string
  fileName: string
}> {
  const chain = Array.isArray(entryChain) ? entryChain.map((value) => normalizeText(value)).filter(Boolean) : []
  if (!chain.length) {
    throw new Error('Archive item path is required')
  }
  const resolvedBundlePath = path.resolve(bundlePath)
  if (!fs.existsSync(resolvedBundlePath) || !fs.statSync(resolvedBundlePath).isFile()) {
    throw new Error(`Bundle not found: ${resolvedBundlePath}`)
  }

  let currentZip: any = await openZipFile(resolvedBundlePath)
  let buffer = Buffer.alloc(0)
  for (let index = 0; index < chain.length; index += 1) {
    const segment = chain[index]
    const entries = await collectZipEntries(currentZip)
    const entry = entries.find((candidate) => normalizeEntryPath(candidate.fileName || candidate.entryName || '') === segment)
    if (!entry) {
      throw new Error(`Archive entry not found: ${segment}`)
    }
    const nextBuffer = await readZipEntryBufferAsync(currentZip, entry)
    if (!nextBuffer) {
      throw new Error(`Unable to read archive entry: ${segment}`)
    }
    buffer = Buffer.from(nextBuffer as unknown as Uint8Array)
    if (index < chain.length - 1) {
      if (buffer.length > MAX_ENTRY_BYTES) {
        throw new Error('Nested archive entry is too large')
      }
      try {
        const nestedZip = await openZipBuffer(buffer)
        try {
          currentZip.close()
        } catch {
          // ignore
        }
        currentZip = nestedZip
      } catch (error) {
        throw new Error(
          error instanceof Error ? error.message : `Unable to open nested archive entry: ${segment}`
        )
      }
    }
  }

  try {
    currentZip.close()
  } catch {
    // ignore
  }

  return {
    buffer,
    contentType: inferContentType(chain[chain.length - 1]),
    fileName: getEntryName(chain[chain.length - 1])
  }
}
