import * as fs from 'fs'
import * as path from 'path'

export type AuditOutcome = 'success' | 'failure' | 'denied'

export interface AuditActor {
  username: string
  authenticated: boolean
  admin: boolean
}

export interface AuditRequestInfo {
  method: string
  path: string
  origin: string
  ip: string
}

export interface AuditLogEntry {
  timestamp: string
  actor: AuditActor
  action: string
  target: string
  outcome: AuditOutcome
  request: AuditRequestInfo
  metadata: Record<string, unknown>
}

export interface AuditLogStore {
  append(entry: AuditLogEntry): Promise<void>
  listRecent(limit?: number, actorUsername?: string): Promise<AuditLogEntry[]>
  listAll(actorUsername?: string): Promise<AuditLogEntry[]>
  close(): Promise<void>
}

const DEFAULT_AUDIT_LOG_FILE = 'activity.log'
const DEFAULT_RECENT_LIMIT = 50
const MAX_RECENT_LIMIT = 200

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeOutcome(value: unknown): AuditOutcome {
  const text = normalizeText(value).toLowerCase()
  if (text === 'failure') {
    return 'failure'
  }
  if (text === 'denied') {
    return 'denied'
  }
  return 'success'
}

function normalizeAuditActor(value: unknown): AuditActor {
  if (!value || typeof value !== 'object') {
    return {
      username: 'anonymous',
      authenticated: false,
      admin: false
    }
  }

  const actor = value as Partial<AuditActor>
  return {
    username: normalizeText(actor.username) || 'anonymous',
    authenticated: Boolean(actor.authenticated),
    admin: Boolean(actor.admin)
  }
}

function normalizeAuditRequest(value: unknown): AuditRequestInfo {
  if (!value || typeof value !== 'object') {
    return {
      method: '',
      path: '',
      origin: '',
      ip: ''
    }
  }

  const request = value as Partial<AuditRequestInfo>
  return {
    method: normalizeText(request.method),
    path: normalizeText(request.path),
    origin: normalizeText(request.origin),
    ip: normalizeText(request.ip)
  }
}

function normalizeAuditMetadata(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {}
  }

  const metadata: Record<string, unknown> = {}
  for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
    if (entry === undefined) {
      continue
    }
    metadata[key] = entry
  }
  return metadata
}

function normalizeAuditEntry(value: unknown): AuditLogEntry | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const entry = value as Partial<AuditLogEntry>
  const timestamp = normalizeText(entry.timestamp)
  const action = normalizeText(entry.action)
  const target = normalizeText(entry.target)

  if (!timestamp || !action) {
    return null
  }

  return {
    timestamp,
    actor: normalizeAuditActor(entry.actor),
    action,
    target,
    outcome: normalizeOutcome(entry.outcome),
    request: normalizeAuditRequest(entry.request),
    metadata: normalizeAuditMetadata(entry.metadata)
  }
}

function normalizeRecentLimit(limit: number | undefined): number {
  if (!Number.isFinite(limit)) {
    return DEFAULT_RECENT_LIMIT
  }

  return Math.max(1, Math.min(MAX_RECENT_LIMIT, Math.floor(Number(limit))))
}

function normalizeUsernameFilter(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

export class FileAuditLogStore implements AuditLogStore {
  private readonly filePath: string

  constructor(logDir: string, fileName = DEFAULT_AUDIT_LOG_FILE) {
    const normalizedDir = String(logDir ?? '').trim()
    if (!normalizedDir) {
      throw new Error('Audit log directory is required')
    }

    const absoluteDir = path.resolve(normalizedDir)
    fs.mkdirSync(absoluteDir, { recursive: true })
    this.filePath = path.join(absoluteDir, fileName)
  }

  async append(entry: AuditLogEntry): Promise<void> {
    const normalizedEntry = normalizeAuditEntry(entry)
    if (!normalizedEntry) {
      return
    }

    fs.appendFileSync(this.filePath, `${JSON.stringify(normalizedEntry)}\n`, 'utf8')
  }

  private readEntries(): AuditLogEntry[] {
    if (!fs.existsSync(this.filePath)) {
      return []
    }

    const content = fs.readFileSync(this.filePath, 'utf8')
    if (!content.trim()) {
      return []
    }

    const entries: AuditLogEntry[] = []
    for (const line of content.split(/\r?\n/g)) {
      const text = normalizeText(line)
      if (!text) {
        continue
      }

      try {
        const parsed = JSON.parse(text)
        const normalized = normalizeAuditEntry(parsed)
        if (normalized) {
          entries.push(normalized)
        }
      } catch {
        continue
      }
    }

    return entries
  }

  async listRecent(limit = DEFAULT_RECENT_LIMIT, actorUsername = ''): Promise<AuditLogEntry[]> {
    const entries = this.readEntries()

    const recentLimit = normalizeRecentLimit(limit)
    const normalizedActorUsername = normalizeUsernameFilter(actorUsername)
    const filteredEntries = normalizedActorUsername
      ? entries.filter(
          (entry) => normalizeUsernameFilter(entry.actor.username) === normalizedActorUsername
        )
      : entries

    return filteredEntries.slice(Math.max(0, filteredEntries.length - recentLimit)).reverse()
  }

  async listAll(actorUsername = ''): Promise<AuditLogEntry[]> {
    const entries = this.readEntries()
    const normalizedActorUsername = normalizeUsernameFilter(actorUsername)
    const filteredEntries = normalizedActorUsername
      ? entries.filter(
          (entry) => normalizeUsernameFilter(entry.actor.username) === normalizedActorUsername
        )
      : entries

    return filteredEntries.slice().reverse()
  }

  async close(): Promise<void> {}
}

export function createFileAuditLogStore(logDir: string, fileName = DEFAULT_AUDIT_LOG_FILE): AuditLogStore {
  return new FileAuditLogStore(logDir, fileName)
}
