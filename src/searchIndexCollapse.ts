import { fork, type ChildProcess } from 'child_process'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import {
  buildMailboxCollapsePartitionMetadata,
  CURRENT_MAILBOX_COLLAPSE_VERSION,
  getMailboxCollapsePartitionKeys,
  type HiddenRuleRecord,
  type MailboxCollapseJobStatus,
  type SearchIndexDocument,
  type SearchIndexStore
} from './searchIndex'

export interface MailboxCollapseCoordinator {
  start(): Promise<MailboxCollapseJobStatus>
  ensureStarted(): Promise<MailboxCollapseJobStatus>
  getStatus(): Promise<MailboxCollapseJobStatus>
  reset(): Promise<MailboxCollapseJobStatus>
}

export interface MailboxCollapseCoordinatorOptions {
  searchIndexStore: SearchIndexStore
}

export interface MailboxCollapseWorkerOptions {
  jobId?: string
  startedAt?: string
  resetBefore?: boolean
  onProgress?: (status: MailboxCollapseJobStatus) => Promise<void> | void
}

function nowIso(): string {
  return new Date().toISOString()
}

function buildIdleStatus(): MailboxCollapseJobStatus {
  return {
    jobId: null,
    status: 'idle',
    version: CURRENT_MAILBOX_COLLAPSE_VERSION,
    startedAt: null,
    completedAt: null,
    updatedAt: nowIso(),
    processedPartitions: 0,
    totalPartitions: 0,
    completedPartitionKeys: [],
    processedWorkUnits: 0,
    totalWorkUnits: 0,
    percentage: 0,
    provisional: true,
    error: null,
    reindexRequired: false
  }
}

function cloneStatus(status: MailboxCollapseJobStatus): MailboxCollapseJobStatus {
  return {
    jobId: status.jobId || null,
    status: status.status,
    version: status.version,
    startedAt: status.startedAt || null,
    completedAt: status.completedAt || null,
    updatedAt: status.updatedAt,
    processedPartitions: status.processedPartitions,
    totalPartitions: status.totalPartitions,
    completedPartitionKeys: [...(status.completedPartitionKeys || [])],
    processedWorkUnits: status.processedWorkUnits,
    totalWorkUnits: status.totalWorkUnits,
    percentage: status.percentage,
    provisional: status.provisional,
    error: status.error || null,
    reindexRequired: status.reindexRequired
  }
}

function percentage(processed: number, total: number): number {
  if (!total) {
    return 100
  }
  return Math.max(0, Math.min(100, Math.round((processed / total) * 100)))
}

function partitionDocuments(documents: SearchIndexDocument[]): Map<string, SearchIndexDocument[]> {
  const partitions = new Map<string, SearchIndexDocument[]>()
  for (const document of documents) {
    for (const partitionKey of getMailboxCollapsePartitionKeys(document)) {
      const records = partitions.get(partitionKey) || []
      records.push(document)
      partitions.set(partitionKey, records)
    }
  }
  return new Map(
    [...partitions.entries()].sort(([left], [right]) => left.localeCompare(right))
  )
}

function requiresReindex(documents: SearchIndexDocument[]): boolean {
  return documents.some((document) => {
    if (document.kind !== 'mail' && document.kind !== 'appointment') {
      return false
    }
    return !document.threadMetadata || typeof document.threadMetadata.messageId !== 'string'
  })
}

async function saveStatus(
  store: SearchIndexStore,
  status: MailboxCollapseJobStatus,
  onProgress?: MailboxCollapseWorkerOptions['onProgress']
): Promise<void> {
  await store.saveMailboxCollapseJob?.(status)
  await onProgress?.(cloneStatus(status))
}

export async function runMailboxCollapseJob(
  store: SearchIndexStore,
  options: MailboxCollapseWorkerOptions = {}
): Promise<MailboxCollapseJobStatus> {
  if (!store.getMailboxCollapseDocuments || !store.writeMailboxCollapsePartition || !store.finalizeMailboxCollapseMetadata) {
    const status = {
      ...buildIdleStatus(),
      status: 'reindex-required' as const,
      error: 'Mailbox collapse storage is unavailable. Reindex the mailbox search index.',
      reindexRequired: true,
      completedAt: nowIso()
    }
    await saveStatus(store, status, options.onProgress)
    return status
  }

  if (options.resetBefore) {
    await store.resetMailboxCollapseMetadata?.()
  }
  const documents = await store.getMailboxCollapseDocuments()
  const existing = (await store.getMailboxCollapseJob?.()) || null
  const partitions = partitionDocuments(documents)
  const totalWorkUnits = [...partitions.values()].reduce((sum, records) => sum + records.length, 0)
  const completed = new Set(
    (existing?.completedPartitionKeys || []).filter((partitionKey) => partitions.has(partitionKey))
  )
  const processedWorkUnits = [...partitions.entries()]
    .filter(([key]) => completed.has(key))
    .reduce((sum, [, records]) => sum + records.length, 0)
  const jobId = options.jobId || existing?.jobId || randomBytes(8).toString('hex')
  const startedAt = options.startedAt || existing?.startedAt || nowIso()
  let status: MailboxCollapseJobStatus = {
    ...(existing || buildIdleStatus()),
    jobId,
    status: 'running',
    version: CURRENT_MAILBOX_COLLAPSE_VERSION,
    startedAt,
    completedAt: null,
    updatedAt: nowIso(),
    processedPartitions: completed.size,
    totalPartitions: partitions.size,
    completedPartitionKeys: [...completed],
    processedWorkUnits,
    totalWorkUnits,
    percentage: percentage(processedWorkUnits, totalWorkUnits),
    provisional: true,
    error: null,
    reindexRequired: false
  }
  await saveStatus(store, status, options.onProgress)

  if (requiresReindex(documents)) {
    status = {
      ...status,
      status: 'reindex-required',
      error: 'Mailbox relationship metadata is missing. Reindex mailboxes before grouping items.',
      reindexRequired: true,
      completedAt: nowIso(),
      updatedAt: nowIso()
    }
    await saveStatus(store, status, options.onProgress)
    return status
  }

  const hiddenRules: HiddenRuleRecord[] = await store.listHiddenRules()
  try {
    for (const [partitionKey, partitionRecords] of partitions) {
      if (completed.has(partitionKey)) {
        continue
      }
      const references = buildMailboxCollapsePartitionMetadata(
        partitionKey,
        partitionRecords,
        hiddenRules
      )
      await store.writeMailboxCollapsePartition(partitionKey, partitionRecords, references)
      completed.add(partitionKey)
      status = {
        ...status,
        processedPartitions: completed.size,
        completedPartitionKeys: [...completed],
        processedWorkUnits: status.processedWorkUnits + partitionRecords.length,
        percentage: percentage(status.processedWorkUnits + partitionRecords.length, totalWorkUnits),
        updatedAt: nowIso()
      }
      await saveStatus(store, status, options.onProgress)
      await new Promise<void>((resolve) => setImmediate(resolve))
    }

    await store.finalizeMailboxCollapseMetadata(CURRENT_MAILBOX_COLLAPSE_VERSION)
    status = {
      ...status,
      status: 'succeeded',
      processedPartitions: partitions.size,
      totalPartitions: partitions.size,
      processedWorkUnits: totalWorkUnits,
      totalWorkUnits,
      percentage: 100,
      provisional: false,
      completedAt: nowIso(),
      updatedAt: nowIso(),
      error: null,
      reindexRequired: false
    }
    await saveStatus(store, status, options.onProgress)
    return status
  } catch (error) {
    status = {
      ...status,
      status: 'failed',
      updatedAt: nowIso(),
      completedAt: nowIso(),
      error: error instanceof Error ? error.message : String(error),
      provisional: true
    }
    await saveStatus(store, status, options.onProgress)
    return status
  }
}

function resolveWorkerScriptPath(): string {
  const candidates = [
    path.resolve(__dirname, 'searchIndexCollapseWorker.js'),
    path.resolve(__dirname, 'searchIndexCollapseWorker.ts'),
    path.resolve(process.cwd(), 'src', 'searchIndexCollapseWorker.ts'),
    path.resolve(process.cwd(), 'dist', 'searchIndexCollapseWorker.js')
  ]
  return candidates.find((candidate) => fs.existsSync(candidate)) || candidates[candidates.length - 1]
}

export function createMailboxCollapseCoordinator(
  options: MailboxCollapseCoordinatorOptions
): MailboxCollapseCoordinator {
  const store = options.searchIndexStore
  let activePromise: Promise<MailboxCollapseJobStatus> | null = null
  let activeWorker: ChildProcess | null = null

  function shouldUseWorker(): boolean {
    return store.kind === 'mongo' && store.isPersistent && Boolean(String(process.env.MONGODB_URI || '').trim())
  }

  async function getStatus(): Promise<MailboxCollapseJobStatus> {
    const status = await store.getMailboxCollapseJob?.()
    return cloneStatus(status || buildIdleStatus())
  }

  async function run(resetBefore = false): Promise<MailboxCollapseJobStatus> {
    if (!shouldUseWorker()) {
      return runMailboxCollapseJob(store)
    }

    return new Promise<MailboxCollapseJobStatus>((resolve) => {
      const worker = fork(resolveWorkerScriptPath(), [], {
        env: {
          ...process.env,
          PST_SEARCH_INDEX_COLLAPSE_JOB_ID: (store.getMailboxCollapseJob
            ? ''
            : randomBytes(8).toString('hex')),
          PST_SEARCH_INDEX_COLLAPSE_RESET: resetBefore ? '1' : '0'
        },
        stdio: ['inherit', 'inherit', 'inherit', 'ipc'],
        execArgv: [...process.execArgv, '-r', 'ts-node/register/transpile-only']
      })
      activeWorker = worker
      let settled = false
      const finish = async (): Promise<void> => {
        if (settled) {
          return
        }
        settled = true
        activeWorker = null
        resolve(await getStatus())
      }
      worker.on('message', (message: { type?: string }) => {
        if (message?.type === 'success' || message?.type === 'failure') {
          void finish()
        }
      })
      worker.once('error', () => void finish())
      worker.once('exit', () => void finish())
    })
  }

  async function start(): Promise<MailboxCollapseJobStatus> {
    if (activePromise) {
      return getStatus()
    }
    const current = await getStatus()
    if (current.status === 'succeeded' && current.version === CURRENT_MAILBOX_COLLAPSE_VERSION) {
      return current
    }
    if (current.status === 'reindex-required') {
      return current
    }
    const resetBefore = current.status === 'idle'
    const jobId = current.jobId || randomBytes(8).toString('hex')
    const startedAt = current.startedAt || nowIso()
    const running: MailboxCollapseJobStatus = {
      ...current,
      jobId,
      status: 'running',
      version: CURRENT_MAILBOX_COLLAPSE_VERSION,
      startedAt,
      completedAt: null,
      updatedAt: nowIso(),
      error: null,
      reindexRequired: false,
      provisional: true
    }
    await store.saveMailboxCollapseJob?.(running)
    activePromise = shouldUseWorker()
      ? run(resetBefore)
      : runMailboxCollapseJob(store, { jobId, startedAt, resetBefore })
    void activePromise.finally(() => {
      activePromise = null
    })
    return running
  }

  async function ensureStarted(): Promise<MailboxCollapseJobStatus> {
    const current = await getStatus()
    if (current.status === 'succeeded' || current.status === 'reindex-required') {
      return current
    }
    return start()
  }

  async function reset(): Promise<MailboxCollapseJobStatus> {
    if (activeWorker) {
      activeWorker.kill()
      activeWorker = null
    }
    activePromise = null
    await store.resetMailboxCollapseMetadata?.()
    const status = buildIdleStatus()
    await store.saveMailboxCollapseJob?.(status)
    return status
  }

  return { start, ensureStarted, getStatus, reset }
}
