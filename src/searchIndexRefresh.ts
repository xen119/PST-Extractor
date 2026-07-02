import { fork } from 'child_process'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import type { ReviewStore } from './reviewStore'
import type {
  SearchIndexRefreshPlan,
  SearchIndexRefreshSource,
  SearchIndexStore
} from './searchIndex'

export type SearchIndexRefreshTrigger = 'startup' | 'manual'
export type SearchIndexRefreshState = 'idle' | 'running' | 'succeeded' | 'failed'

export interface SearchIndexRefreshSummary {
  mailboxCount: number
  messageCount: number
  changedCount: number
  skippedCount: number
  removedCount: number
  failedCount: number
}

export interface SearchIndexRefreshStatus {
  source: SearchIndexRefreshSource
  jobId: string | null
  status: SearchIndexRefreshState
  trigger: SearchIndexRefreshTrigger | null
  startedAt: string | null
  completedAt: string | null
  updatedAt: string
  summary: SearchIndexRefreshSummary | null
  error: string | null
}

export interface SearchIndexRefreshCoordinator {
  start(source: SearchIndexRefreshSource, trigger: SearchIndexRefreshTrigger): Promise<SearchIndexRefreshStatus>
  getStatus(source: SearchIndexRefreshSource): SearchIndexRefreshStatus
}

export interface SearchIndexRefreshCoordinatorOptions {
  pstRootDir: string
  reviewStore: ReviewStore
  searchIndexStore: SearchIndexStore
  onJobComplete?: (status: SearchIndexRefreshStatus) => void
}

interface RefreshJobContext {
  source: SearchIndexRefreshSource
  jobId: string
  trigger: SearchIndexRefreshTrigger
  startedAt: string
  stagingDocumentsCollectionName: string
}

interface WorkerMessage {
  type: 'success' | 'failure'
  plan?: SearchIndexRefreshPlan
  error?: string
}

class SearchIndexRefreshInProgressError extends Error {
  statusCode = 409

  constructor(source: SearchIndexRefreshSource) {
    super(`Search index refresh already in progress for ${source}`)
  }
}

function normalizeText(value: unknown): string {
  return String(value ?? '').trim()
}

function normalizeRefreshSource(value: unknown): SearchIndexRefreshSource {
  return normalizeText(value).toLowerCase() === 'items' ? 'items' : 'mailboxes'
}

function nowIso(): string {
  return new Date().toISOString()
}

function buildInitialStatus(source: SearchIndexRefreshSource): SearchIndexRefreshStatus {
  return {
    source,
    jobId: null,
    status: 'idle',
    trigger: null,
    startedAt: null,
    completedAt: null,
    updatedAt: nowIso(),
    summary: null,
    error: null
  }
}

function resolveWorkerScriptPath(): string {
  const candidates = [
    path.resolve(__dirname, 'searchIndexRefreshWorker.js'),
    path.resolve(__dirname, 'searchIndexRefreshWorker.ts'),
    path.resolve(process.cwd(), 'src', 'searchIndexRefreshWorker.ts'),
    path.resolve(process.cwd(), 'dist', 'searchIndexRefreshWorker.js')
  ]

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate
    }
  }

  return candidates[candidates.length - 1]
}

function createStagingDocumentsCollectionName(source: SearchIndexRefreshSource, jobId: string): string {
  return `pst_search_documents_staging_${source}_${jobId}`
}

function isMongoSearchIndexStore(store: SearchIndexStore): store is SearchIndexStore & {
  kind: 'mongo'
  promoteStagedDocuments?: (
    stagingDocumentsCollectionName: string,
    changedMailboxKeys?: string[],
    removedMailboxKeys?: string[]
  ) => Promise<void>
} {
  return store.kind === 'mongo'
}

function isPersistentReviewStore(store: ReviewStore): boolean {
  return store.isPersistent
}

function buildSummaryFromPlan(plan: SearchIndexRefreshPlan): SearchIndexRefreshSummary {
  return {
    mailboxCount: plan.mailboxCount,
    messageCount: plan.messageCount,
    changedCount: plan.changedCount,
    skippedCount: plan.skippedCount,
    removedCount: plan.removedCount,
    failedCount: plan.failedCount
  }
}

function createWorkerTransport(
  context: RefreshJobContext,
  options: SearchIndexRefreshCoordinatorOptions
): Promise<SearchIndexRefreshPlan> {
  return new Promise((resolve, reject) => {
    const workerEnv: NodeJS.ProcessEnv = {
      ...process.env,
      PST_SEARCH_INDEX_REFRESH_JOB_ID: context.jobId,
      PST_SEARCH_INDEX_REFRESH_SOURCE: context.source,
      PST_SEARCH_INDEX_REFRESH_TRIGGER: context.trigger,
      PST_SEARCH_INDEX_PST_ROOT_DIR: options.pstRootDir,
      PST_SEARCH_INDEX_STAGING_DOCUMENTS_COLLECTION: context.stagingDocumentsCollectionName
    }

    const worker = fork(resolveWorkerScriptPath(), [], {
      env: workerEnv,
      stdio: ['inherit', 'inherit', 'inherit', 'ipc'],
      execArgv: [...process.execArgv, '-r', 'ts-node/register/transpile-only']
    })

    let settled = false

    const finish = (error: Error | null, plan?: SearchIndexRefreshPlan) => {
      if (settled) {
        return
      }
      settled = true
      worker.removeAllListeners()
      if (error) {
        reject(error)
        return
      }
      resolve(
        plan || {
          source: context.source,
          mailboxCount: 0,
          messageCount: 0,
          changedCount: 0,
          skippedCount: 0,
          removedCount: 0,
          failedCount: 0,
          changedMailboxKeys: [],
          removedMailboxKeys: [],
          fingerprints: []
        }
      )
    }

    worker.on('message', (message: WorkerMessage) => {
      if (!message || typeof message !== 'object') {
        return
      }
      if (message.type === 'success') {
        finish(null, message.plan)
        return
      }
      if (message.type === 'failure') {
        finish(new Error(normalizeText(message.error) || 'Search index refresh failed'))
      }
    })

    worker.once('error', (error) => {
      finish(error instanceof Error ? error : new Error(String(error)))
    })

    worker.once('exit', (code, signal) => {
      if (settled) {
        return
      }
      if (code === 0) {
        finish(null)
        return
      }
      finish(
        new Error(
          signal
            ? `Search index refresh worker exited with signal ${signal}`
            : `Search index refresh worker exited with code ${code}`
        )
      )
    })
  })
}

export function createSearchIndexRefreshCoordinator(
  options: SearchIndexRefreshCoordinatorOptions
): SearchIndexRefreshCoordinator {
  const statusBySource = new Map<SearchIndexRefreshSource, SearchIndexRefreshStatus>([
    ['mailboxes', buildInitialStatus('mailboxes')],
    ['items', buildInitialStatus('items')]
  ])
  const activeJobs = new Map<SearchIndexRefreshSource, Promise<void>>()

  async function promoteStagedSnapshot(
    stagingDocumentsCollectionName: string,
    changedMailboxKeys: string[],
    removedMailboxKeys: string[]
  ): Promise<void> {
    if (!isMongoSearchIndexStore(options.searchIndexStore)) {
      return
    }

    if (typeof options.searchIndexStore.promoteStagedDocuments === 'function') {
      await options.searchIndexStore.promoteStagedDocuments(
        stagingDocumentsCollectionName,
        changedMailboxKeys,
        removedMailboxKeys
      )
    }
  }

  async function runInProcessRefresh(context: RefreshJobContext): Promise<SearchIndexRefreshPlan> {
    const { refreshSearchIndexSourceFromCatalog } = await import('./searchIndex')
    return refreshSearchIndexSourceFromCatalog(
      options.pstRootDir,
      context.source,
      options.reviewStore,
      options.searchIndexStore,
      {
        pruneRemovedFiles: true,
        updateFingerprints: true
      }
    )
  }

  async function executeJob(job: RefreshJobContext): Promise<void> {
    const shouldUseWorker =
      isMongoSearchIndexStore(options.searchIndexStore) && isPersistentReviewStore(options.reviewStore)
    const startedAt = job.startedAt

    try {
      const plan = shouldUseWorker ? await createWorkerTransport(job, options) : await runInProcessRefresh(job)

      if (shouldUseWorker) {
        await promoteStagedSnapshot(job.stagingDocumentsCollectionName, plan.changedMailboxKeys, plan.removedMailboxKeys)
        await options.searchIndexStore.replaceFileFingerprints(job.source, plan.fingerprints)
      }

      statusBySource.set(job.source, {
        source: job.source,
        jobId: job.jobId,
        status: 'succeeded',
        trigger: job.trigger,
        startedAt,
        completedAt: nowIso(),
        updatedAt: nowIso(),
        summary: buildSummaryFromPlan(plan),
        error: null
      })
      try {
        options.onJobComplete?.({ ...statusBySource.get(job.source)! })
      } catch {
        // Ignore audit callback failures.
      }
    } catch (error) {
      statusBySource.set(job.source, {
        source: job.source,
        jobId: job.jobId,
        status: 'failed',
        trigger: job.trigger,
        startedAt,
        completedAt: nowIso(),
        updatedAt: nowIso(),
        summary: null,
        error: error instanceof Error ? error.message : String(error)
      })
      try {
        options.onJobComplete?.({ ...statusBySource.get(job.source)! })
      } catch {
        // Ignore audit callback failures.
      }
    } finally {
      activeJobs.delete(job.source)
    }
  }

  return {
    start(source: SearchIndexRefreshSource, trigger: SearchIndexRefreshTrigger): Promise<SearchIndexRefreshStatus> {
      const normalizedSource = normalizeRefreshSource(source)
      if (activeJobs.has(normalizedSource)) {
        throw new SearchIndexRefreshInProgressError(normalizedSource)
      }

      const job: RefreshJobContext = {
        source: normalizedSource,
        jobId: randomBytes(8).toString('hex'),
        trigger,
        startedAt: nowIso(),
        stagingDocumentsCollectionName: ''
      }
      job.stagingDocumentsCollectionName = createStagingDocumentsCollectionName(normalizedSource, job.jobId)
      statusBySource.set(normalizedSource, {
        source: normalizedSource,
        jobId: job.jobId,
        status: 'running',
        trigger,
        startedAt: job.startedAt,
        completedAt: null,
        updatedAt: job.startedAt,
        summary: null,
        error: null
      })
      const activeJob = executeJob(job)
      activeJobs.set(normalizedSource, activeJob)
      void activeJob.catch(() => undefined)
      return Promise.resolve({ ...statusBySource.get(normalizedSource)! })
    },

    getStatus(source: SearchIndexRefreshSource): SearchIndexRefreshStatus {
      const normalizedSource = normalizeRefreshSource(source)
      return { ...(statusBySource.get(normalizedSource) || buildInitialStatus(normalizedSource)) }
    }
  }
}
