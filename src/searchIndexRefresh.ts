import { fork, type ChildProcess } from 'child_process'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import type { ReviewStore } from './reviewStore'
import type { SearchIndexStore } from './searchIndex'

export type SearchIndexRefreshTrigger = 'startup' | 'manual'
export type SearchIndexRefreshState = 'idle' | 'running' | 'succeeded' | 'failed'

export interface SearchIndexRefreshSummary {
  mailboxCount: number
  messageCount: number
}

export interface SearchIndexRefreshStatus {
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
  start(trigger: SearchIndexRefreshTrigger): Promise<SearchIndexRefreshStatus>
  getStatus(): SearchIndexRefreshStatus
}

export interface SearchIndexRefreshCoordinatorOptions {
  pstRootDir: string
  reviewStore: ReviewStore
  searchIndexStore: SearchIndexStore
  onJobComplete?: (status: SearchIndexRefreshStatus) => void
}

interface RefreshJobContext {
  jobId: string
  trigger: SearchIndexRefreshTrigger
  startedAt: string
  stagingDocumentsCollectionName: string
}

interface WorkerMessage {
  type: 'success' | 'failure'
  summary?: SearchIndexRefreshSummary
  error?: string
}

class SearchIndexRefreshInProgressError extends Error {
  statusCode = 409

  constructor() {
    super('Search index refresh already in progress')
  }
}

function normalizeText(value: unknown): string {
  return String(value ?? '').trim()
}

function nowIso(): string {
  return new Date().toISOString()
}

function buildInitialStatus(): SearchIndexRefreshStatus {
  return {
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

function createStagingDocumentsCollectionName(jobId: string): string {
  return `pst_search_documents_staging_${jobId}`
}

function isMongoSearchIndexStore(store: SearchIndexStore): store is SearchIndexStore & {
  kind: 'mongo'
  promoteStagedDocuments?: (stagingDocumentsCollectionName: string) => Promise<void>
} {
  return store.kind === 'mongo'
}

function isPersistentReviewStore(store: ReviewStore): boolean {
  return store.isPersistent
}

function createWorkerTransport(
  context: RefreshJobContext,
  options: SearchIndexRefreshCoordinatorOptions
): Promise<SearchIndexRefreshSummary> {
  return new Promise((resolve, reject) => {
    const workerEnv: NodeJS.ProcessEnv = {
      ...process.env,
      PST_SEARCH_INDEX_REFRESH_JOB_ID: context.jobId,
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

    const finish = (error: Error | null, summary?: SearchIndexRefreshSummary) => {
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
        summary || {
          mailboxCount: 0,
          messageCount: 0
        }
      )
    }

    worker.on('message', (message: WorkerMessage) => {
      if (!message || typeof message !== 'object') {
        return
      }
      if (message.type === 'success') {
        finish(null, message.summary)
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
  let status = buildInitialStatus()
  let activeJob: Promise<void> | null = null

  async function promoteStagedSnapshot(stagingDocumentsCollectionName: string): Promise<void> {
    if (!isMongoSearchIndexStore(options.searchIndexStore)) {
      return
    }

    if (typeof options.searchIndexStore.promoteStagedDocuments === 'function') {
      await options.searchIndexStore.promoteStagedDocuments(stagingDocumentsCollectionName)
    }
  }

  async function runInProcessRefresh(context: RefreshJobContext): Promise<SearchIndexRefreshSummary> {
    const { refreshSearchIndexFromCatalog } = await import('./searchIndex')
    return refreshSearchIndexFromCatalog(options.pstRootDir, options.reviewStore, options.searchIndexStore)
  }

  async function executeJob(job: RefreshJobContext): Promise<void> {
    const shouldUseWorker = isMongoSearchIndexStore(options.searchIndexStore) && isPersistentReviewStore(options.reviewStore)
    const startedAt = job.startedAt

    try {
      const summary = shouldUseWorker
        ? await createWorkerTransport(job, options)
        : await runInProcessRefresh(job)

      if (shouldUseWorker) {
        await promoteStagedSnapshot(job.stagingDocumentsCollectionName)
      }

      status = {
        jobId: job.jobId,
        status: 'succeeded',
        trigger: job.trigger,
        startedAt,
        completedAt: nowIso(),
        updatedAt: nowIso(),
        summary,
        error: null
      }
      try {
        options.onJobComplete?.({ ...status })
      } catch {
        // Ignore audit callback failures.
      }
    } catch (error) {
      status = {
        jobId: job.jobId,
        status: 'failed',
        trigger: job.trigger,
        startedAt,
        completedAt: nowIso(),
        updatedAt: nowIso(),
        summary: null,
        error: error instanceof Error ? error.message : String(error)
      }
      try {
        options.onJobComplete?.({ ...status })
      } catch {
        // Ignore audit callback failures.
      }
    } finally {
      activeJob = null
    }
  }

  return {
    start(trigger: SearchIndexRefreshTrigger): Promise<SearchIndexRefreshStatus> {
      if (activeJob) {
        throw new SearchIndexRefreshInProgressError()
      }

      const job: RefreshJobContext = {
        jobId: randomBytes(8).toString('hex'),
        trigger,
        startedAt: nowIso(),
        stagingDocumentsCollectionName: ''
      }
      job.stagingDocumentsCollectionName = createStagingDocumentsCollectionName(job.jobId)
      status = {
        jobId: job.jobId,
        status: 'running',
        trigger,
        startedAt: job.startedAt,
        completedAt: null,
        updatedAt: job.startedAt,
        summary: null,
        error: null
      }
      activeJob = executeJob(job)
      void activeJob.catch(() => undefined)
      return Promise.resolve({ ...status })
    },

    getStatus(): SearchIndexRefreshStatus {
      return { ...status }
    }
  }
}
