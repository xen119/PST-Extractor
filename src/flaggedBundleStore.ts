import { GridFSBucket, MongoClient, ObjectId, type Collection } from 'mongodb'
import { Readable } from 'stream'

export type FlaggedBundleScope = 'all' | 'search' | 'pst'
export type FlaggedBundleGroupType = 'mailbox' | 'archive'
export type FlaggedBundleJobStatus = 'running' | 'succeeded' | 'failed'
export type FlaggedBundleJobStage =
  | 'collecting'
  | 'mailbox'
  | 'archive'
  | 'finalizing'
  | 'succeeded'
  | 'failed'

export interface FlaggedBundleExportScope {
  scope: FlaggedBundleScope
  scopePath: string
  scopeLabel: string
  sessionId: string
  sessionFileName: string
}

export interface FlaggedBundleProgress {
  stage: FlaggedBundleJobStage
  totalItems: number
  processedItems: number
  failedItems: number
  percent: number
  currentGroup: FlaggedBundleGroupType | null
  currentLabel: string
}

export interface FlaggedBundleArtifactRecord {
  artifactId: string
  fileId: string
  fileName: string
  partNumber: number
  partCount: number
  itemCount: number
  sizeBytes: number
  exceedsMaxSize: boolean
}

export interface FlaggedBundleGroupRecord {
  groupType: FlaggedBundleGroupType
  label: string
  itemCount: number
  failedCount: number
  artifacts: FlaggedBundleArtifactRecord[]
}

export interface FlaggedBundleJobRecord {
  exportId: string
  ownerUsername: string
  workspaceKey: string
  workspaceLockKey: string | null
  generatedAt: string
  startedAt: string
  completedAt: string | null
  updatedAt: string
  status: FlaggedBundleJobStatus
  scope: FlaggedBundleExportScope
  maxSizeBytes: number
  progress: FlaggedBundleProgress
  error: string | null
  groups: FlaggedBundleGroupRecord[]
}

export interface FlaggedBundleArtifactDownload {
  artifact: FlaggedBundleArtifactRecord
  stream: NodeJS.ReadableStream
}

export interface FlaggedBundleStore {
  kind: 'memory' | 'mongo'
  isPersistent: boolean
  createJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord>
  getJob(exportId: string): Promise<FlaggedBundleJobRecord | null>
  listJobsForWorkspace(ownerUsername: string, workspaceKey: string): Promise<FlaggedBundleJobRecord[]>
  saveJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord | null>
  addArtifact(
    exportId: string,
    groupType: FlaggedBundleGroupType,
    artifact: FlaggedBundleArtifactRecord,
    buffer: Buffer
  ): Promise<FlaggedBundleJobRecord | null>
  openArtifactDownload(exportId: string, artifactId: string): Promise<FlaggedBundleArtifactDownload | null>
  deleteJob(exportId: string): Promise<boolean>
  failRunningJobs(reason: string): Promise<number>
  close(): Promise<void>
}

export interface MongoFlaggedBundleStoreConnectOptions {
  bucketName?: string
}

const DEFAULT_COLLECTION_NAME = 'pst_flagged_bundle_jobs'

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function buildWorkspaceLockKey(ownerUsername: string, workspaceKey: string): string {
  return `${normalizeText(ownerUsername).toLowerCase()}::${normalizeText(workspaceKey)}`
}

function cloneProgress(progress: FlaggedBundleProgress): FlaggedBundleProgress {
  return { ...progress }
}

function cloneArtifact(artifact: FlaggedBundleArtifactRecord): FlaggedBundleArtifactRecord {
  return { ...artifact }
}

function cloneGroup(group: FlaggedBundleGroupRecord): FlaggedBundleGroupRecord {
  return {
    ...group,
    artifacts: group.artifacts.map(cloneArtifact)
  }
}

function cloneJob(job: FlaggedBundleJobRecord): FlaggedBundleJobRecord {
  return {
    ...job,
    scope: { ...job.scope },
    progress: cloneProgress(job.progress),
    groups: job.groups.map(cloneGroup)
  }
}

function normalizeGroupType(value: unknown): FlaggedBundleGroupType | null {
  if (value === 'mailbox' || value === 'archive') {
    return value
  }
  return null
}

function normalizeJobRecord(value: unknown): FlaggedBundleJobRecord | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const source = value as Partial<FlaggedBundleJobRecord> & {
    scope?: Partial<FlaggedBundleExportScope>
    progress?: Partial<FlaggedBundleProgress>
    groups?: Array<Partial<FlaggedBundleGroupRecord> & { artifacts?: Partial<FlaggedBundleArtifactRecord>[] }>
  }

  const exportId = normalizeText(source.exportId)
  const ownerUsername = normalizeText(source.ownerUsername)
  const workspaceKey = normalizeText(source.workspaceKey)
  const scope = (source.scope || {}) as Partial<FlaggedBundleExportScope>
  const progress = (source.progress || {}) as Partial<FlaggedBundleProgress>
  if (!exportId || !ownerUsername || !workspaceKey) {
    return null
  }

  return {
    exportId,
    ownerUsername,
    workspaceKey,
    workspaceLockKey: normalizeText(source.workspaceLockKey) || null,
    generatedAt: normalizeText(source.generatedAt),
    startedAt: normalizeText(source.startedAt),
    completedAt: normalizeText(source.completedAt) || null,
    updatedAt: normalizeText(source.updatedAt),
    status: source.status === 'succeeded' || source.status === 'failed' ? source.status : 'running',
    scope: {
      scope: scope.scope === 'search' || scope.scope === 'pst' ? scope.scope : 'all',
      scopePath: normalizeText(scope.scopePath),
      scopeLabel: normalizeText(scope.scopeLabel),
      sessionId: normalizeText(scope.sessionId),
      sessionFileName: normalizeText(scope.sessionFileName)
    },
    maxSizeBytes: Math.max(1, Math.floor(Number(source.maxSizeBytes) || 0)),
    progress: {
      stage:
        progress.stage === 'mailbox' ||
        progress.stage === 'archive' ||
        progress.stage === 'finalizing' ||
        progress.stage === 'succeeded' ||
        progress.stage === 'failed'
          ? progress.stage
          : 'collecting',
      totalItems: Math.max(0, Math.floor(Number(progress.totalItems) || 0)),
      processedItems: Math.max(0, Math.floor(Number(progress.processedItems) || 0)),
      failedItems: Math.max(0, Math.floor(Number(progress.failedItems) || 0)),
      percent: Math.max(0, Math.min(100, Math.floor(Number(progress.percent) || 0))),
      currentGroup:
        progress.currentGroup === 'mailbox' || progress.currentGroup === 'archive'
          ? progress.currentGroup
          : null,
      currentLabel: normalizeText(progress.currentLabel)
    },
    error: normalizeText(source.error) || null,
    groups: Array.isArray(source.groups)
      ? source.groups.map((group) => ({
          groupType: group.groupType === 'archive' ? 'archive' : 'mailbox',
          label: normalizeText(group.label),
          itemCount: Math.max(0, Math.floor(Number(group.itemCount) || 0)),
          failedCount: Math.max(0, Math.floor(Number(group.failedCount) || 0)),
          artifacts: Array.isArray(group.artifacts)
            ? group.artifacts
                .map((artifact) => {
                  const artifactId = normalizeText(artifact.artifactId)
                  const fileId = normalizeText(artifact.fileId)
                  const fileName = normalizeText(artifact.fileName)
                  if (!artifactId || !fileId || !fileName) {
                    return null
                  }

                  return {
                    artifactId,
                    fileId,
                    fileName,
                    partNumber: Math.max(1, Math.floor(Number(artifact.partNumber) || 0)),
                    partCount: Math.max(1, Math.floor(Number(artifact.partCount) || 0)),
                    itemCount: Math.max(0, Math.floor(Number(artifact.itemCount) || 0)),
                    sizeBytes: Math.max(0, Math.floor(Number(artifact.sizeBytes) || 0)),
                    exceedsMaxSize: Boolean(artifact.exceedsMaxSize)
                  }
                })
                .filter((artifact): artifact is FlaggedBundleArtifactRecord => Boolean(artifact))
            : []
        }))
      : []
  }
}

function createJobProgress(totalItems = 0): FlaggedBundleProgress {
  return {
    stage: 'collecting',
    totalItems: Math.max(0, Math.floor(totalItems)),
    processedItems: 0,
    failedItems: 0,
    percent: 0,
    currentGroup: null,
    currentLabel: ''
  }
}

function updateProgress(
  progress: FlaggedBundleProgress,
  patch: Partial<FlaggedBundleProgress>
): FlaggedBundleProgress {
  const totalItems = Math.max(0, Math.floor(patch.totalItems ?? progress.totalItems))
  const processedItems = Math.max(0, Math.floor(patch.processedItems ?? progress.processedItems))
  const failedItems = Math.max(0, Math.floor(patch.failedItems ?? progress.failedItems))
  const percent = totalItems > 0 ? Math.min(100, Math.floor((processedItems / totalItems) * 100)) : 0
  return {
    ...progress,
    ...patch,
    totalItems,
    processedItems,
    failedItems,
    percent,
    currentGroup: patch.currentGroup === undefined ? progress.currentGroup : patch.currentGroup,
    currentLabel: patch.currentLabel === undefined ? progress.currentLabel : normalizeText(patch.currentLabel)
  }
}

function normalizeArtifactBuffer(content: Buffer): Buffer {
  return Buffer.isBuffer(content) ? Buffer.from(content) : Buffer.from(content || [])
}

class FlaggedBundleJobInProgressError extends Error {
  statusCode = 409

  constructor(ownerUsername: string, workspaceKey: string) {
    super(`Flagged bundle export already in progress for ${normalizeText(ownerUsername)} / ${normalizeText(workspaceKey)}`)
  }
}

class MemoryFlaggedBundleStore implements FlaggedBundleStore {
  kind: 'memory' = 'memory'
  isPersistent = false
  private readonly jobs = new Map<string, FlaggedBundleJobRecord>()
  private readonly artifacts = new Map<string, Map<string, Buffer>>()

  async createJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord> {
    const lockKey = buildWorkspaceLockKey(job.ownerUsername, job.workspaceKey)
    for (const existing of this.jobs.values()) {
      if (
        existing.status === 'running' &&
        existing.ownerUsername === job.ownerUsername &&
        existing.workspaceKey === job.workspaceKey
      ) {
        throw new FlaggedBundleJobInProgressError(job.ownerUsername, job.workspaceKey)
      }
    }

    const next = cloneJob(job)
    this.jobs.set(next.exportId, next)
    return cloneJob(next)
  }

  async getJob(exportId: string): Promise<FlaggedBundleJobRecord | null> {
    const job = this.jobs.get(normalizeText(exportId))
    return job ? cloneJob(job) : null
  }

  async listJobsForWorkspace(ownerUsername: string, workspaceKey: string): Promise<FlaggedBundleJobRecord[]> {
    const lockKey = buildWorkspaceLockKey(ownerUsername, workspaceKey)
    return [...this.jobs.values()]
      .filter((job) => job.workspaceLockKey === lockKey || (job.ownerUsername === ownerUsername && job.workspaceKey === workspaceKey))
      .sort((left, right) => Date.parse(right.generatedAt) - Date.parse(left.generatedAt))
      .map(cloneJob)
  }

  async saveJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord | null> {
    const existing = this.jobs.get(job.exportId)
    if (!existing) {
      return null
    }

    this.jobs.set(job.exportId, cloneJob(job))
    return cloneJob(job)
  }

  async addArtifact(
    exportId: string,
    groupType: FlaggedBundleGroupType,
    artifact: FlaggedBundleArtifactRecord,
    buffer: Buffer
  ): Promise<FlaggedBundleJobRecord | null> {
    const job = this.jobs.get(normalizeText(exportId))
    if (!job) {
      return null
    }

    const group = job.groups.find((item) => item.groupType === groupType)
    if (!group) {
      return null
    }

    const storedArtifact = cloneArtifact(artifact)
    group.artifacts.push(storedArtifact)
    let artifactMap = this.artifacts.get(job.exportId)
    if (!artifactMap) {
      artifactMap = new Map<string, Buffer>()
      this.artifacts.set(job.exportId, artifactMap)
    }
    artifactMap.set(storedArtifact.artifactId, normalizeArtifactBuffer(buffer))
    job.updatedAt = new Date().toISOString()
    this.jobs.set(job.exportId, cloneJob(job))
    return cloneJob(job)
  }

  async openArtifactDownload(exportId: string, artifactId: string): Promise<FlaggedBundleArtifactDownload | null> {
    const job = this.jobs.get(normalizeText(exportId))
    if (!job) {
      return null
    }

    for (const group of job.groups) {
      const artifact = group.artifacts.find((entry) => entry.artifactId === artifactId)
      if (!artifact) {
        continue
      }
      const artifactMap = this.artifacts.get(job.exportId)
      const buffer = artifactMap?.get(artifact.artifactId) || null
      if (!buffer) {
        return null
      }
      return {
        artifact: cloneArtifact(artifact),
        stream: Readable.from(buffer)
      }
    }

    return null
  }

  async deleteJob(exportId: string): Promise<boolean> {
    const key = normalizeText(exportId)
    const removed = this.jobs.delete(key)
    this.artifacts.delete(key)
    return removed
  }

  async failRunningJobs(reason: string): Promise<number> {
    let count = 0
    for (const [exportId, job] of this.jobs.entries()) {
      if (job.status !== 'running') {
        continue
      }
      job.status = 'failed'
      job.error = normalizeText(reason) || 'Flagged bundle export interrupted'
      job.completedAt = new Date().toISOString()
      job.updatedAt = job.completedAt
      job.workspaceLockKey = null
      job.progress = updateProgress(job.progress, {
        stage: 'failed',
        currentGroup: null,
        currentLabel: ''
      })
      this.jobs.set(exportId, cloneJob(job))
      count += 1
    }
    return count
  }

  async close(): Promise<void> {
    this.jobs.clear()
    this.artifacts.clear()
  }
}

interface StoredFlaggedBundleJobDocument extends FlaggedBundleJobRecord {}

function normalizeMongoError(error: unknown): Error {
  if (error instanceof Error) {
    if (error.message.includes('E11000')) {
      return new FlaggedBundleJobInProgressError('anonymous', 'workspace')
    }
    return error
  }
  return new Error(String(error))
}

export class MongoFlaggedBundleStore implements FlaggedBundleStore {
  kind: 'mongo' = 'mongo'
  isPersistent = true
  private readonly bucket: GridFSBucket

  private constructor(
    private readonly collection: Collection<StoredFlaggedBundleJobDocument>,
    private readonly client: MongoClient,
    bucketName = 'pst_flagged_bundle_exports'
  ) {
    this.bucket = new GridFSBucket(client.db(), { bucketName })
  }

  static async connect(
    uri: string,
    dbName = 'pst-extractor',
    options: MongoFlaggedBundleStoreConnectOptions = {}
  ): Promise<MongoFlaggedBundleStore> {
    const client = await MongoClient.connect(uri)
    const collection = client.db(dbName).collection<StoredFlaggedBundleJobDocument>(DEFAULT_COLLECTION_NAME)
    await collection.createIndex({ exportId: 1 }, { unique: true, name: 'flagged_bundle_export_id' })
    await collection.createIndex({ ownerUsername: 1, workspaceKey: 1, generatedAt: -1 }, { name: 'flagged_bundle_workspace' })
    await collection.createIndex(
      { workspaceLockKey: 1 },
      {
        unique: true,
        name: 'flagged_bundle_running_lock',
        partialFilterExpression: { workspaceLockKey: { $type: 'string' } }
      }
    )
    const store = new MongoFlaggedBundleStore(collection, client, options.bucketName)
    await store.failRunningJobs('Flagged bundle export interrupted by server restart')
    return store
  }

  private async readJob(exportId: string): Promise<FlaggedBundleJobRecord | null> {
    const job = normalizeJobRecord(await this.collection.findOne({ exportId: normalizeText(exportId) }))
    return job ? cloneJob(job) : null
  }

  async createJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord> {
    const normalized = cloneJob(job)
    const runningExisting = await this.collection.findOne({
      ownerUsername: normalized.ownerUsername,
      workspaceKey: normalized.workspaceKey,
      status: 'running'
    })
    if (runningExisting) {
      throw new FlaggedBundleJobInProgressError(normalized.ownerUsername, normalized.workspaceKey)
    }
    try {
      await this.collection.insertOne(normalized)
      return cloneJob(normalized)
    } catch (error) {
      const normalizedError = normalizeMongoError(error)
      if (normalizedError instanceof FlaggedBundleJobInProgressError) {
        throw new FlaggedBundleJobInProgressError(job.ownerUsername, job.workspaceKey)
      }
      throw normalizedError
    }
  }

  async getJob(exportId: string): Promise<FlaggedBundleJobRecord | null> {
    return this.readJob(exportId)
  }

  async listJobsForWorkspace(ownerUsername: string, workspaceKey: string): Promise<FlaggedBundleJobRecord[]> {
    const jobs = await this.collection
      .find({
        ownerUsername: normalizeText(ownerUsername),
        workspaceKey: normalizeText(workspaceKey)
      })
      .sort({ generatedAt: -1, exportId: -1 })
      .toArray()

    return jobs.map((job) => cloneJob(normalizeJobRecord(job) || job))
  }

  async saveJob(job: FlaggedBundleJobRecord): Promise<FlaggedBundleJobRecord | null> {
    const normalized = cloneJob(job)
    const result = await this.collection.replaceOne({ exportId: normalized.exportId }, normalized)
    if (!result.matchedCount) {
      return null
    }
    return cloneJob(normalized)
  }

  async addArtifact(
    exportId: string,
    groupType: FlaggedBundleGroupType,
    artifact: FlaggedBundleArtifactRecord,
    buffer: Buffer
  ): Promise<FlaggedBundleJobRecord | null> {
    const job = await this.readJob(exportId)
    if (!job) {
      return null
    }

    const group = job.groups.find((entry) => entry.groupType === groupType)
    if (!group) {
      return null
    }

    const uploadId = new ObjectId()
    await new Promise<void>((resolve, reject) => {
      const upload = this.bucket.openUploadStreamWithId(uploadId, artifact.fileName, {
        metadata: {
          exportId: job.exportId,
          artifactId: artifact.artifactId,
          groupType,
          partNumber: artifact.partNumber,
          partCount: artifact.partCount,
          itemCount: artifact.itemCount,
          sizeBytes: artifact.sizeBytes,
          exceedsMaxSize: artifact.exceedsMaxSize
        }
      })
      upload.once('error', reject)
      upload.once('finish', () => resolve())
      upload.end(normalizeArtifactBuffer(buffer))
    })

    group.artifacts.push({
      ...cloneArtifact(artifact),
      fileId: uploadId.toHexString()
    })
    job.updatedAt = new Date().toISOString()
    const saved = await this.saveJob(job)
    if (!saved) {
      try {
        await this.bucket.delete(uploadId)
      } catch {
        // ignore orphan cleanup failures
      }
      return null
    }

    return saved
  }

  async openArtifactDownload(exportId: string, artifactId: string): Promise<FlaggedBundleArtifactDownload | null> {
    const job = await this.readJob(exportId)
    if (!job) {
      return null
    }

    for (const group of job.groups) {
      const artifact = group.artifacts.find((entry) => entry.artifactId === artifactId)
      if (!artifact || !artifact.fileId) {
        continue
      }

      const fileId = new ObjectId(artifact.fileId)
      return {
        artifact: cloneArtifact(artifact),
        stream: this.bucket.openDownloadStream(fileId)
      }
    }

    return null
  }

  async deleteJob(exportId: string): Promise<boolean> {
    const job = await this.readJob(exportId)
    if (!job) {
      return false
    }

    await this.collection.deleteOne({ exportId: job.exportId })
    for (const group of job.groups) {
      for (const artifact of group.artifacts) {
        if (!artifact.fileId) {
          continue
        }
        try {
          await this.bucket.delete(new ObjectId(artifact.fileId))
        } catch {
          // Ignore missing blob cleanup failures.
        }
      }
    }
    return true
  }

  async failRunningJobs(reason: string): Promise<number> {
    const jobs = await this.collection.find({ status: 'running' }).toArray()
    let count = 0
    for (const job of jobs) {
      const normalized = normalizeJobRecord(job)
      if (!normalized || normalized.status !== 'running') {
        continue
      }
      normalized.status = 'failed'
      normalized.error = normalizeText(reason) || 'Flagged bundle export interrupted'
      normalized.completedAt = new Date().toISOString()
      normalized.updatedAt = normalized.completedAt
      normalized.workspaceLockKey = null
      normalized.progress = updateProgress(normalized.progress, {
        stage: 'failed',
        currentGroup: null,
        currentLabel: ''
      })
      await this.collection.replaceOne({ exportId: normalized.exportId }, normalized)
      count += 1
    }
    return count
  }

  async close(): Promise<void> {
    await this.client.close()
  }
}

export function createMemoryFlaggedBundleStore(): FlaggedBundleStore {
  return new MemoryFlaggedBundleStore()
}

export async function createFlaggedBundleStoreFromEnv(
  env: NodeJS.ProcessEnv = process.env,
  options: MongoFlaggedBundleStoreConnectOptions = {}
): Promise<FlaggedBundleStore> {
  const uri = normalizeText(env.MONGODB_URI)
  if (!uri) {
    return createMemoryFlaggedBundleStore()
  }

  const dbName = normalizeText(env.MONGODB_DB) || 'pst-extractor'
  return MongoFlaggedBundleStore.connect(uri, dbName, options)
}

export function createFlaggedBundleJobInProgressError(
  ownerUsername: string,
  workspaceKey: string
): Error & { statusCode: number } {
  const error = new FlaggedBundleJobInProgressError(ownerUsername, workspaceKey) as Error & { statusCode: number }
  error.statusCode = 409
  return error
}

export function buildFlaggedBundleWorkspaceKey(
  scope: FlaggedBundleScope,
  scopePath: string,
  sessionId: string
): string {
  const normalizedScope = scope === 'search' || scope === 'pst' ? scope : 'all'
  return [normalizedScope, normalizeText(scopePath), normalizeText(sessionId)].join('::')
}
