import * as React from 'react'
import { AlertTriangle, Download, Loader2, RefreshCw, Trash2, X } from 'lucide-react'
import { api } from '@/api'
import type { FlaggedBundleJob, FlaggedBundleJobStatus, FlaggedBundleScope } from '@/types'
import { readWorkspaceStorageItem, writeWorkspaceStorageItem } from '@/lib/workspace'
import { cn, formatBytes, formatDate, normalizeText } from '@/lib/utils'
import {
  Badge,
  Button,
  Dialog,
  DialogContent,
  DialogDescription,
  DialogTitle,
  Input,
  ScrollArea,
  Separator
} from '@/components/ui'

const SIZE_PRESETS = [
  { value: '50', label: '50 MB', bytes: 50 * 1024 * 1024 },
  { value: '100', label: '100 MB', bytes: 100 * 1024 * 1024 },
  { value: '250', label: '250 MB', bytes: 250 * 1024 * 1024 },
  { value: '500', label: '500 MB', bytes: 500 * 1024 * 1024 },
  { value: 'custom', label: 'Custom', bytes: 0 }
] as const

type SizePresetValue = (typeof SIZE_PRESETS)[number]['value']

function isSizePresetValue(value: string): value is SizePresetValue {
  return SIZE_PRESETS.some((preset) => preset.value === value)
}

function getSizePresetBytes(preset: SizePresetValue, customMb: string): number {
  if (preset !== 'custom') {
    return SIZE_PRESETS.find((item) => item.value === preset)?.bytes || 250 * 1024 * 1024
  }

  const parsed = Number.parseFloat(normalizeText(customMb))
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return 250 * 1024 * 1024
  }

  return Math.max(1, Math.floor(parsed * 1024 * 1024))
}

function buildWorkspaceQuery(scope: FlaggedBundleDialogScope | null): {
  scope: FlaggedBundleScope
  scopePath?: string
  sessionId?: string
} | null {
  if (!scope) {
    return null
  }

  return {
    scope: scope.scope,
    scopePath: scope.scopePath || undefined,
    sessionId: scope.sessionId || undefined
  }
}

function isTerminalJobStatus(status: FlaggedBundleJobStatus): boolean {
  return status !== 'running'
}

function getJobStatusLabel(status: FlaggedBundleJobStatus): string {
  switch (status) {
    case 'running':
      return 'Running'
    case 'succeeded':
      return 'Ready'
    case 'failed':
      return 'Failed'
    default:
      return status
  }
}

function mergeJobList(jobs: FlaggedBundleJob[], nextJob: FlaggedBundleJob): FlaggedBundleJob[] {
  return [nextJob, ...jobs.filter((job) => job.exportId !== nextJob.exportId)].sort(
    (left, right) => Date.parse(right.generatedAt) - Date.parse(left.generatedAt)
  )
}

export interface FlaggedBundleDialogScope {
  scope: FlaggedBundleScope
  scopePath: string
  scopeLabel: string
  sessionId: string
  sessionFileName: string
}

export interface FlaggedBundleDialogProps {
  open: boolean
  authenticated: boolean
  username: string
  scope: FlaggedBundleDialogScope | null
  onOpenChange: (open: boolean) => void
}

export function FlaggedBundleDialog({
  open,
  authenticated,
  username,
  scope,
  onOpenChange
}: FlaggedBundleDialogProps) {
  const [preset, setPreset] = React.useState<SizePresetValue>('250')
  const [customMb, setCustomMb] = React.useState('250')
  const [busy, setBusy] = React.useState(false)
  const [loadingJobs, setLoadingJobs] = React.useState(false)
  const [deletingExportId, setDeletingExportId] = React.useState('')
  const [error, setError] = React.useState('')
  const [jobs, setJobs] = React.useState<FlaggedBundleJob[]>([])

  const workspaceQuery = React.useMemo(
    () => buildWorkspaceQuery(scope),
    [scope?.scope, scope?.scopePath, scope?.sessionId]
  )
  const currentScopeLabel = scope?.scopeLabel || 'All cases/searches'
  const currentScopeDescription = React.useMemo(() => {
    if (!scope) {
      return 'Select a case or search before generating downloads.'
    }
    if (scope.scope === 'pst') {
      return scope.sessionFileName || 'Selected PST'
    }
    return scope.scopePath || scope.scopeLabel || currentScopeLabel
  }, [currentScopeLabel, scope])

  const runningJob = React.useMemo(
    () => jobs.find((job) => job.status === 'running') || null,
    [jobs]
  )
  const hasRunningJob = Boolean(runningJob)

  React.useEffect(() => {
    if (!open) {
      setBusy(false)
      setLoadingJobs(false)
      setDeletingExportId('')
      setError('')
      setJobs([])
      return
    }

    const storedPreset = readWorkspaceStorageItem('flaggedBundleSizePreset', authenticated, username, '250')
    const storedCustom = readWorkspaceStorageItem('flaggedBundleCustomSizeMb', authenticated, username, '250')
    setPreset(isSizePresetValue(storedPreset) ? storedPreset : '250')
    setCustomMb(storedCustom || '250')
    setError('')
  }, [authenticated, open, username])

  React.useEffect(() => {
    if (!open || !workspaceQuery) {
      setJobs([])
      return
    }

    const query = workspaceQuery

    let cancelled = false

    async function refreshJobs(silent = false): Promise<void> {
      if (!silent) {
        setLoadingJobs(true)
      }

      try {
        const response = await api.exports.listFlaggedBundles(query)
        if (cancelled) {
          return
        }
        setJobs(response.jobs)
        if (!silent) {
          setError('')
        }
      } catch (bundleError) {
        if (cancelled) {
          return
        }
        if (!silent) {
          setError(bundleError instanceof Error ? bundleError.message : 'Unable to load bundle history')
        }
      } finally {
        if (!cancelled && !silent) {
          setLoadingJobs(false)
        }
      }
    }

    void refreshJobs(false)
    const timer = window.setInterval(() => {
      void refreshJobs(true)
    }, 5000)

    return () => {
      cancelled = true
      window.clearInterval(timer)
    }
  }, [open, scope?.scope, scope?.scopePath, scope?.sessionId, workspaceQuery])

  const sizeBytes = getSizePresetBytes(preset, customMb)

  async function handleGenerate(event: React.FormEvent): Promise<void> {
    event.preventDefault()
    if (!scope || busy || hasRunningJob) {
      return
    }

    setBusy(true)
    setError('')
    try {
      writeWorkspaceStorageItem('flaggedBundleSizePreset', authenticated, username, preset)
      writeWorkspaceStorageItem('flaggedBundleCustomSizeMb', authenticated, username, customMb)
      const response = await api.exports.prepareFlaggedBundle({
        scope: scope.scope,
        scopePath: scope.scopePath || undefined,
        sessionId: scope.sessionId || undefined,
        maxSizeBytes: sizeBytes
      })
      setJobs((previous) => mergeJobList(previous, response))
    } catch (bundleError) {
      setError(bundleError instanceof Error ? bundleError.message : 'Unable to start flagged bundle export')
    } finally {
      setBusy(false)
    }
  }

  async function handleDelete(exportId: string): Promise<void> {
    setDeletingExportId(exportId)
    setError('')
    try {
      await api.exports.deleteFlaggedBundle(exportId)
      setJobs((previous) => previous.filter((job) => job.exportId !== exportId))
    } catch (bundleError) {
      setError(bundleError instanceof Error ? bundleError.message : 'Unable to delete flagged bundle export')
    } finally {
      setDeletingExportId('')
    }
  }

  function renderArtifactCard(job: FlaggedBundleJob, artifact: FlaggedBundleJob['groups'][number]['artifacts'][number]) {
    return (
      <div
        key={artifact.artifactId}
        className={cn(
          'rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface)] p-4',
          artifact.exceedsMaxSize && 'border-[color:rgba(245,158,11,0.35)]'
        )}
      >
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="min-w-0">
            <div className="font-semibold text-[color:var(--text)]">
              Part {artifact.partNumber}
              {artifact.partCount ? ` of ${artifact.partCount}` : ''}
            </div>
            <div className="mt-1 text-sm text-[color:var(--muted)]">{artifact.fileName}</div>
            <div className="mt-2 flex flex-wrap items-center gap-2 text-xs text-[color:var(--soft)]">
              <span>
                {artifact.itemCount} item{artifact.itemCount === 1 ? '' : 's'}
              </span>
              <span>•</span>
              <span>{formatBytes(artifact.sizeBytes)}</span>
              {artifact.exceedsMaxSize ? (
                <>
                  <span>•</span>
                  <span>Item larger than selected limit</span>
                </>
              ) : null}
            </div>
          </div>
          <a
            className="button button-primary inline-flex items-center gap-2"
            href={artifact.downloadUrl}
            download={artifact.fileName}
          >
            <Download className="h-4 w-4" />
            Download
          </a>
        </div>
      </div>
    )
  }

  function renderJob(job: FlaggedBundleJob) {
    const isRunning = job.status === 'running'
    const progress = job.progress.percent
    const progressLabel =
      job.progress.currentLabel ||
      (isRunning ? 'Generating flagged bundle' : getJobStatusLabel(job.status))
    const summaryText = [
      `${job.progress.processedItems}/${job.progress.totalItems || 0} items`,
      `${job.progress.failedItems} failed`,
      formatBytes(job.maxSizeBytes)
    ].join(' • ')

    return (
      <section
        key={job.exportId}
        className="space-y-4 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface)] p-5 shadow-sm"
      >
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="min-w-0 space-y-2">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-lg font-semibold text-[color:var(--text)]">
                {job.scope.scopeLabel || 'Flagged bundle'}
              </div>
              <Badge
                className={cn(
                  'border-[color:var(--line)]',
                  isRunning && 'bg-[color:rgba(37,99,235,0.12)] text-[color:var(--accent)]',
                  job.status === 'succeeded' &&
                    'bg-[color:rgba(16,185,129,0.12)] text-[color:var(--success)]',
                  job.status === 'failed' &&
                    'bg-[color:rgba(220,38,38,0.08)] text-[color:var(--danger)]'
                )}
              >
                {getJobStatusLabel(job.status)}
              </Badge>
              <Badge className="border-[color:var(--line)] bg-[color:var(--surface)] text-[color:var(--muted)]">
                {formatDate(job.generatedAt)}
              </Badge>
            </div>
            <div className="text-sm text-[color:var(--muted)]">{job.scope.scopePath || currentScopeDescription}</div>
            <div className="flex flex-wrap items-center gap-2 text-xs text-[color:var(--soft)]">
              <span>{summaryText}</span>
              {job.error ? (
                <>
                  <span>•</span>
                  <span className="text-[color:var(--danger)]">{job.error}</span>
                </>
              ) : null}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {isRunning ? (
              <Badge className="border-[color:rgba(37,99,235,0.24)] bg-[color:rgba(37,99,235,0.12)] text-[color:var(--accent)]">
                <Loader2 className="mr-1 h-3.5 w-3.5 animate-spin" />
                Working
              </Badge>
            ) : null}
            {isTerminalJobStatus(job.status) ? (
              <Button
                variant="ghost"
                onClick={() => {
                  void handleDelete(job.exportId)
                }}
                disabled={deletingExportId === job.exportId}
              >
                {deletingExportId === job.exportId ? (
                  <RefreshCw className="h-4 w-4 animate-spin" />
                ) : (
                  <Trash2 className="h-4 w-4" />
                )}
                Delete
              </Button>
            ) : null}
          </div>
        </div>

        {isRunning ? (
          <div className="space-y-2 rounded-2xl border border-[color:var(--line)] bg-[color:rgba(15,23,42,0.03)] p-4">
            <div className="flex flex-wrap items-center justify-between gap-2 text-sm">
              <div className="font-medium text-[color:var(--text)]">{progressLabel}</div>
              <div className="text-[color:var(--muted)]">{progress}%</div>
            </div>
            <div className="h-2 overflow-hidden rounded-full bg-[color:rgba(148,163,184,0.25)]">
              <div
                className="h-full rounded-full bg-[color:var(--accent)] transition-[width]"
                style={{ width: `${Math.max(0, Math.min(100, progress))}%` }}
              />
            </div>
            <div className="flex flex-wrap items-center gap-2 text-xs text-[color:var(--soft)]">
              <span>{job.progress.processedItems} processed</span>
              <span>•</span>
              <span>{job.progress.failedItems} failed</span>
              <span>•</span>
              <span>{progressLabel}</span>
            </div>
          </div>
        ) : null}

        {job.status === 'failed' && job.progress.failedItems > 0 ? (
          <div className="flex items-start gap-2 rounded-2xl border border-[color:rgba(220,38,38,0.22)] bg-[color:rgba(220,38,38,0.08)] p-3 text-sm text-[color:var(--danger)]">
            <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
            <div>Some parts of this bundle could not be created. Any completed downloads remain available until you delete them.</div>
          </div>
        ) : null}

        <div className="space-y-4">
          {job.groups.map((group) => (
            <section key={group.groupType} className="space-y-3">
              <div className="flex flex-wrap items-center gap-2">
                <div className="text-base font-semibold text-[color:var(--text)]">{group.label}</div>
                <Badge className="border-[color:var(--line)] bg-[color:var(--surface)] text-[color:var(--muted)]">
                  {group.itemCount} item{group.itemCount === 1 ? '' : 's'}
                </Badge>
                {group.failedCount ? (
                  <Badge className="border-[color:rgba(220,38,38,0.22)] bg-[color:rgba(220,38,38,0.08)] text-[color:var(--danger)]">
                    {group.failedCount} failed
                  </Badge>
                ) : null}
              </div>

              {group.artifacts.length ? (
                <div className="space-y-3">{group.artifacts.map((artifact) => renderArtifactCard(job, artifact))}</div>
              ) : (
                <div className="rounded-2xl border border-dashed border-[color:var(--line)] bg-[color:var(--surface)] p-4 text-sm text-[color:var(--muted)]">
                  {job.status === 'running'
                    ? 'This export is still being generated.'
                    : 'No flagged items were exported for this group.'}
                </div>
              )}
            </section>
          ))}
        </div>
      </section>
    )
  }

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent
        showCloseButton={false}
        className="h-[min(96vh,900px)] w-[min(96vw,1180px)] overflow-hidden p-0"
      >
        <div className="flex h-full min-h-0 flex-col">
          <div className="flex items-start justify-between gap-4 border-b border-[color:var(--line)] px-6 py-5">
            <div className="min-w-0">
              <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                Review export
              </div>
              <DialogTitle className="mt-1 text-2xl">Download flagged bundle</DialogTitle>
              <DialogDescription>
                Generate separate ZIP downloads for mailbox items and Teams / SharePoint items. Completed downloads stay available until you delete them.
              </DialogDescription>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                <span className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                  Current scope
                </span>
                <Badge className="border-[color:var(--line)] bg-[color:var(--surface)] text-[color:var(--text)]">
                  {currentScopeLabel}
                </Badge>
                <span className="text-sm text-[color:var(--muted)]">{currentScopeDescription}</span>
              </div>
            </div>
            <Button variant="ghost" onClick={() => onOpenChange(false)}>
              <X className="h-4 w-4" />
              Close
            </Button>
          </div>

          <div className="grid min-h-0 flex-1 gap-0 lg:grid-cols-[24rem_minmax(0,1fr)]">
            <div className="border-b border-[color:var(--line)] p-6 lg:border-b-0 lg:border-r">
              <form className="space-y-5" onSubmit={handleGenerate}>
                <div className="space-y-2">
                  <div className="text-sm font-semibold text-[color:var(--text)]">ZIP segment size</div>
                  <div className="text-sm text-[color:var(--muted)]">
                    Choose the maximum size for each generated ZIP. Items are never split across files.
                  </div>
                </div>

                <label className="space-y-2 text-sm font-medium text-[color:var(--text)]">
                  <span>Preset</span>
                  <select
                    className="input"
                    value={preset}
                    onChange={(event) => {
                      const nextValue = event.target.value
                      if (!isSizePresetValue(nextValue)) {
                        return
                      }
                      setPreset(nextValue)
                      writeWorkspaceStorageItem('flaggedBundleSizePreset', authenticated, username, nextValue)
                    }}
                  >
                    {SIZE_PRESETS.map((item) => (
                      <option key={item.value} value={item.value}>
                        {item.label}
                      </option>
                    ))}
                  </select>
                </label>

                {preset === 'custom' ? (
                  <label className="space-y-2 text-sm font-medium text-[color:var(--text)]">
                    <span>Custom size in MB</span>
                    <Input
                      type="number"
                      min="1"
                      step="1"
                      value={customMb}
                      onChange={(event) => {
                        setCustomMb(event.target.value)
                        writeWorkspaceStorageItem(
                          'flaggedBundleCustomSizeMb',
                          authenticated,
                          username,
                          event.target.value
                        )
                      }}
                    />
                  </label>
                ) : null}

                <div className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface)] p-4 text-sm text-[color:var(--muted)]">
                  <div className="font-medium text-[color:var(--text)]">Selected limit</div>
                  <div>{formatBytes(sizeBytes)}</div>
                </div>

                <Button type="submit" disabled={busy || !scope || hasRunningJob || loadingJobs}>
                  {busy || loadingJobs ? (
                    <RefreshCw className="h-4 w-4 animate-spin" />
                  ) : (
                    <Download className="h-4 w-4" />
                  )}
                  {hasRunningJob ? 'Generation in progress' : busy ? 'Starting...' : 'Generate downloads'}
                </Button>

                {hasRunningJob ? (
                  <div className="rounded-2xl border border-[color:rgba(37,99,235,0.18)] bg-[color:rgba(37,99,235,0.08)] p-3 text-sm text-[color:var(--accent)]">
                    A bundle is already being generated for this workspace. You can close this dialog and come back later to check progress.
                  </div>
                ) : null}

                {error ? (
                  <div className="rounded-2xl border border-[color:rgba(220,38,38,0.22)] bg-[color:rgba(220,38,38,0.08)] p-3 text-sm text-[color:var(--danger)]">
                    {error}
                  </div>
                ) : null}
              </form>
            </div>

            <ScrollArea className="min-h-0 p-6">
              <div className="space-y-6">
                {runningJob ? (
                  <section className="space-y-3 rounded-3xl border border-[color:rgba(37,99,235,0.18)] bg-[color:rgba(37,99,235,0.06)] p-5">
                    <div className="flex flex-wrap items-center justify-between gap-3">
                      <div>
                        <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">
                          Current export
                        </div>
                        <div className="mt-1 text-lg font-semibold text-[color:var(--text)]">
                          {runningJob.scope.scopeLabel || 'Flagged bundle'}
                        </div>
                      <div className="mt-1 text-sm text-[color:var(--muted)]">
                          {runningJob.progress.currentLabel || 'Preparing bundle'}
                        </div>
                      </div>
                      <Badge className="border-[color:rgba(37,99,235,0.24)] bg-[color:rgba(37,99,235,0.12)] text-[color:var(--accent)]">
                        <Loader2 className="mr-1 h-3.5 w-3.5 animate-spin" />
                        {runningJob.progress.currentLabel || getJobStatusLabel(runningJob.status)}
                      </Badge>
                    </div>
                    <div className="h-2 overflow-hidden rounded-full bg-[color:rgba(148,163,184,0.25)]">
                      <div
                        className="h-full rounded-full bg-[color:var(--accent)] transition-[width]"
                        style={{ width: `${Math.max(0, Math.min(100, runningJob.progress.percent))}%` }}
                      />
                    </div>
                    <div className="flex flex-wrap items-center gap-2 text-xs text-[color:var(--soft)]">
                      <span>{runningJob.progress.processedItems} processed</span>
                      <span>•</span>
                      <span>{runningJob.progress.failedItems} failed</span>
                      <span>•</span>
                      <span>{formatBytes(runningJob.maxSizeBytes)}</span>
                    </div>
                  </section>
                ) : null}

                {jobs.length ? (
                  jobs.map((job) => renderJob(job))
                ) : (
                  <div className="empty-state h-full">
                    {loadingJobs ? (
                      <>
                        <div className="text-lg font-semibold text-[color:var(--text)]">
                          Loading export history...
                        </div>
                        <div className="mt-2 text-sm text-[color:var(--muted)]">
                          Checking whether this workspace already has a running or completed bundle.
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="text-lg font-semibold text-[color:var(--text)]">
                          No downloads generated yet
                        </div>
                        <div className="mt-2 text-sm text-[color:var(--muted)]">
                          Generate a bundle to begin tracking progress and downloading ZIP parts.
                        </div>
                      </>
                    )}
                  </div>
                )}

                <Separator />
                <div className="text-xs text-[color:var(--soft)]">
                  Export history stays available until you explicitly delete it.
                </div>
              </div>
            </ScrollArea>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}
