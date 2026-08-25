import { createSearchIndexStoreFromEnv } from './searchIndex'
import { runMailboxCollapseJob } from './searchIndexCollapse'

async function main(): Promise<void> {
  const store = await createSearchIndexStoreFromEnv(process.env)
  try {
    const status = await runMailboxCollapseJob(store, {
      jobId: process.env.PST_SEARCH_INDEX_COLLAPSE_JOB_ID || undefined,
      resetBefore: process.env.PST_SEARCH_INDEX_COLLAPSE_RESET === '1'
    })
    if (typeof process.send === 'function') {
      process.send({ type: status.status === 'failed' || status.status === 'reindex-required' ? 'failure' : 'success' })
    }
  } catch (error) {
    await store.saveMailboxCollapseJob?.({
      ...(await store.getMailboxCollapseJob?.()),
      ...{
        jobId: process.env.PST_SEARCH_INDEX_COLLAPSE_JOB_ID || null,
        status: 'failed',
        version: 6,
        startedAt: null,
        completedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        processedPartitions: 0,
        totalPartitions: 0,
        completedPartitionKeys: [],
        processedWorkUnits: 0,
        totalWorkUnits: 0,
        percentage: 0,
        provisional: true,
        error: error instanceof Error ? error.message : String(error),
        reindexRequired: false
      }
    })
    if (typeof process.send === 'function') {
      process.send({ type: 'failure', error: error instanceof Error ? error.message : String(error) })
    }
    process.exitCode = 1
  } finally {
    await store.close()
  }
}

if (require.main === module) {
  void main()
}
