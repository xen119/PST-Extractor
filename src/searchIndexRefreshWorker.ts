import { createReviewStoreFromEnv } from './reviewStore'
import {
  createSearchIndexStoreFromEnv,
  refreshSearchIndexSourceFromCatalog,
  type SearchIndexRefreshSource
} from './searchIndex'

function normalizeRefreshSource(value: unknown): SearchIndexRefreshSource {
  return String(value ?? '').trim().toLowerCase() === 'items' ? 'items' : 'mailboxes'
}

function normalizeText(value: unknown): string {
  return String(value ?? '').trim()
}

async function main(): Promise<void> {
  try {
    const pstRootDir = normalizeText(process.env.PST_SEARCH_INDEX_PST_ROOT_DIR)
    const stagingDocumentsCollectionName = normalizeText(
      process.env.PST_SEARCH_INDEX_STAGING_DOCUMENTS_COLLECTION
    )
    const stagingMailboxDetailsCollectionName = normalizeText(
      process.env.PST_SEARCH_INDEX_STAGING_MAILBOX_DETAILS_COLLECTION
    )
    const source = normalizeRefreshSource(process.env.PST_SEARCH_INDEX_REFRESH_SOURCE)

    if (!pstRootDir) {
      throw new Error('PST root directory is required')
    }
    if (!stagingDocumentsCollectionName) {
      throw new Error('Staging documents collection name is required')
    }
    if (!stagingMailboxDetailsCollectionName) {
      throw new Error('Staging mailbox detail snapshot collection name is required')
    }

    const reviewStore = await createReviewStoreFromEnv(process.env)
    const searchIndexStore = await createSearchIndexStoreFromEnv(process.env, {
      documentsCollectionName: stagingDocumentsCollectionName,
      mailboxDetailsCollectionName: stagingMailboxDetailsCollectionName
    })

    const plan = await refreshSearchIndexSourceFromCatalog(
      pstRootDir,
      source,
      reviewStore,
      searchIndexStore,
      {
        pruneRemovedFiles: false,
        updateFingerprints: false
      }
    )

    if (typeof process.send === 'function') {
      process.send({
        type: 'success',
        plan
      })
    }
    await searchIndexStore.close()
    await reviewStore.close()
    process.exitCode = 0
  } catch (error) {
    if (typeof process.send === 'function') {
      process.send({
        type: 'failure',
        error: error instanceof Error ? error.message : String(error)
      })
    }
    process.exitCode = 1
  }
}

if (require.main === module) {
  void main()
}
