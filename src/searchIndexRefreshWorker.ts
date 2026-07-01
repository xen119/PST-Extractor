import { createReviewStoreFromEnv } from './reviewStore'
import { createSearchIndexStoreFromEnv, refreshSearchIndexFromCatalog } from './searchIndex'

function normalizeText(value: unknown): string {
  return String(value ?? '').trim()
}

async function main(): Promise<void> {
  try {
    const pstRootDir = normalizeText(process.env.PST_SEARCH_INDEX_PST_ROOT_DIR)
    const stagingDocumentsCollectionName = normalizeText(
      process.env.PST_SEARCH_INDEX_STAGING_DOCUMENTS_COLLECTION
    )

    if (!pstRootDir) {
      throw new Error('PST root directory is required')
    }
    if (!stagingDocumentsCollectionName) {
      throw new Error('Staging documents collection name is required')
    }

    const reviewStore = await createReviewStoreFromEnv(process.env)
    const searchIndexStore = await createSearchIndexStoreFromEnv(process.env, {
      documentsCollectionName: stagingDocumentsCollectionName
    })

    const summary = await refreshSearchIndexFromCatalog(pstRootDir, reviewStore, searchIndexStore)
    if (typeof process.send === 'function') {
      process.send({
        type: 'success',
        summary
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
