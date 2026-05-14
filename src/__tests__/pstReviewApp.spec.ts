import * as fs from 'fs'
import * as http from 'http'
import * as os from 'os'
import * as path from 'path'
import { AddressInfo } from 'net'
import { buildOpenApiDocument } from '../openApi'
import { createPstReviewApp } from '../pstReviewApp'
import { MemoryReviewStore } from '../reviewStore'

const resolve = path.resolve

const enronPath = resolve('./src/__tests__/testdata/enron.pst')
const publicDir = resolve('./example/public')

function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix))
}

function findMailFolder(node: {
  id: string
  mailMessageCount?: number
  children?: Array<{
    id: string
    mailMessageCount?: number
    children?: unknown[]
  }>
}): { id: string; mailMessageCount?: number } | null {
  if ((node.mailMessageCount || 0) > 0) {
    return node
  }

  for (const child of node.children || []) {
    const match = findMailFolder(child)
    if (match) {
      return match
    }
  }

  return null
}

async function readJson(response: Response): Promise<any> {
  const text = await response.text()
  return text ? JSON.parse(text) : null
}

async function requestJson(url: string, init?: RequestInit): Promise<any> {
  const response = await fetch(url, init)
  const payload = await readJson(response)
  if (!response.ok) {
    throw new Error(payload?.error || response.statusText)
  }
  return payload
}

async function startApp(pstRootDir: string) {
  const reviewStore = new MemoryReviewStore()
  const app = createPstReviewApp({
    publicDir,
    pstRootDir,
    reviewStore,
    openApiSpec: buildOpenApiDocument({
      version: 'test',
      reviewStorageMode: reviewStore.kind
    })
  })

  const server = app.listen(0)
  await new Promise<void>((resolveListening) => {
    server.once('listening', resolveListening)
  })

  const address = server.address() as AddressInfo
  return {
    baseUrl: `http://127.0.0.1:${address.port}`,
    reviewStore,
    server
  }
}

describe('pst review api', () => {
  let rootDir: string | null = null
  let server: http.Server | null = null
  let reviewStore: MemoryReviewStore | null = null

  afterEach(async () => {
    if (server) {
      await new Promise<void>((resolveClose) => {
        server!.close(() => resolveClose())
      })
      server = null
    }
    if (reviewStore) {
      await reviewStore.close()
      reviewStore = null
    }
    if (rootDir) {
      fs.rmSync(rootDir, { recursive: true, force: true })
      rootDir = null
    }
  })

  it('serves mailbox catalog, extraction, review endpoints, and swagger docs', async () => {
    rootDir = makeTempDir('pst-review-api-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    fs.copyFileSync(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore

    const catalog = await requestJson(`${started.baseUrl}/api/psts`)
    expect(catalog.rootExists).toBe(true)
    expect(catalog.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'sample.pst'
    ])

    await expect(
      requestJson(`${started.baseUrl}/api/psts/open`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({})
      })
    ).rejects.toThrow('Mailbox file name is required')

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ fileName: 'sample.pst' })
    })

    const folder = findMailFolder(opened.tree)
    expect(folder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?page=1&pageSize=20`
    )
    expect(folderPage.page.items.length).toBeGreaterThan(0)

    const message = folderPage.page.items.find((item: { isMailLike: boolean }) => item.isMailLike)
    expect(message).toBeTruthy()

    const detail = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}`
    )
    expect(detail.detail.review).toEqual({
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    })

    const patch = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['Urgent', 'urgent', 'Follow up']
        })
      }
    )
    expect(patch.review.flagged).toBe(true)
    expect(patch.review.tags).toEqual(['Urgent', 'Follow up'])

    const review = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/review`
    )
    expect(review.review.flagged).toBe(true)
    expect(review.review.tags).toEqual(['Urgent', 'Follow up'])

    const queue = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/review?reviewFlagged=1`
    )
    expect(queue.items.map((item: { messageId: string }) => item.messageId)).toContain(message.id)

    const extraction = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/extract?fields=summary,review,attachments`
    )
    expect(extraction.record.summary.subject).toBe(detail.detail.subject)
    expect(extraction.record.review.flagged).toBe(true)

    const folderExtraction = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages/extract?fields=summary,review&pageSize=5`
    )
    expect(folderExtraction.items.some((item: { messageId: string; record: { review: { flagged: boolean } } }) => item.messageId === message.id && item.record.review.flagged)).toBe(true)

    const deleted = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/review`,
      {
        method: 'DELETE'
      }
    )
    expect(deleted.review.flagged).toBe(false)
    expect(deleted.review.tags).toEqual([])

    const openApi = await requestJson(`${started.baseUrl}/api/openapi.json`)
    expect(openApi.paths['/api/psts/open']).toBeDefined()
    expect(openApi.paths['/api/sessions/{sessionId}/messages/{messageId}/review']).toBeDefined()

    const docsResponse = await fetch(`${started.baseUrl}/api/docs`)
    const docsHtml = await docsResponse.text()
    expect(docsHtml).toContain('SwaggerUIBundle')
    expect(docsHtml).toContain('/api/openapi.json')
  })
})
