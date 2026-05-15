import * as fs from 'fs'
import * as http from 'http'
import * as os from 'os'
import * as path from 'path'
import { AddressInfo } from 'net'
import { buildOpenApiDocument } from '../openApi'
import { createPstReviewApp, type ApiSecurityConfig } from '../pstReviewApp'
import { MemoryReviewStore } from '../reviewStore'

const resolve = path.resolve

const enronPath = resolve('./src/__tests__/testdata/enron.pst')
const outlookPath = resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
const publicDir = resolve('./example/public')

jest.setTimeout(30000)

function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix))
}

function stageFixture(sourcePath: string, targetPath: string): void {
  try {
    fs.linkSync(sourcePath, targetPath)
  } catch {
    fs.copyFileSync(sourcePath, targetPath)
  }
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

async function startApp(pstRootDir: string, apiSecurity?: ApiSecurityConfig) {
  const reviewStore = new MemoryReviewStore()
  const app = createPstReviewApp({
    publicDir,
    pstRootDir,
    reviewStore,
    openApiSpec: buildOpenApiDocument({
      version: 'test',
      reviewStorageMode: reviewStore.kind
    }),
    apiSecurity
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
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(pstDir, 'Case2', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'sample.pst'))
    stageFixture(outlookPath, path.join(pstDir, 'Case2', 'Search1', 'archive.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore

    const catalog = await requestJson(`${started.baseUrl}/api/psts?scopePath=Case2/Search1`)
    expect(catalog.rootExists).toBe(true)
    expect(catalog.scopes.map((scope: { scopeLabel: string }) => scope.scopeLabel)).toEqual([
      'Case1 / Search1',
      'Case2 / Search1'
    ])
    expect(catalog.scopePath).toBe('Case2/Search1')
    expect(catalog.scopeLabel).toBe('Case2 / Search1')
    expect(catalog.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'archive.ost'
    ])

    await expect(
      requestJson(`${started.baseUrl}/api/psts/open`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          scopePath: '../Case2/Search1',
          fileName: 'archive.ost'
        })
      })
    ).rejects.toThrow('Scope path must stay within the PST folder')

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case2/Search1',
        fileName: 'archive.ost'
      })
    })
    expect(opened.scopePath).toBe('Case2/Search1')
    expect(opened.scopeLabel).toBe('Case2 / Search1')

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

    const recipientMatch = String(
      detail.detail.resolvedDisplayTo || detail.detail.displayTo || ''
    ).match(/<([^>]+)>/)
    expect(recipientMatch).toBeTruthy()
    const recipientEmail = recipientMatch ? recipientMatch[1] : ''

    const searchAll = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=signature&mailOnly=1`
    )
    expect(searchAll.scope).toBe('all')
    expect(searchAll.page.items.some((item: { id: string }) => item.id === message.id)).toBe(true)

    const searchSelected = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=Case2/Search1&query=signature&mailOnly=1`
    )
    expect(searchSelected.scope).toBe('search')
    expect(searchSelected.page.scopePath).toBe('Case2/Search1')
    expect(searchSelected.page.items.some((item: { id: string }) => item.id === message.id)).toBe(
      true
    )

    const searchMailbox = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=signature&mailOnly=1`
    )
    expect(searchMailbox.scope).toBe('pst')
    expect(searchMailbox.page.items.some((item: { id: string }) => item.id === message.id)).toBe(
      true
    )

    if (recipientEmail) {
      const recipientSearch = await requestJson(
        `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=${encodeURIComponent(
          recipientEmail
        )}&mailOnly=1`
      )
      expect(
        recipientSearch.page.items.some((item: { id: string }) => item.id === message.id)
      ).toBe(true)
    }

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

  it('opens PST root mailboxes when scopePath is empty', async () => {
    rootDir = makeTempDir('pst-review-api-root-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(outlookPath, path.join(pstDir, 'root.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore

    const catalog = await requestJson(`${started.baseUrl}/api/psts`)
    expect(catalog.scopePath).toBe('')
    expect(catalog.scopeLabel).toBe('PST root')
    expect(catalog.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'root.ost'
    ])

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: '',
        fileName: 'root.ost'
      })
    })

    expect(opened.scopePath).toBe('')
    expect(opened.scopeLabel).toBe('PST root')
    expect(opened.fileName).toBe('root.ost')
  })

  it('bypasses auth for loopback requests', async () => {
    rootDir = makeTempDir('pst-review-api-bypass-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const authCheck = jest.fn(async (_req, _res, next) => next())
    const started = await startApp(pstDir, {
      bypassIps: ['127.0.0.1', '::1', '::ffff:127.0.0.1', 'localhost'],
      m365Auth: {
        CheckTokens: authCheck
      }
    })
    server = started.server
    reviewStore = started.reviewStore

    const catalog = await requestJson(`${started.baseUrl}/api/psts`)
    expect(catalog.files).toHaveLength(1)
    expect(authCheck).not.toHaveBeenCalled()
  })

  it('rejects remote requests without a token when auth is required', async () => {
    rootDir = makeTempDir('pst-review-api-auth-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, {
      bypassIps: [],
      m365Auth: {
        CheckTokens: async (req, res, next) => {
          if (!req.headers['x-graph-token']) {
            return res.status(401).json({
              success: false,
              message: 'No proof of possession.'
            })
          }
          return next()
        }
      }
    })
    server = started.server
    reviewStore = started.reviewStore

    const response = await fetch(`${started.baseUrl}/api/psts`, {
      headers: {
        'X-Forwarded-For': '203.0.113.10'
      }
    })
    const payload = await readJson(response)
    expect(response.status).toBe(401)
    expect(payload.message).toBe('No proof of possession.')
  })

  it('allows remote requests when auth succeeds and responds to cors preflight', async () => {
    rootDir = makeTempDir('pst-review-api-cors-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const authCheck = jest.fn(async (_req, _res, next) => next())
    const started = await startApp(pstDir, {
      bypassIps: [],
      allowedOrigins: ['https://app.example.test'],
      m365Auth: {
        CheckTokens: authCheck
      }
    })
    server = started.server
    reviewStore = started.reviewStore

    const preflight = await fetch(`${started.baseUrl}/api/psts`, {
      method: 'OPTIONS',
      headers: {
        Origin: 'https://app.example.test',
        'Access-Control-Request-Method': 'GET',
        'Access-Control-Request-Headers': 'content-type, x-graph-token'
      }
    })

    expect(preflight.status).toBe(204)
    expect(preflight.headers.get('access-control-allow-origin')).toBe('https://app.example.test')
    expect(preflight.headers.get('access-control-allow-methods')).toContain('GET')
    expect(preflight.headers.get('access-control-allow-headers')).toContain('x-graph-token')
    expect(authCheck).not.toHaveBeenCalled()

    const response = await fetch(`${started.baseUrl}/api/psts`, {
      headers: {
        Origin: 'https://app.example.test',
        'X-Graph-Token': 'demo-token',
        'X-Forwarded-For': '203.0.113.10'
      }
    })

    expect(response.ok).toBe(true)
    expect(response.headers.get('access-control-allow-origin')).toBe('https://app.example.test')
    expect(authCheck).toHaveBeenCalled()
  })
})
