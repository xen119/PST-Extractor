import * as fs from 'fs'
import * as http from 'http'
import * as os from 'os'
import * as path from 'path'
import AdmZip from 'adm-zip'
import { AddressInfo } from 'net'
import { generateTotpCode } from '../authSecurity'
import { createMemoryAuthUserStore, type AuthUserStore } from '../authUsers'
import { createMemoryAppSettingsStore, type AppSettingsStore } from '../appSettings'
import { buildOpenApiDocument } from '../openApi'
import { createPstReviewApp, type ApiSecurityConfig, type AppAuthConfig } from '../pstReviewApp'
import { MemoryReviewStore } from '../reviewStore'
import { MemorySearchIndexStore, type SearchIndexDocument, type SearchIndexFileFingerprint } from '../searchIndex'

const resolve = path.resolve

const enronPath = resolve('./src/__tests__/testdata/enron.pst')
const outlookPath = resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
const publicDir = resolve('./example/public')

jest.setTimeout(30000)

interface StartAppOptions {
  searchIndexStore?: MemorySearchIndexStore
  skipInitialRefresh?: boolean
  backgroundInitialRefresh?: boolean
}

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

function stageArchiveBundle(targetPath: string, entries: Array<[string, string]>): void {
  const zip = new AdmZip()
  for (const [entryName, content] of entries) {
    zip.addFile(entryName, Buffer.from(content, 'utf8'))
  }
  zip.writeZip(targetPath)
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

function findFolderByDisplayName(
  node: {
    id: string
    displayName?: string
    children?: Array<{
      id: string
      displayName?: string
      children?: unknown[]
    }>
  },
  displayName: string
): { id: string; displayName?: string } | null {
  if (node.displayName === displayName) {
    return node
  }

  for (const child of node.children || []) {
    const match = findFolderByDisplayName(child, displayName)
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

async function requestBuffer(url: string, init?: RequestInit): Promise<{ response: Response; buffer: Buffer }> {
  const response = await fetch(url, init)
  const buffer = Buffer.from(await response.arrayBuffer())
  if (!response.ok) {
    const text = buffer.toString('utf8')
    let payload: any = null
    try {
      payload = text ? JSON.parse(text) : null
    } catch {
      payload = null
    }
    throw new Error((payload && payload.error) || response.statusText)
  }
  return { response, buffer }
}

function getSetCookieHeader(response: Response): string {
  return response.headers.get('set-cookie') || ''
}

function getCookiePair(setCookieHeader: string): string {
  return setCookieHeader.split(';')[0] || ''
}

function readAuditLogEntries(filePath: string): any[] {
  if (!fs.existsSync(filePath)) {
    return []
  }

  return fs
    .readFileSync(filePath, 'utf8')
    .split(/\r?\n/g)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => JSON.parse(line))
}

function createDeferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void
  let reject!: (reason?: unknown) => void
  const promise = new Promise<T>((promiseResolve, promiseReject) => {
    resolve = promiseResolve
    reject = promiseReject
  })

  return { promise, resolve, reject }
}

function parseStoredZipEntries(buffer: Buffer): Map<string, Buffer> {
  const entries = new Map<string, Buffer>()
  let offset = 0
  while (offset + 4 <= buffer.length) {
    const signature = buffer.readUInt32LE(offset)
    if (signature === 0x04034b50) {
      const compression = buffer.readUInt16LE(offset + 8)
      expect(compression).toBe(0)
      const nameLength = buffer.readUInt16LE(offset + 26)
      const extraLength = buffer.readUInt16LE(offset + 28)
      const compressedSize = buffer.readUInt32LE(offset + 18)
      const fileName = buffer
        .slice(offset + 30, offset + 30 + nameLength)
        .toString('utf8')
      const dataStart = offset + 30 + nameLength + extraLength
      const dataEnd = dataStart + compressedSize
      entries.set(fileName, buffer.slice(dataStart, dataEnd))
      offset = dataEnd
      continue
    }
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break
    }
    offset += 1
  }
  return entries
}

function makeSearchIndexDocument(overrides: Partial<SearchIndexDocument> = {}): SearchIndexDocument {
  const now = new Date('2024-01-01T00:00:00.000Z').toISOString()
  return {
    mailboxKey: 'C:/PST/Case1/Search1/alpha.pst',
    scopePath: 'Case1/Search1',
    scopeLabel: 'Case1 / Search1',
    fileName: 'alpha.pst',
    mailboxName: 'Alpha',
    messageId: 'message:1',
    descriptorId: '1',
    folderId: 'folder:1',
    folderPath: 'Inbox',
    order: 1,
    messageClass: 'IPM.Note',
    kind: 'mail',
    subject: 'Project Alpha',
    originalSubject: 'Re: Project Alpha',
    senderName: 'Alice Example',
    senderEmailAddress: 'alice@example.com',
    recipientText: 'Bob Example <bob@example.com>',
    displayTo: 'Bob Example <bob@example.com>',
    displayCC: '',
    displayBCC: '',
    resolvedDisplayTo: 'Bob Example <bob@example.com>',
    resolvedDisplayCC: '',
    resolvedDisplayBCC: '',
    clientSubmitTime: now,
    creationTime: now,
    modificationTime: now,
    messageDeliveryTime: now,
    sortDate: now,
    sortDateMs: Date.parse(now),
    importance: 1,
    hasAttachments: false,
    isRead: true,
    isMailLike: true,
    bodySearchText: 'project alpha signature',
    searchText: 'project alpha alice example alice@example.com bob example bob@example.com signature ipm.note mail',
    searchTokens: ['project', 'alpha', 'signature'],
    addressValues: ['alice@example.com', 'bob@example.com'],
    subjectValues: ['project alpha', 're: project alpha'],
    review: {
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    },
    reviewStates: [],
    reviewTagValues: [],
    updatedAt: now,
    sourceType: 'mailbox',
    ...overrides
  }
}

function makeSearchIndexFingerprint(
  overrides: Partial<SearchIndexFileFingerprint> & { source: SearchIndexFileFingerprint['source'] }
): SearchIndexFileFingerprint {
  const now = new Date('2024-01-01T00:00:00.000Z').toISOString()
  return {
    source: overrides.source,
    mailboxKey: 'C:/PST/Case1/Search1/alpha.pst',
    fileName: 'alpha.pst',
    scopePath: 'Case1/Search1',
    scopeLabel: 'Case1 / Search1',
    size: 1024,
    modifiedAt: now,
    updatedAt: now,
    ...overrides
  }
}

async function startApp(
  pstRootDir: string,
  apiSecurity?: ApiSecurityConfig,
  auth?: AppAuthConfig,
  authUserStore?: AuthUserStore,
  appSettingsStore?: AppSettingsStore,
  smtpTransportFactory?: (settings: any) => any,
  options?: StartAppOptions
) {
  const auditLogDir = path.join(path.dirname(pstRootDir), 'logs')
  const reviewStore = new MemoryReviewStore()
  const searchIndexStore = options?.searchIndexStore || new MemorySearchIndexStore()
  const app = createPstReviewApp({
    publicDir,
    pstRootDir,
    reviewStore,
    searchIndexStore,
    openApiSpec: buildOpenApiDocument({
      version: 'test',
      reviewStorageMode: reviewStore.kind
    }),
    auth,
    authUserStore,
    appSettingsStore,
    auditLogDir,
    smtpTransportFactory,
    apiSecurity
  })

  const server = app.listen(0)
  await new Promise<void>((resolveListening) => {
    server.once('listening', resolveListening)
  })

  const searchIndexRefreshCoordinator = app.get('searchIndexRefreshCoordinator') as
    | {
        start(
          source: 'mailboxes' | 'items',
          trigger: 'startup' | 'manual'
        ): Promise<{ status: string }>
        getStatus(source: 'mailboxes' | 'items'): { status: string }
      }
    | undefined
  if (!options?.skipInitialRefresh && searchIndexRefreshCoordinator) {
    const startPromises = [
      searchIndexRefreshCoordinator.start('mailboxes', 'startup'),
      searchIndexRefreshCoordinator.start('items', 'startup')
    ]
    if (!options?.backgroundInitialRefresh) {
      const started = await Promise.all(startPromises)
      if (started.some((entry) => entry.status === 'running')) {
        await waitForSearchIndexRefreshCompletion(searchIndexRefreshCoordinator, 'mailboxes')
        await waitForSearchIndexRefreshCompletion(searchIndexRefreshCoordinator, 'items')
      }
    }
  }

  const address = server.address() as AddressInfo
  return {
    baseUrl: `http://127.0.0.1:${address.port}`,
    reviewStore,
    searchIndexStore,
    appSettingsStore,
    auditLogDir,
    auditLogPath: path.join(auditLogDir, 'activity.log'),
    server
  }
}

async function waitForSearchIndexRefreshCompletion(
  coordinator: { getStatus(source: 'mailboxes' | 'items'): { status: string } },
  source: 'mailboxes' | 'items'
): Promise<void> {
  for (;;) {
    const status = coordinator.getStatus(source)
    if (status.status !== 'running') {
      return
    }
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
}

async function waitForRefreshStatus(baseUrl: string, cookie: string): Promise<void> {
  for (let attempt = 0; attempt < 200; attempt += 1) {
    const response = await fetch(`${baseUrl}/api/search/index/refresh/status?source=mailboxes`, {
      headers: {
        Cookie: cookie
      }
    })
    const payload = await readJson(response)
    if (response.ok && payload?.status?.status !== 'running') {
      return
    }
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error('Search index refresh did not finish in time')
}

describe('pst review api', () => {
  let rootDir: string | null = null
  let server: http.Server | null = null
  let reviewStore: MemoryReviewStore | null = null
  let searchIndexStore: MemorySearchIndexStore | null = null

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
    if (searchIndexStore) {
      await searchIndexStore.close()
      searchIndexStore = null
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
    searchIndexStore = started.searchIndexStore

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

    const { PSTMessage } = require('../PSTMessage.class')
    const originalGetAttachment = PSTMessage.prototype.getAttachment
    const attachmentSpy = jest
      .spyOn(PSTMessage.prototype, 'getAttachment')
      .mockImplementation(function (this: PSTMessage, index: number) {
        if (index === 0) {
          return {
            embeddedPSTMessage: null,
            fileInputStream: null,
            filename: 'missing.txt',
            longFilename: '',
            longPathname: '',
            pathname: '',
            mimeTag: 'text/plain',
            contentId: '',
            attachMethod: 1,
            filesize: 0
          } as any
        }
        return originalGetAttachment.call(this, index)
      })

    const attachmentResponse = await fetch(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        'message:2110308'
      )}/attachments/0`
    )
    const attachmentPayload = await readJson(attachmentResponse)
    expect(attachmentResponse.status).toBe(404)
    expect(attachmentPayload.error).toContain('Attachment bytes are not stored in this PST')
    attachmentSpy.mockRestore()

    const attachmentDetail = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        'message:2110308'
      )}`
    )
    expect(attachmentDetail.detail.attachments[0].isDownloadable).toBe(true)
    expect(attachmentDetail.detail.attachments[0].downloadUrl).toContain('/attachments/0')

    const recipientMatch = String(
      detail.detail.resolvedDisplayTo || detail.detail.displayTo || ''
    ).match(/<([^>]+)>/)
    const searchSources = [
      recipientMatch ? recipientMatch[1] : '',
      detail.detail.senderEmailAddress || '',
      detail.detail.subject || message.subject || '',
      detail.detail.originalSubject || '',
      detail.detail.bodyPrefix || '',
      detail.detail.bodyText || ''
    ]
    const searchCandidates = Array.from(
      new Set(
        searchSources
          .flatMap((source) =>
            String(source)
              .split(/[^A-Za-z0-9@._%+-]+/)
              .map((part) => part.trim())
              .filter((part) => part.length >= 4)
          )
          .filter(Boolean)
      )
    )
    let searchAll: any = null
    let searchTerm = ''
    for (const candidate of searchCandidates) {
      const attempt = await requestJson(
        `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
          candidate
        )}&mailOnly=1&pageSize=5000`
      )
      if (attempt.page.total > 0) {
        searchAll = attempt
        searchTerm = candidate
        break
      }
    }
    expect(searchAll).toBeTruthy()
    expect(searchTerm).toBeTruthy()
    expect(searchAll.scope).toBe('all')
    expect(searchAll.page.mode).toBe('and')
    expect(searchAll.page.total).toBeGreaterThan(0)
    const searchResult =
      searchAll.page.items.find((item: { messageId: string }) => item.messageId === message.id) ||
      searchAll.page.items.find((item: { messageId: string }) => Boolean(item.messageId))
    expect(searchResult).toBeTruthy()
    const resultRecipientMatch = String(
      searchResult.resolvedDisplayTo || searchResult.displayTo || ''
    ).match(/<([^>]+)>/)
    const hiddenRuleKind = resultRecipientMatch ? 'address' : 'subject'

    const searchSelected = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=Case2/Search1&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=5000`
    )
    expect(searchSelected.scope).toBe('search')
    expect(searchSelected.page.scopePath).toBe('Case2/Search1')
    expect(searchSelected.page.mode).toBe('and')
    expect(searchSelected.page.total).toBeGreaterThan(0)

    const searchMailbox = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=5000`
    )
    expect(searchMailbox.scope).toBe('pst')
    expect(searchMailbox.page.total).toBeGreaterThan(0)

    const refreshedIndexResponse = await fetch(`${started.baseUrl}/api/search/index/refresh?source=mailboxes`, {
      method: 'POST'
    })
    const refreshedIndex = await readJson(refreshedIndexResponse)
    expect(refreshedIndexResponse.status).toBe(202)
    expect(refreshedIndex.status.status).toBe('running')

    const searchAnd = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=${encodeURIComponent(
        `+ ${searchTerm}`
      )}&mailOnly=1&pageSize=5000`
    )
    expect(searchAnd.page.mode).toBe('and')
    expect(searchAnd.page.total).toBeGreaterThan(0)

    const searchOr = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=${encodeURIComponent(
        `${searchTerm} | missingterm`
      )}&mailOnly=1&pageSize=5000`
    )
    expect(searchOr.page.mode).toBe('or')
    expect(searchOr.page.total).toBeGreaterThan(0)

    if (searchSources[0]) {
      const recipientSearch = await requestJson(
        `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&query=${encodeURIComponent(
          searchSources[0]
        )}&mailOnly=1&pageSize=5000`
      )
      expect(recipientSearch.page.total).toBeGreaterThan(0)
    }

    const hiddenSubject =
      hiddenRuleKind === 'address'
        ? resultRecipientMatch
          ? resultRecipientMatch[1]
          : searchSources[0]
        : searchResult?.subject || detail.detail.subject || searchTerm
    expect(folderPage.page.total).toBeGreaterThan(0)
    expect(
      folderPage.page.items.some((item: { id: string }) => item.id === message.id)
    ).toBe(true)

    const hiddenSubjectResponse = await requestJson(`${started.baseUrl}/api/search/filters`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        kind: hiddenRuleKind,
        value: hiddenSubject,
        label: hiddenSubject
      })
    })
    expect(hiddenSubjectResponse.rule.kind).toBe(hiddenRuleKind)
    expect(hiddenSubjectResponse.rule.value).toBe(hiddenSubject.toLowerCase())

    const folderAfterHidden = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?mailOnly=1&pageSize=5000`
    )
    expect(folderAfterHidden.page.total).toBeLessThan(folderPage.page.total)
    expect(
      folderAfterHidden.page.items.some((item: { id: string }) => item.id === message.id)
    ).toBe(false)

    const hiddenFilters = await requestJson(`${started.baseUrl}/api/search/filters`)
    expect(hiddenFilters.items.some((item: { value: string }) => item.value === hiddenSubject.toLowerCase())).toBe(true)

    const searchHidden = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=5000`
    )
    if (searchResult) {
      expect(
        searchHidden.page.items.some(
          (item: { messageId: string }) => item.messageId === searchResult.messageId
        )
      ).toBe(
        false
      )
    }

    await requestJson(
      `${started.baseUrl}/api/search/filters/${encodeURIComponent(hiddenSubjectResponse.rule.filterId)}`,
      {
        method: 'DELETE'
      }
    )

    const searchRestored = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=5000`
    )
    expect(searchRestored.page.total).toBeGreaterThanOrEqual(searchHidden.page.total)
    if (searchResult) {
      expect(
        searchRestored.page.items.some(
          (item: { messageId: string }) => item.messageId === searchResult.messageId
        )
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
    expect(openApi.paths['/api/items/{itemId}/review']).toBeDefined()
    expect(openApi.paths['/api/exports/flagged.zip']).toBeDefined()
    expect(openApi.paths['/api/auth/password-reset/request']).toBeDefined()
    expect(openApi.paths['/api/auth/password-reset/{token}']).toBeDefined()
    expect(openApi.paths['/api/auth/password-reset/{token}/confirm']).toBeDefined()
    expect(openApi.paths['/api/auth/users/{username}/mfa/enforce']).toBeDefined()
    expect(openApi.paths['/api/auth/users/{username}/access']).toBeDefined()
    const authStatusSchema = (openApi.components as { schemas?: Record<string, { properties?: Record<string, unknown> }> } | undefined)?.schemas?.AuthStatus
    expect(authStatusSchema?.properties?.passwordResetAvailable).toBeDefined()
    expect(authStatusSchema?.properties?.loginFailedCount).toBeDefined()
    expect(authStatusSchema?.properties?.lockedUntil).toBeDefined()

    const docsResponse = await fetch(`${started.baseUrl}/api/docs`)
    const docsHtml = await docsResponse.text()
    expect(docsHtml).toContain('SwaggerUIBundle')
    expect(docsHtml).toContain('/api/openapi.json')
  })

  it('serves mailbox detail from snapshots and refreshes them after review changes', async () => {
    rootDir = makeTempDir('pst-review-detail-cache-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    const mailboxPath = path.join(pstDir, 'Case1', 'Search1', 'sample.pst')
    stageFixture(enronPath, mailboxPath)

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })

    const folder = findMailFolder(opened.tree)
    expect(folder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?page=1&pageSize=20`
    )
    const message = folderPage.page.items.find((item: { isMailLike: boolean }) => item.isMailLike)
    expect(message).toBeTruthy()
    if (!message) {
      throw new Error('Expected a mailbox message')
    }

    const snapshotBeforeDetail = await started.searchIndexStore.findMailboxDetail(mailboxPath, message.id)
    expect(snapshotBeforeDetail).toBeTruthy()

    const { PSTUtil } = require('../PSTUtil.class')
    const loadSpy = jest.spyOn(PSTUtil, 'detectAndLoadPSTObject')
    const reviewSpy = jest.spyOn(started.reviewStore, 'getReview')

    const detailOne = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(message.id)}`
    )
    const extractOne = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/extract?fields=all`
    )
    const detailTwo = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(message.id)}`
    )

    expect(detailTwo.detail).toEqual(detailOne.detail)
    expect(extractOne.record.review.flagged).toBe(detailOne.detail.review.flagged)
    expect(loadSpy).not.toHaveBeenCalled()
    expect(reviewSpy).toHaveBeenCalledTimes(1)

    await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true
        })
      }
    )

    const detailAfterUpdate = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(message.id)}`
    )
    const extractAfterUpdate = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/extract?fields=all`
    )

    expect(detailAfterUpdate.detail.review.flagged).toBe(true)
    expect(extractAfterUpdate.record.review.flagged).toBe(true)
    expect(loadSpy).not.toHaveBeenCalled()
    expect(reviewSpy).toHaveBeenCalledTimes(2)
  })

  it('backfills a missing mailbox snapshot from PST once and reuses it thereafter', async () => {
    rootDir = makeTempDir('pst-review-detail-backfill-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    const mailboxPath = path.join(pstDir, 'Case1', 'Search1', 'sample.pst')
    stageFixture(enronPath, mailboxPath)

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })

    const folder = findMailFolder(opened.tree)
    expect(folder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?page=1&pageSize=20`
    )
    const message = folderPage.page.items.find((item: { isMailLike: boolean }) => item.isMailLike)
    expect(message).toBeTruthy()
    if (!message) {
      throw new Error('Expected a mailbox message')
    }

    await started.searchIndexStore.deleteMailboxDetails(mailboxPath)

    const { PSTUtil } = require('../PSTUtil.class')
    const loadSpy = jest.spyOn(PSTUtil, 'detectAndLoadPSTObject')
    const upsertSpy = jest.spyOn(started.searchIndexStore, 'upsertMailboxDetail')

    const detailOne = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(message.id)}`
    )
    const loadCountAfterFirstDetail = loadSpy.mock.calls.length
    const upsertCountAfterFirstDetail = upsertSpy.mock.calls.length
    const detailTwo = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(message.id)}`
    )

    expect(detailTwo.detail).toEqual(detailOne.detail)
    expect(loadCountAfterFirstDetail).toBeGreaterThan(0)
    expect(loadSpy.mock.calls.length).toBe(loadCountAfterFirstDetail)
    expect(upsertSpy).toHaveBeenCalledTimes(1)
    expect(upsertCountAfterFirstDetail).toBe(1)
    expect(await started.searchIndexStore.findMailboxDetail(mailboxPath, message.id)).toBeTruthy()
  })

  it('moves PSTs into and out of the removed archive without deleting them', async () => {
    rootDir = makeTempDir('pst-review-archive-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'sample.pst'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore
    const searchIndexRefreshSpy = jest.spyOn(started.searchIndexStore, 'replaceMailboxDocuments')

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })

    const folder = findMailFolder(opened.tree)
    expect(folder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?page=1&pageSize=20`
    )
    const message = folderPage.page.items.find((item: { isMailLike: boolean }) => item.isMailLike)
    expect(message).toBeTruthy()

    const detail = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}`
    )
    const searchTerm = String(detail.detail.subject || message.subject || '').trim()
    expect(searchTerm).toBeTruthy()

    const beforeRemove = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=100`
    )
    expect(beforeRemove.page.total).toBeGreaterThan(0)

    const removal = await requestJson(`${started.baseUrl}/api/psts/remove`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })
    expect(removal.removed.scopePath).toBe('Case1/Search1')
    expect(removal.closedSessionIds).toContain(opened.sessionId)

    await expect(
      requestJson(`${started.baseUrl}/api/sessions/${opened.sessionId}`)
    ).rejects.toThrow('Session not found')

    const activeCatalog = await requestJson(`${started.baseUrl}/api/psts`)
    expect(activeCatalog.scopes).toHaveLength(0)
    expect(activeCatalog.files).toHaveLength(0)

    const removedCatalog = await requestJson(`${started.baseUrl}/api/psts/removed`)
    expect(removedCatalog.scopes.map((scope: { scopeLabel: string }) => scope.scopeLabel)).toEqual([
      'Case1 / Search1'
    ])
    expect(removedCatalog.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'sample.pst'
    ])

    const afterRemove = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=100`
    )
    expect(afterRemove.page.total).toBeLessThan(beforeRemove.page.total)
    expect(
      afterRemove.page.items.every((item: { fileName: string }) => item.fileName !== 'sample.pst')
    ).toBe(true)

    const restore = await requestJson(`${started.baseUrl}/api/psts/restore`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })
    expect(restore.restored.scopePath).toBe('Case1/Search1')
    expect(searchIndexRefreshSpy).not.toHaveBeenCalled()

    const restoredCatalog = await requestJson(`${started.baseUrl}/api/psts`)
    expect(restoredCatalog.scopes.map((scope: { scopeLabel: string }) => scope.scopeLabel)).toEqual([
      'Case1 / Search1'
    ])
    expect(restoredCatalog.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'sample.pst'
    ])

    const afterRestore = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=${encodeURIComponent(
        searchTerm
      )}&mailOnly=1&pageSize=100`
    )
    expect(afterRestore.page.total).toBe(afterRemove.page.total)
  })

  it('rejects concurrent search index refreshes while one is already running', async () => {
    rootDir = makeTempDir('pst-review-index-refresh-lock-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const firstReplaceStarted = createDeferred<void>()
    const releaseRefresh = createDeferred<void>()
    class BlockingSearchIndexStore extends MemorySearchIndexStore {
      private started = false

      override async replaceMailboxDocuments(
        mailboxKey: string,
        documents: Parameters<MemorySearchIndexStore['replaceMailboxDocuments']>[1]
      ): Promise<void> {
        if (!this.started) {
          this.started = true
          firstReplaceStarted.resolve()
          await releaseRefresh.promise
        }
        return super.replaceMailboxDocuments(mailboxKey, documents)
      }
    }

    const blockingSearchIndexStore = new BlockingSearchIndexStore()
    const started = await startApp(
      pstDir,
      undefined,
      {
        username: 'admin',
        password: 'pst-extractor',
        sessionTtlMinutes: 180
      },
      undefined,
      undefined,
      undefined,
      {
        searchIndexStore: blockingSearchIndexStore,
        skipInitialRefresh: true
      }
    )
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)

    const firstRefreshResponsePromise = fetch(`${started.baseUrl}/api/search/index/refresh?source=mailboxes`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair
      }
    })

    await firstReplaceStarted.promise

    const secondRefreshResponse = await fetch(`${started.baseUrl}/api/search/index/refresh?source=mailboxes`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair
      }
    })
    const secondRefreshPayload = await readJson(secondRefreshResponse)
    expect(secondRefreshResponse.status).toBe(409)
    expect(secondRefreshPayload.error).toBe('Search index refresh already in progress for mailboxes')

    releaseRefresh.resolve()
    const firstRefreshResponse = await firstRefreshResponsePromise
    const firstRefreshPayload = await readJson(firstRefreshResponse)
    expect(firstRefreshResponse.status).toBe(202)
    expect(firstRefreshPayload.status.status).toBe('running')
  })

  it('starts listening before the initial search index refresh finishes', async () => {
    rootDir = makeTempDir('pst-review-index-startup-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const firstReplaceStarted = createDeferred<void>()
    const releaseRefresh = createDeferred<void>()
    class BlockingSearchIndexStore extends MemorySearchIndexStore {
      private started = false

      override async replaceMailboxDocuments(
        mailboxKey: string,
        documents: Parameters<MemorySearchIndexStore['replaceMailboxDocuments']>[1]
      ): Promise<void> {
        if (!this.started) {
          this.started = true
          firstReplaceStarted.resolve()
          await releaseRefresh.promise
        }
        return super.replaceMailboxDocuments(mailboxKey, documents)
      }
    }

    const blockingSearchIndexStore = new BlockingSearchIndexStore()
    const startPromise = startApp(
      pstDir,
      undefined,
      {
        username: 'admin',
        password: 'pst-extractor',
        sessionTtlMinutes: 180
      },
      undefined,
      undefined,
      undefined,
      {
        searchIndexStore: blockingSearchIndexStore,
        backgroundInitialRefresh: true
      }
    )

    let started = false
    startPromise.then(() => {
      started = true
    })

    await firstReplaceStarted.promise
    await new Promise((resolve) => setTimeout(resolve, 50))
    expect(started).toBe(true)

    releaseRefresh.resolve()
    const app = await startPromise
    server = app.server
    reviewStore = app.reviewStore
    searchIndexStore = app.searchIndexStore
  })

  it('allows appointment items to be flagged and cleared', async () => {
    rootDir = makeTempDir('pst-review-api-appt-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(outlookPath, path.join(pstDir, 'Case1', 'Search1', 'archive.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'archive.ost'
      })
    })

    const calendarFolder = findFolderByDisplayName(opened.tree, 'Calendar')
    expect(calendarFolder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        calendarFolder!.id
      )}/messages?mailOnly=0&pageSize=20`
    )
    const appointment = folderPage.page.items.find(
      (item: { kind: string }) => item.kind === 'appointment'
    )
    expect(appointment).toBeTruthy()

    const patch = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        appointment.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true
        })
      }
    )
    expect(patch.review.flagged).toBe(true)

    const review = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        appointment.id
      )}/review`
    )
    expect(review.review.flagged).toBe(true)

    const queue = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/review?reviewFlagged=1`
    )
    expect(queue.items.map((item: { messageId: string }) => item.messageId)).toContain(
      appointment.id
    )

    const deleted = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        appointment.id
      )}/review`,
      {
        method: 'DELETE'
      }
    )
    expect(deleted.review.flagged).toBe(false)
  })

  it('exports a flagged mail and appointment bundle as a zip archive', async () => {
    rootDir = makeTempDir('pst-review-api-bundle-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(outlookPath, path.join(pstDir, 'Case1', 'Search1', 'archive.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'archive.ost'
      })
    })

    const mailFolder = findMailFolder(opened.tree)
    expect(mailFolder).toBeTruthy()
    const mailFolderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        mailFolder!.id
      )}/messages?pageSize=20`
    )
    const mailItem = mailFolderPage.page.items.find((item: { kind: string }) => item.kind === 'mail')
    expect(mailItem).toBeTruthy()

    await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        mailItem.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true
        })
      }
    )

    const calendarFolder = findFolderByDisplayName(opened.tree, 'Calendar')
    expect(calendarFolder).toBeTruthy()
    const calendarPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        calendarFolder!.id
      )}/messages?mailOnly=0&pageSize=20`
    )
    const appointment = calendarPage.page.items.find(
      (item: { kind: string }) => item.kind === 'appointment'
    )
    expect(appointment).toBeTruthy()

    await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        appointment.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true
        })
      }
    )

    const { response, buffer } = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=pst&sessionId=${opened.sessionId}`
    )
    expect(response.headers.get('content-type')).toContain('application/zip')
    const entries = parseStoredZipEntries(buffer)
    expect(entries.has('manifest.json')).toBe(true)
    const manifest = JSON.parse(entries.get('manifest.json')!.toString('utf8'))
    expect(manifest.scope.scope).toBe('pst')
    expect(manifest.counts.total).toBe(2)
    expect(manifest.counts.exported).toBe(2)
    expect(manifest.counts.failed).toBe(0)
    expect(manifest.items).toHaveLength(2)
    expect(manifest.items.some((item: { outputType: string }) => item.outputType === 'eml')).toBe(
      true
    )
    expect(manifest.items.some((item: { outputType: string }) => item.outputType === 'ics')).toBe(
      true
    )
    expect([...entries.keys()].some((name) => name.endsWith('.eml'))).toBe(true)
    expect([...entries.keys()].some((name) => name.endsWith('.ics'))).toBe(true)
    expect([...entries.keys()].some((name) => name.endsWith('.pst'))).toBe(false)
  })

  it('honors bundle scope when exporting flagged items', async () => {
    rootDir = makeTempDir('pst-review-api-bundle-scope-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search2'), { recursive: true })
    stageFixture(outlookPath, path.join(pstDir, 'Case1', 'Search1', 'first.ost'))
    stageFixture(outlookPath, path.join(pstDir, 'Case1', 'Search2', 'second.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const openedOne = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'first.ost'
      })
    })
    const openedTwo = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search2',
        fileName: 'second.ost'
      })
    })

    const firstFolder = findMailFolder(openedOne.tree)
    const secondFolder = findMailFolder(openedTwo.tree)
    expect(firstFolder).toBeTruthy()
    expect(secondFolder).toBeTruthy()

    const firstPage = await requestJson(
      `${started.baseUrl}/api/sessions/${openedOne.sessionId}/folders/${encodeURIComponent(
        firstFolder!.id
      )}/messages?pageSize=20`
    )
    const secondPage = await requestJson(
      `${started.baseUrl}/api/sessions/${openedTwo.sessionId}/folders/${encodeURIComponent(
        secondFolder!.id
      )}/messages?pageSize=20`
    )
    const firstMail = firstPage.page.items.find((item: { kind: string }) => item.kind === 'mail')
    const secondMail = secondPage.page.items.find((item: { kind: string }) => item.kind === 'mail')
    expect(firstMail).toBeTruthy()
    expect(secondMail).toBeTruthy()

    await requestJson(
      `${started.baseUrl}/api/sessions/${openedOne.sessionId}/messages/${encodeURIComponent(
        firstMail.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ flagged: true })
      }
    )
    await requestJson(
      `${started.baseUrl}/api/sessions/${openedTwo.sessionId}/messages/${encodeURIComponent(
        secondMail.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ flagged: true })
      }
    )

    const searchScopeBundle = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}`
    )
    const searchEntries = parseStoredZipEntries(searchScopeBundle.buffer)
    const searchManifest = JSON.parse(searchEntries.get('manifest.json')!.toString('utf8'))
    expect(searchManifest.scope.scope).toBe('search')
    expect(searchManifest.counts.total).toBe(1)
    expect(searchManifest.items).toHaveLength(1)
    expect(searchManifest.items[0].scopePath).toBe('Case1/Search1')
    expect([...searchEntries.keys()].some((name) => name.includes('first.ost'))).toBe(true)
    expect([...searchEntries.keys()].some((name) => name.includes('second.ost'))).toBe(false)

    const allScopeBundle = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=all`
    )
    const allEntries = parseStoredZipEntries(allScopeBundle.buffer)
    const allManifest = JSON.parse(allEntries.get('manifest.json')!.toString('utf8'))
    expect(allManifest.scope.scope).toBe('all')
    expect(allManifest.counts.total).toBe(2)
    expect(allManifest.items).toHaveLength(2)
    expect([...allEntries.keys()].some((name) => name.includes('first.ost'))).toBe(true)
    expect([...allEntries.keys()].some((name) => name.includes('second.ost'))).toBe(true)
  })

  it('flags archive items and exports them in the flagged bundle', async () => {
    rootDir = makeTempDir('pst-review-api-archive-bundle-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'mailbox.pst'))
    stageArchiveBundle(path.join(pstDir, 'Case1', 'Search1', 'Items.1.001.BONUS_AND_COMMISSION_DECISION_MAKING.zip'), [
      ['Exchange/Thread/TeamsMessagesData/chat.json', JSON.stringify({ subject: 'Launch', body: 'Teams chat' })],
      ['SharePoint/Docs/report.txt', 'Quarterly report'],
      ['SharePoint/Docs/deck.docx', 'Quarterly deck']
    ])
    fs.mkdirSync(path.join(pstDir, 'Case2', 'Search2'), { recursive: true })
    stageArchiveBundle(path.join(pstDir, 'Case2', 'Search2', 'Items.1.001.OTHER.zip'), [
      ['Exchange/Thread/TeamsMessagesData/chat.json', JSON.stringify({ subject: 'Launch', body: 'Other teams chat' })]
    ])

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const scopedTeamsSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=all&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=teams&query=launch&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    expect(scopedTeamsSearch.page.items).toHaveLength(1)
    expect(scopedTeamsSearch.page.scopePath).toBe('Case1/Search1')
    expect(scopedTeamsSearch.page.items[0].scopePath).toBe('Case1/Search1')

    const teamsSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=teams&query=launch&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    expect(teamsSearch.page.items).toHaveLength(1)
    expect(teamsSearch.page.items[0].sourceType).toBe('teams')

    const teamsReview = await requestJson(
      `${started.baseUrl}/api/items/${encodeURIComponent(teamsSearch.page.items[0].id)}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['Teams']
        })
      }
    )
    expect(teamsReview.review.flagged).toBe(true)
    expect(teamsReview.review.tags).toEqual(['Teams'])

    const sharepointSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=sharepoint&query=report&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    expect(sharepointSearch.page.items).toHaveLength(1)
    expect(sharepointSearch.page.items[0].sourceType).toBe('sharepoint')

    const officeSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=sharepoint&query=deck&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    const officeItem = officeSearch.page.items.find((item: { contentType?: string }) =>
      String(item.contentType || '').includes('openxmlformats')
    )
    expect(officeItem).toBeTruthy()
    if (!officeItem) {
      throw new Error('Expected an office archive item')
    }

    const officeDetail = await requestJson(
      `${started.baseUrl}/api/items/${encodeURIComponent(officeItem.id)}`
    )
    expect(officeDetail.detail.previewUrl).toContain('/api/items/')
    expect(officeDetail.detail.previewUrl).toContain('/preview')

    const officePreview = await fetch(`${started.baseUrl}${officeDetail.detail.previewUrl}`)
    const officePreviewBody = await officePreview.text()
    expect(officePreview.status).toBe(200)
    expect(['application/pdf', 'text/html; charset=utf-8']).toContain(
      officePreview.headers.get('content-type')
    )
    expect(officePreviewBody.length).toBeGreaterThan(0)

    const sharepointReview = await requestJson(
      `${started.baseUrl}/api/items/${encodeURIComponent(sharepointSearch.page.items[0].id)}/review`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['SharePoint']
        })
      }
    )
    expect(sharepointReview.review.flagged).toBe(true)
    expect(sharepointReview.review.tags).toEqual(['SharePoint'])

    const { response, buffer } = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}`
    )
    expect(response.headers.get('content-type')).toContain('application/zip')
    const entries = parseStoredZipEntries(buffer)
    const manifest = JSON.parse(entries.get('manifest.json')!.toString('utf8'))
    expect(manifest.counts.total).toBe(2)
    expect(manifest.counts.exported).toBe(2)
    expect(manifest.items.every((item: { sourceType: string }) => item.sourceType === 'archive')).toBe(
      true
    )
    expect(manifest.items.every((item: { outputType: string }) => item.outputType === 'raw')).toBe(true)
    expect([...entries.keys()].some((name) => name.endsWith('chat.json'))).toBe(true)
    expect([...entries.keys()].some((name) => name.endsWith('report.txt'))).toBe(true)
    expect([...entries.keys()].some((name) => name.endsWith('.eml'))).toBe(false)
    expect([...entries.keys()].some((name) => name.endsWith('.ics'))).toBe(false)

    const { response: csvResponse, buffer: csvBuffer } = await requestBuffer(
      `${started.baseUrl}/api/exports/items.csv`
    )
    expect(csvResponse.headers.get('content-type')).toContain('text/csv')
    const csvLines = csvBuffer
      .toString('utf8')
      .split(/\r?\n/g)
      .map((line) => line.trim())
      .filter(Boolean)
    expect(csvLines[0]).toContain('sourceType')
    expect(csvLines.some((line) => line.startsWith('"mailbox",'))).toBe(true)
    expect(csvLines.some((line) => line.startsWith('"teams",'))).toBe(true)
    expect(csvLines.some((line) => line.startsWith('"sharepoint",'))).toBe(true)
  })

  it('does not leak archive search results across case/search scopes', async () => {
    rootDir = makeTempDir('pst-review-api-archive-scope-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(pstDir, 'Case2', 'Search2'), { recursive: true })
    stageArchiveBundle(path.join(pstDir, 'Case2', 'Search2', 'Items.1.001.OTHER.zip'), [
      ['Exchange/Thread/TeamsMessagesData/chat.json', JSON.stringify({ subject: 'Launch', body: 'Scoped chat' })]
    ])

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const scopedTeamsSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=teams&query=launch&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    expect(scopedTeamsSearch.page.total).toBe(0)
    expect(scopedTeamsSearch.page.items).toHaveLength(0)

    const exactTeamsSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case2/Search2'
      )}&sourceType=teams&query=launch&mode=and&page=1&pageSize=20&mailOnly=0&sort=date-desc`
    )
    expect(exactTeamsSearch.page.total).toBe(1)
    expect(exactTeamsSearch.page.items).toHaveLength(1)
    expect(exactTeamsSearch.page.items[0].scopePath).toBe('Case2/Search2')
  })

  it('isolates review state and flagged bundles per authenticated user', async () => {
    rootDir = makeTempDir('pst-review-api-review-scope-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'review-scope.pst'))

    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' },
      { username: 'bob', password: 'bob-password' }
    ])
    const started = await startApp(
      pstDir,
      undefined,
      {
        username: 'admin',
        password: 'pst-extractor',
        sessionTtlMinutes: 180
      },
      authUserStore
    )
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    async function login(username: string, password: string): Promise<string> {
      const response = await fetch(`${started.baseUrl}/api/auth/login`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ username, password })
      })
      expect(response.status).toBe(200)
      return getCookiePair(getSetCookieHeader(response))
    }

    const adminCookie = await login('admin', 'pst-extractor')
    const bobCookie = await login('bob', 'bob-password')

    await requestJson(`${started.baseUrl}/api/auth/users/bob/access`, {
      method: 'PUT',
      headers: {
        Cookie: adminCookie,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        assignedCasePaths: ['Case1']
      })
    })

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        Cookie: adminCookie,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'review-scope.pst'
      })
    })

    const mailFolder = findMailFolder(opened.tree)
    expect(mailFolder).toBeTruthy()
    const mailFolderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        mailFolder!.id
      )}/messages?pageSize=20`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const mailItem = mailFolderPage.page.items.find((item: { kind: string }) => item.kind === 'mail')
    expect(mailItem).toBeTruthy()

    await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        mailItem.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          Cookie: adminCookie,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['Admin']
        })
      }
    )

    await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        mailItem.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          Cookie: bobCookie,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['Bob']
        })
      }
    )

    const adminReview = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        mailItem.id
      )}/review`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(adminReview.review.flagged).toBe(true)
    expect(adminReview.review.tags).toEqual(['Admin'])

    const bobReview = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        mailItem.id
      )}/review`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    expect(bobReview.review.flagged).toBe(true)
    expect(bobReview.review.tags).toEqual(['Bob'])

    const adminQueue = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/review?reviewFlagged=1`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(adminQueue.items.map((item: { messageId: string }) => item.messageId)).toContain(
      mailItem.id
    )
    expect(
      adminQueue.items.every((item: { reviewerUsername?: string }) => item.reviewerUsername === 'admin')
    ).toBe(true)

    const bobQueue = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/review?reviewFlagged=1`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    expect(bobQueue.items.map((item: { messageId: string }) => item.messageId)).toContain(
      mailItem.id
    )
    expect(
      bobQueue.items.every((item: { reviewerUsername?: string }) => item.reviewerUsername === 'bob')
    ).toBe(true)

    const adminSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&reviewFlagged=1&pageSize=20`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(adminSearch.page.total).toBeGreaterThan(0)
    expect(
      adminSearch.page.items.every((item: { review: { tags: string[] } }) =>
        item.review.tags.includes('Admin')
      )
    ).toBe(true)

    const bobSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=pst&sessionId=${opened.sessionId}&reviewFlagged=1&pageSize=20`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    expect(bobSearch.page.total).toBeGreaterThan(0)
    expect(
      bobSearch.page.items.every((item: { review: { tags: string[] } }) =>
        item.review.tags.includes('Bob')
      )
    ).toBe(true)

    const adminBundle = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=pst&sessionId=${opened.sessionId}`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const adminEntries = parseStoredZipEntries(adminBundle.buffer)
    const adminManifest = JSON.parse(adminEntries.get('manifest.json')!.toString('utf8'))
    expect(adminManifest.counts.total).toBe(1)
    expect(adminManifest.items).toHaveLength(1)
    expect(adminManifest.items[0].review.tags).toEqual(['Admin'])

    const bobBundle = await requestBuffer(
      `${started.baseUrl}/api/exports/flagged.zip?scope=pst&sessionId=${opened.sessionId}`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    const bobEntries = parseStoredZipEntries(bobBundle.buffer)
    const bobManifest = JSON.parse(bobEntries.get('manifest.json')!.toString('utf8'))
    expect(bobManifest.counts.total).toBe(1)
    expect(bobManifest.items).toHaveLength(1)
    expect(bobManifest.items[0].review.tags).toEqual(['Bob'])

    const adminItemsCsvBeforeClear = await requestBuffer(
      `${started.baseUrl}/api/exports/items.csv`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const adminItemsCsvBeforeClearText = adminItemsCsvBeforeClear.buffer.toString('utf8')
    const adminMailRowBeforeClear = adminItemsCsvBeforeClearText
      .split(/\r?\n/g)
      .find((line) => line.includes(`"${mailItem.id}"`))
    expect(adminItemsCsvBeforeClearText.split(/\r?\n/g)[0]).toContain('flagged')
    if (!adminMailRowBeforeClear) {
      throw new Error('Expected mail row in items CSV before clearing flags')
    }
    expect(adminMailRowBeforeClear).toContain(',"true",')

    const clearFlags = await requestJson(
      `${started.baseUrl}/api/reviews/clear-flags?workspaceMode=folder&sessionId=${opened.sessionId}&folderId=${encodeURIComponent(
        mailFolder!.id
      )}`,
      {
        method: 'POST',
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(clearFlags.clearedCount).toBeGreaterThan(0)

    const adminFolderAfterClear = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        mailFolder!.id
      )}/messages?page=1&pageSize=20`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const adminFolderItemAfterClear = adminFolderAfterClear.page.items.find(
      (item: { id: string }) => item.id === mailItem.id
    )
    if (!adminFolderItemAfterClear) {
      throw new Error('Expected mail item in admin folder after clearing flags')
    }
    expect(adminFolderItemAfterClear.review.flagged).toBe(false)

    const bobFolderAfterClear = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        mailFolder!.id
      )}/messages?page=1&pageSize=20`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    const bobFolderItemAfterClear = bobFolderAfterClear.page.items.find(
      (item: { id: string }) => item.id === mailItem.id
    )
    if (!bobFolderItemAfterClear) {
      throw new Error('Expected mail item in bob folder after clearing flags')
    }
    expect(bobFolderItemAfterClear.review.flagged).toBe(true)

    const adminItemsCsvAfterClear = await requestBuffer(
      `${started.baseUrl}/api/exports/items.csv`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const adminItemsCsvAfterClearText = adminItemsCsvAfterClear.buffer.toString('utf8')
    const adminMailRowAfterClear = adminItemsCsvAfterClearText
      .split(/\r?\n/g)
      .find((line) => line.includes(`"${mailItem.id}"`))
    if (!adminMailRowAfterClear) {
      throw new Error('Expected mail row in items CSV after clearing flags')
    }
    expect(adminMailRowAfterClear).toContain(',"false",')

    const bobItemsCsvAfterClear = await requestBuffer(
      `${started.baseUrl}/api/exports/items.csv`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    const bobItemsCsvAfterClearText = bobItemsCsvAfterClear.buffer.toString('utf8')
    const bobMailRowAfterClear = bobItemsCsvAfterClearText
      .split(/\r?\n/g)
      .find((line) => line.includes(`"${mailItem.id}"`))
    if (!bobMailRowAfterClear) {
      throw new Error('Expected mail row in items CSV for bob after clearing flags')
    }
    expect(bobMailRowAfterClear).toContain(',"true",')

    const auditEntries = readAuditLogEntries(started.auditLogPath)
    expect(auditEntries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          action: 'message.review.update',
          actor: expect.objectContaining({ username: 'admin' }),
          metadata: expect.objectContaining({ reviewerUsername: 'admin' })
        }),
        expect.objectContaining({
          action: 'message.review.update',
          actor: expect.objectContaining({ username: 'bob' }),
          metadata: expect.objectContaining({ reviewerUsername: 'bob' })
        }),
        expect.objectContaining({
          action: 'bundle.export',
          actor: expect.objectContaining({ username: 'admin' }),
          metadata: expect.objectContaining({ reviewerUsername: 'admin' })
        }),
        expect.objectContaining({
          action: 'bundle.export',
          actor: expect.objectContaining({ username: 'bob' }),
          metadata: expect.objectContaining({ reviewerUsername: 'bob' })
        })
      ])
    )
  })

  it('opens PST root mailboxes when scopePath is empty', async () => {
    rootDir = makeTempDir('pst-review-api-root-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(outlookPath, path.join(pstDir, 'root.ost'))

    const started = await startApp(pstDir)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

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
    searchIndexStore = started.searchIndexStore

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
    searchIndexStore = started.searchIndexStore

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
    searchIndexStore = started.searchIndexStore

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

  it('skips the external auth middleware for protected routes after local sign-in', async () => {
    rootDir = makeTempDir('pst-review-api-local-auth-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const authCheck = jest.fn(async (req, res, next) => {
      if (!req.headers['x-graph-token']) {
        return res.status(401).json({
          success: false,
          message: 'No proof of possession.'
        })
      }
      return next()
    })
    const started = await startApp(pstDir, {
      bypassIps: [],
      m365Auth: {
        CheckTokens: authCheck
      }
    }, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    })
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)

    const catalog = await requestJson(`${started.baseUrl}/api/psts`, {
      headers: {
        Cookie: cookiePair
      }
    })
    expect(catalog.files).toHaveLength(1)

    const filters = await requestJson(`${started.baseUrl}/api/search/filters`, {
      headers: {
        Cookie: cookiePair
      }
    })
    expect(filters.items).toEqual([])
    expect(authCheck).not.toHaveBeenCalled()
  })

  it('authenticates the default viewer account and protects session cookies', async () => {
    rootDir = makeTempDir('pst-review-api-auth-login-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    })
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const failedLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'wrong-password'
      })
    })
    const failedLoginPayload = await readJson(failedLoginResponse)
    expect(failedLoginResponse.status).toBe(401)
    expect(failedLoginPayload.error).toBe('Invalid username or password')

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const setCookie = getSetCookieHeader(loginResponse)
    const cookiePair = getCookiePair(setCookie)

    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)
    expect(loginPayload.enabled).toBe(true)
    expect(loginPayload.canManageUsers).toBe(true)
    expect(loginPayload.mfaEnabled).toBe(false)
    expect(loginPayload.mfaEnforced).toBe(false)
    expect(loginPayload.user.username).toBe('admin')
    expect(setCookie).toContain('HttpOnly')
    expect(cookiePair).toContain('pst-review-session=')

    const meResponse = await fetch(`${started.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const mePayload = await readJson(meResponse)
    expect(meResponse.status).toBe(200)
    expect(mePayload.authenticated).toBe(true)
    expect(mePayload.canManageUsers).toBe(true)
    expect(mePayload.mfaEnabled).toBe(false)
    expect(mePayload.mfaEnforced).toBe(false)
    expect(mePayload.user.username).toBe('admin')

    const emailInviteResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'jane.doe',
        recipientEmail: 'jane.doe@example.com'
      })
    })
    const emailInvitePayload = await readJson(emailInviteResponse)
    expect(emailInviteResponse.status).toBe(200)
    const emailInviteToken = String(new URL(emailInvitePayload.inviteUrl).pathname.split('/').pop() || '')

    const acceptEmailInviteResponse = await fetch(
      `${started.baseUrl}/api/auth/invites/${encodeURIComponent(emailInviteToken)}/accept`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          password: 'jane-password',
          confirmPassword: 'jane-password'
        })
      }
    )
    const acceptEmailInvitePayload = await readJson(acceptEmailInviteResponse)
    expect(acceptEmailInviteResponse.status).toBe(200)
    expect(acceptEmailInvitePayload.user.username).toBe('jane.doe')

    const emailLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'jane.doe@example.com',
        password: 'jane-password'
      })
    })
    const emailLoginPayload = await readJson(emailLoginResponse)
    const emailLoginCookiePair = getCookiePair(getSetCookieHeader(emailLoginResponse))
    expect(emailLoginResponse.status).toBe(200)
    expect(emailLoginPayload.authenticated).toBe(true)
    expect(emailLoginPayload.user.username).toBe('jane.doe')
    expect(emailLoginCookiePair).toContain('pst-review-session=')

    const catalog = await requestJson(`${started.baseUrl}/api/psts`, {
      headers: {
        Cookie: cookiePair
      }
    })
    expect(catalog.files).toHaveLength(1)

    const logoutResponse = await fetch(`${started.baseUrl}/api/auth/logout`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair
      }
    })
    const logoutPayload = await readJson(logoutResponse)
    expect(logoutResponse.status).toBe(200)
    expect(logoutPayload.authenticated).toBe(false)

    const postLogoutMeResponse = await fetch(`${started.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const postLogoutMePayload = await readJson(postLogoutMeResponse)
    expect(postLogoutMeResponse.status).toBe(401)
    expect(postLogoutMePayload.error).toBe('Authentication required')
  })

  it('adds additional viewer users and persists them across restarts', async () => {
    rootDir = makeTempDir('pst-review-api-auth-users-')
    const pstDir = path.join(rootDir, 'PST')
    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' }
    ])
    const appSettingsStore = createMemoryAppSettingsStore({
      enabled: true,
      host: 'smtp.example.test',
      port: 587,
      secure: false,
      username: 'smtp-user',
      password: 'smtp-secret',
      fromName: 'DV PST Mail Explorer',
      fromAddress: 'noreply@example.test',
      replyTo: 'support@example.test'
    })
    const sentMessages: Array<Record<string, unknown>> = []
    const smtpTransportFactory = () => ({
      sendMail: async (message: Record<string, unknown>) => {
        sentMessages.push(message)
        return {
          messageId: 'invite-message-id',
          accepted: [String(message.to || '')],
          rejected: []
        }
      },
      close: async () => undefined
    })
    fs.mkdirSync(pstDir)
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180,
      publicBaseUrl: 'https://portal.example.test'
    }, authUserStore, appSettingsStore, smtpTransportFactory)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const unauthUsersResponse = await fetch(`${started.baseUrl}/api/auth/users`)
    const unauthUsersPayload = await readJson(unauthUsersResponse)
    expect(unauthUsersResponse.status).toBe(401)
    expect(unauthUsersPayload.error).toBe('Authentication required')

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)
    expect(loginPayload.canManageUsers).toBe(true)
    expect(loginPayload.mfaEnabled).toBe(false)

    const createResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'alice',
        recipientEmail: 'alice@example.com'
      })
    })
    const createPayload = await readJson(createResponse)
    expect(createResponse.status).toBe(200)
    expect(createPayload.user.username).toBe('alice')
    expect(createPayload.user.inviteStatus).toBe('pending')
    expect(createPayload.emailSent).toBe(true)
    expect(createPayload.inviteUrl).toContain('/invite/')
    expect(new URL(createPayload.inviteUrl).origin).toBe('https://portal.example.test')
    expect(sentMessages.length).toBe(1)
    expect(String(sentMessages[0].html || '')).toContain('Click here to setup your access to DV PST Mail Explorer')
    expect(String(sentMessages[0].text || '')).toContain('Click here to setup your access to DV PST Mail Explorer')

    await requestJson(`${started.baseUrl}/api/auth/users/alice/access`, {
      method: 'PUT',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        assignedCasePaths: ['Case1']
      })
    })

    const enforceResponse = await fetch(`${started.baseUrl}/api/auth/users/alice/mfa/enforce`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        enforced: true
      })
    })
    const enforcePayload = await readJson(enforceResponse)
    expect(enforceResponse.status).toBe(200)
    expect(enforcePayload.user.username).toBe('alice')
    expect(enforcePayload.user.mfaEnforced).toBe(true)

    const inviteToken = String(new URL(createPayload.inviteUrl).pathname.split('/').pop() || '')
    const inviteLookupResponse = await fetch(
      `${started.baseUrl}/api/auth/invites/${encodeURIComponent(inviteToken)}`
    )
    const inviteLookupPayload = await readJson(inviteLookupResponse)
    expect(inviteLookupResponse.status).toBe(200)
    expect(inviteLookupPayload.invite.username).toBe('alice')
    expect(inviteLookupPayload.invite.inviteStatus).toBe('pending')
    expect(inviteLookupPayload.invite.mfaEnforced).toBe(true)

    const acceptResponse = await fetch(
      `${started.baseUrl}/api/auth/invites/${encodeURIComponent(inviteToken)}/accept`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          password: 'secret123',
          confirmPassword: 'secret123'
        })
      }
    )
    const acceptPayload = await readJson(acceptResponse)
    const aliceCookiePair = getCookiePair(getSetCookieHeader(acceptResponse))
    expect(acceptResponse.status).toBe(200)
    expect(acceptPayload.user.username).toBe('alice')
    expect(acceptPayload.user.mfaEnforced).toBe(true)
    expect(acceptPayload.mfaAvailable).toBe(true)

    const inviteLookupAfterAcceptResponse = await fetch(
      `${started.baseUrl}/api/auth/invites/${encodeURIComponent(inviteToken)}`
    )
    const inviteLookupAfterAcceptPayload = await readJson(inviteLookupAfterAcceptResponse)
    expect(inviteLookupAfterAcceptResponse.status).toBe(404)
    expect(inviteLookupAfterAcceptPayload.error).toBe('Invite not found')

    const inviteAcceptAgainResponse = await fetch(
      `${started.baseUrl}/api/auth/invites/${encodeURIComponent(inviteToken)}/accept`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          password: 'secret123',
          confirmPassword: 'secret123'
        })
      }
    )
    const inviteAcceptAgainPayload = await readJson(inviteAcceptAgainResponse)
    expect(inviteAcceptAgainResponse.status).toBe(409)
    expect(inviteAcceptAgainPayload.error).toBe('Invite already accepted')

    const mfaStartResponse = await fetch(`${started.baseUrl}/api/auth/mfa/enrollment/start`, {
      method: 'POST',
      headers: {
        Cookie: aliceCookiePair
      }
    })
    const mfaStartPayload = await readJson(mfaStartResponse)
    expect(mfaStartResponse.status).toBe(200)
    expect(mfaStartPayload.secret).toBeTruthy()
    expect(mfaStartPayload.qrCodeDataUrl).toContain('data:image/png;base64,')

    const totpCode = generateTotpCode(mfaStartPayload.secret)
    const mfaCompleteResponse = await fetch(
      `${started.baseUrl}/api/auth/mfa/enrollment/complete`,
      {
        method: 'POST',
        headers: {
          Cookie: aliceCookiePair,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          code: totpCode
        })
      }
    )
    const mfaCompletePayload = await readJson(mfaCompleteResponse)
    expect(mfaCompleteResponse.status).toBe(200)
    expect(mfaCompletePayload.user.username).toBe('alice')
    expect(mfaCompletePayload.recoveryCodes.length).toBeGreaterThan(0)

    const usersResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const usersPayload = await readJson(usersResponse)
    expect(usersResponse.status).toBe(200)
    expect(usersPayload.users.map((user: { username: string }) => user.username)).toEqual(
      expect.arrayContaining(['admin', 'alice'])
    )
    expect(
      usersPayload.users.find((user: { username: string }) => user.username === 'alice')
    ).toEqual(
      expect.objectContaining({
        inviteStatus: 'active',
        mfaEnabled: true,
        mfaEnforced: true
      })
    )

    const aliceLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'alice',
        password: 'secret123'
      })
    })
    const aliceLoginPayload = await readJson(aliceLoginResponse)
    const aliceChallengeCookiePair = getCookiePair(getSetCookieHeader(aliceLoginResponse))
    expect(aliceLoginResponse.status).toBe(200)
    expect(aliceLoginPayload.authenticated).toBe(false)
    expect(aliceLoginPayload.mfaRequired).toBe(true)
    expect(aliceLoginPayload.mfaEnabled).toBe(false)
    expect(aliceLoginPayload.user.username).toBe('alice')

    const aliceChallengeResponse = await fetch(`${started.baseUrl}/api/auth/mfa/challenge`, {
      method: 'POST',
      headers: {
        Cookie: aliceChallengeCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        code: generateTotpCode(mfaStartPayload.secret)
      })
    })
    const aliceChallengePayload = await readJson(aliceChallengeResponse)
    const aliceSessionCookiePair = getCookiePair(getSetCookieHeader(aliceChallengeResponse))
    expect(aliceChallengeResponse.status).toBe(200)
    expect(aliceChallengePayload.authenticated).toBe(true)
    expect(aliceChallengePayload.mfaEnabled).toBe(true)
    expect(aliceChallengePayload.mfaEnforced).toBe(true)
    expect(aliceChallengePayload.user.username).toBe('alice')

    const aliceMeResponse = await fetch(`${started.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: aliceSessionCookiePair
      }
    })
    const aliceMePayload = await readJson(aliceMeResponse)
    expect(aliceMeResponse.status).toBe(200)
    expect(aliceMePayload.authenticated).toBe(true)
    expect(aliceMePayload.mfaEnabled).toBe(true)
    expect(aliceMePayload.mfaEnforced).toBe(true)
    expect(aliceMePayload.user.username).toBe('alice')

    const aliceRecoveryLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'alice',
        password: 'secret123'
      })
    })
    const aliceRecoveryLoginPayload = await readJson(aliceRecoveryLoginResponse)
    const aliceRecoveryChallengeCookiePair = getCookiePair(getSetCookieHeader(aliceRecoveryLoginResponse))
    expect(aliceRecoveryLoginResponse.status).toBe(200)
    expect(aliceRecoveryLoginPayload.mfaRequired).toBe(true)

    const recoveryCode = String(mfaCompletePayload.recoveryCodes[0] || '')
    const aliceRecoveryChallengeResponse = await fetch(
      `${started.baseUrl}/api/auth/mfa/challenge`,
      {
        method: 'POST',
        headers: {
          Cookie: aliceRecoveryChallengeCookiePair,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          code: recoveryCode
        })
      }
    )
    const aliceRecoveryChallengePayload = await readJson(aliceRecoveryChallengeResponse)
    const aliceRecoveryCookiePair = getCookiePair(getSetCookieHeader(aliceRecoveryChallengeResponse))
    expect(aliceRecoveryChallengeResponse.status).toBe(200)
    expect(aliceRecoveryChallengePayload.authenticated).toBe(true)
    expect(aliceRecoveryChallengePayload.user.username).toBe('alice')

    const aliceEnforceResponse = await fetch(`${started.baseUrl}/api/auth/users/alice/mfa/enforce`, {
      method: 'POST',
      headers: {
        Cookie: aliceCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        enforced: false
      })
    })
    const aliceEnforcePayload = await readJson(aliceEnforceResponse)
    expect(aliceEnforceResponse.status).toBe(403)
    expect(aliceEnforcePayload.error).toBe('Admin access required')

    const aliceOpenResponse = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        Cookie: aliceRecoveryCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: 'Case1/Search1',
        fileName: 'sample.pst'
      })
    })
    expect(aliceOpenResponse.sessionId).toBeTruthy()

    const aliceRefreshResponse = await fetch(`${started.baseUrl}/api/search/index/refresh?source=mailboxes`, {
      method: 'POST',
      headers: {
        Cookie: aliceRecoveryCookiePair
      }
    })
    const aliceRefreshPayload = await readJson(aliceRefreshResponse)
    expect(aliceRefreshResponse.status).toBe(403)
    expect(aliceRefreshPayload.error).toBe('Admin access required')

    const aliceFilteredActivityLogResponse = await fetch(
      `${started.baseUrl}/api/activity-log?username=alice&limit=20`,
      {
        headers: {
          Cookie: cookiePair
        }
      }
    )
    const aliceFilteredActivityLogPayload = await readJson(aliceFilteredActivityLogResponse)
    expect(aliceFilteredActivityLogResponse.status).toBe(200)
    expect(aliceFilteredActivityLogPayload.entries.length).toBeGreaterThan(0)
    expect(
      aliceFilteredActivityLogPayload.entries.every((entry: { actor?: { username?: string } }) =>
        entry.actor?.username === 'alice'
      )
    ).toBe(true)
    expect(aliceFilteredActivityLogPayload.entries.map((entry: { action: string }) => entry.action)).toEqual(
      expect.arrayContaining(['auth.login', 'mailbox.open'])
    )

    const aliceActivityLogResponse = await fetch(`${started.baseUrl}/api/activity-log`, {
      headers: {
        Cookie: aliceCookiePair
      }
    })
    const aliceActivityLogPayload = await readJson(aliceActivityLogResponse)
    expect(aliceActivityLogResponse.status).toBe(403)
    expect(aliceActivityLogPayload.error).toBe('Admin access required')

    const aliceUsersResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      headers: {
        Cookie: aliceCookiePair
      }
    })
    const aliceUsersPayload = await readJson(aliceUsersResponse)
    expect(aliceUsersResponse.status).toBe(403)
    expect(aliceUsersPayload.error).toBe('Admin access required')

    const aliceCreateResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      method: 'POST',
      headers: {
        Cookie: aliceCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'charlie',
        recipientEmail: 'charlie@example.com'
      })
    })
    const aliceCreatePayload = await readJson(aliceCreateResponse)
    expect(aliceCreateResponse.status).toBe(403)
    expect(aliceCreatePayload.error).toBe('Admin access required')

    await new Promise<void>((resolveClose) => {
      started.server.close(() => resolveClose())
    })
    server = null
    await started.reviewStore.close()
    reviewStore = null
    await started.searchIndexStore.close()
    searchIndexStore = null

    const restarted = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    }, authUserStore)
    server = restarted.server
    reviewStore = restarted.reviewStore
    searchIndexStore = restarted.searchIndexStore

    const aliceRestartLoginResponse = await fetch(`${restarted.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'alice',
        password: 'secret123'
      })
    })
    const aliceRestartLoginPayload = await readJson(aliceRestartLoginResponse)
    expect(aliceRestartLoginResponse.status).toBe(200)
    expect(aliceRestartLoginPayload.authenticated).toBe(false)
    expect(aliceRestartLoginPayload.mfaRequired).toBe(true)
    expect(aliceRestartLoginPayload.mfaEnabled).toBe(false)
    expect(aliceRestartLoginPayload.mfaEnforced).toBe(true)
    expect(aliceRestartLoginPayload.user.username).toBe('alice')

    const aliceRestartChallengeResponse = await fetch(
      `${restarted.baseUrl}/api/auth/mfa/challenge`,
      {
        method: 'POST',
        headers: {
          Cookie: getCookiePair(getSetCookieHeader(aliceRestartLoginResponse)),
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          code: generateTotpCode(mfaStartPayload.secret)
        })
      }
    )
    const aliceRestartChallengePayload = await readJson(aliceRestartChallengeResponse)
    const aliceRestartSessionCookiePair = getCookiePair(getSetCookieHeader(aliceRestartChallengeResponse))
    expect(aliceRestartChallengeResponse.status).toBe(200)
    expect(aliceRestartChallengePayload.authenticated).toBe(true)
    expect(aliceRestartChallengePayload.mfaEnabled).toBe(true)
    expect(aliceRestartChallengePayload.mfaEnforced).toBe(true)
    expect(aliceRestartChallengePayload.user.username).toBe('alice')

    const adminRestartLoginResponse = await fetch(`${restarted.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const adminRestartLoginPayload = await readJson(adminRestartLoginResponse)
    const adminRestartCookiePair = getCookiePair(getSetCookieHeader(adminRestartLoginResponse))
    expect(adminRestartLoginResponse.status).toBe(200)
    expect(adminRestartLoginPayload.authenticated).toBe(true)
    expect(adminRestartLoginPayload.canManageUsers).toBe(true)
    expect(adminRestartLoginPayload.mfaEnabled).toBe(false)
    expect(adminRestartLoginPayload.mfaEnforced).toBe(false)

    const deleteAliceResponse = await fetch(`${restarted.baseUrl}/api/auth/users/alice`, {
      method: 'DELETE',
      headers: {
        Cookie: adminRestartCookiePair
      }
    })
    const deleteAlicePayload = await readJson(deleteAliceResponse)
    expect(deleteAliceResponse.status).toBe(200)
    expect(deleteAlicePayload.user.username).toBe('alice')

    const aliceMeAfterDeleteResponse = await fetch(`${restarted.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: aliceRestartSessionCookiePair
      }
    })
    const aliceMeAfterDeletePayload = await readJson(aliceMeAfterDeleteResponse)
    expect(aliceMeAfterDeleteResponse.status).toBe(401)
    expect(aliceMeAfterDeletePayload.error).toBe('Authentication required')

    const aliceLoginAfterDeleteResponse = await fetch(`${restarted.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'alice',
        password: 'secret123'
      })
    })
    const aliceLoginAfterDeletePayload = await readJson(aliceLoginAfterDeleteResponse)
    expect(aliceLoginAfterDeleteResponse.status).toBe(401)
    expect(aliceLoginAfterDeletePayload.error).toBe('Invalid username or password')

    const usersAfterDeleteResponse = await fetch(`${restarted.baseUrl}/api/auth/users`, {
      headers: {
        Cookie: adminRestartCookiePair
      }
    })
    const usersAfterDeletePayload = await readJson(usersAfterDeleteResponse)
    expect(usersAfterDeleteResponse.status).toBe(200)
    expect(usersAfterDeletePayload.users.map((user: { username: string }) => user.username)).toEqual(
      expect.not.arrayContaining(['alice'])
    )
  })

  it('supports self-service mfa enrollment after a non-admin login', async () => {
    rootDir = makeTempDir('pst-review-api-self-service-mfa-')
    const pstDir = path.join(rootDir, 'PST')
    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' }
    ])
    await authUserStore.addUser('bob', 'secret123')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    }, authUserStore)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const bobLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'bob',
        password: 'secret123'
      })
    })
    const bobLoginPayload = await readJson(bobLoginResponse)
    const bobCookiePair = getCookiePair(getSetCookieHeader(bobLoginResponse))
    expect(bobLoginResponse.status).toBe(200)
    expect(bobLoginPayload.authenticated).toBe(true)
    expect(bobLoginPayload.mfaEnabled).toBe(false)

    const bobMfaStartResponse = await fetch(`${started.baseUrl}/api/auth/mfa/enrollment/start`, {
      method: 'POST',
      headers: {
        Cookie: bobCookiePair
      }
    })
    const bobMfaStartPayload = await readJson(bobMfaStartResponse)
    expect(bobMfaStartResponse.status).toBe(200)
    expect(bobMfaStartPayload.secret).toBeTruthy()

    const bobMfaCompleteResponse = await fetch(
      `${started.baseUrl}/api/auth/mfa/enrollment/complete`,
      {
        method: 'POST',
        headers: {
          Cookie: bobCookiePair,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          code: generateTotpCode(bobMfaStartPayload.secret)
        })
      }
    )
    const bobMfaCompletePayload = await readJson(bobMfaCompleteResponse)
    expect(bobMfaCompleteResponse.status).toBe(200)
    expect(bobMfaCompletePayload.user.username).toBe('bob')
    expect(bobMfaCompletePayload.user.mfaEnabled).toBe(true)

    const bobMeResponse = await fetch(`${started.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: bobCookiePair
      }
    })
    const bobMePayload = await readJson(bobMeResponse)
    expect(bobMeResponse.status).toBe(200)
    expect(bobMePayload.authenticated).toBe(true)
    expect(bobMePayload.mfaEnabled).toBe(true)

    const bobLoginAfterMfaResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'bob',
        password: 'secret123'
      })
    })
    const bobLoginAfterMfaPayload = await readJson(bobLoginAfterMfaResponse)
    expect(bobLoginAfterMfaResponse.status).toBe(200)
    expect(bobLoginAfterMfaPayload.authenticated).toBe(false)
    expect(bobLoginAfterMfaPayload.mfaRequired).toBe(true)
    expect(bobLoginAfterMfaPayload.mfaEnabled).toBe(false)

    const bobChallengeResponse = await fetch(`${started.baseUrl}/api/auth/mfa/challenge`, {
      method: 'POST',
      headers: {
        Cookie: getCookiePair(getSetCookieHeader(bobLoginAfterMfaResponse)),
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        code: generateTotpCode(bobMfaStartPayload.secret)
      })
    })
    const bobChallengePayload = await readJson(bobChallengeResponse)
    expect(bobChallengeResponse.status).toBe(200)
    expect(bobChallengePayload.authenticated).toBe(true)
    expect(bobChallengePayload.mfaEnabled).toBe(true)
  })

  it('restricts catalog and search access to assigned cases before any mailbox is opened', async () => {
    rootDir = makeTempDir('pst-review-api-case-access-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(path.join(pstDir, 'Case1', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(pstDir, 'Case2', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(pstDir, 'Case1', 'Search1', 'case1.pst'))
    stageFixture(enronPath, path.join(pstDir, 'Case2', 'Search1', 'case2.pst'))

    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' },
      { username: 'bob', password: 'password123' }
    ])

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    }, authUserStore)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const adminLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const adminCookie = getCookiePair(getSetCookieHeader(adminLoginResponse))
    expect(adminLoginResponse.status).toBe(200)

    const bobLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'bob',
        password: 'password123'
      })
    })
    const bobCookie = getCookiePair(getSetCookieHeader(bobLoginResponse))
    expect(bobLoginResponse.status).toBe(200)

    const emptyCatalogResponse = await requestJson(`${started.baseUrl}/api/psts`, {
      headers: {
        Cookie: bobCookie
      }
    })
    expect(emptyCatalogResponse.scopes).toEqual([])
    expect(emptyCatalogResponse.files).toEqual([])
    expect(emptyCatalogResponse.message).toBe('No cases assigned.')

    const accessResponse = await fetch(`${started.baseUrl}/api/auth/users/bob/access`, {
      method: 'PUT',
      headers: {
        Cookie: adminCookie,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        assignedCasePaths: ['Case1']
      })
    })
    const accessPayload = await readJson(accessResponse)
    expect(accessResponse.status).toBe(200)
    expect(accessPayload.user.assignedCasePaths).toEqual(['Case1'])

    const catalogResponse = await requestJson(`${started.baseUrl}/api/psts`, {
      headers: {
        Cookie: bobCookie
      }
    })
    expect(catalogResponse.scopes.map((scope: { scopePath: string }) => scope.scopePath)).toEqual([
      'Case1/Search1'
    ])
    expect(catalogResponse.files.map((file: { fileName: string }) => file.fileName)).toEqual([
      'case1.pst'
    ])

    const disallowedCatalogResponse = await fetch(
      `${started.baseUrl}/api/psts?scopePath=${encodeURIComponent('Case2/Search1')}`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    const disallowedCatalogPayload = await readJson(disallowedCatalogResponse)
    expect(disallowedCatalogResponse.status).toBe(403)
    expect(disallowedCatalogPayload.error).toBe('Case access required')

    const searchBeforeOpenResponse = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=sample&pageSize=5`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    expect(searchBeforeOpenResponse.scope).toBe('all')
    expect(searchBeforeOpenResponse.page).toBeTruthy()

    const disallowedSearchResponse = await fetch(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case2/Search1'
      )}&query=sample&pageSize=5`,
      {
        headers: {
          Cookie: bobCookie
        }
      }
    )
    const disallowedSearchPayload = await readJson(disallowedSearchResponse)
    expect(disallowedSearchResponse.status).toBe(403)
    expect(disallowedSearchPayload.error).toBe('Case access required')
  })

  it('searches from persisted fingerprints without crawling the PST tree', async () => {
    rootDir = makeTempDir('pst-review-api-search-fingerprints-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir, { recursive: true })

    const fingerprintStore = new MemorySearchIndexStore()
    const mailboxKey = path.resolve(pstDir, 'Case1', 'Search1', 'alpha.pst')
    const teamsKey = path.resolve(pstDir, 'Case1', 'Search1', 'Items.1.001.TEAMS.zip')

    await fingerprintStore.replaceMailboxDocuments(mailboxKey, [
      makeSearchIndexDocument({
        mailboxKey,
        fileName: 'alpha.pst',
        mailboxName: 'Alpha',
        messageId: 'message:mail-1',
        descriptorId: 'mail-1',
        folderId: 'folder:mail-1',
        folderPath: 'Inbox',
        subject: 'Signature note',
        originalSubject: 'Signature note',
        senderName: 'Alice Example',
        senderEmailAddress: 'alice@example.com',
        recipientText: 'Bob Example <bob@example.com>',
        displayTo: 'Bob Example <bob@example.com>',
        resolvedDisplayTo: 'Bob Example <bob@example.com>',
        bodySearchText: 'signature note',
        searchText: 'signature note alice example alice@example.com bob example bob@example.com ipm.note mail',
        searchTokens: ['signature', 'note'],
        addressValues: ['alice@example.com', 'bob@example.com'],
        subjectValues: ['signature note'],
        sourceType: 'mailbox'
      })
    ])
    await fingerprintStore.replaceMailboxDocuments(teamsKey, [
      makeSearchIndexDocument({
        mailboxKey: teamsKey,
        fileName: 'Items.1.001.TEAMS.zip',
        mailboxName: 'Teams Bundle',
        messageId: 'message:teams-1',
        descriptorId: 'teams-1',
        folderId: 'folder:teams-1',
        folderPath: 'Teams',
        subject: 'Launch plan',
        originalSubject: 'Launch plan',
        senderName: 'Team Bot',
        senderEmailAddress: 'bot@example.com',
        recipientText: '',
        displayTo: '',
        resolvedDisplayTo: '',
        bodySearchText: 'launch plan',
        searchText: 'launch plan team bot bot@example.com ipm.note mail',
        searchTokens: ['launch', 'plan'],
        addressValues: ['bot@example.com'],
        subjectValues: ['launch plan'],
        sourceType: 'teams',
        kind: 'other'
      })
    ])
    await fingerprintStore.replaceFileFingerprints('mailboxes', [
      makeSearchIndexFingerprint({
        source: 'mailboxes',
        mailboxKey,
        fileName: 'alpha.pst',
        scopePath: 'Case1/Search1',
        scopeLabel: 'Case1 / Search1',
        size: 1024,
        modifiedAt: null
      })
    ])
    await fingerprintStore.replaceFileFingerprints('items', [
      makeSearchIndexFingerprint({
        source: 'items',
        mailboxKey: teamsKey,
        fileName: 'Items.1.001.TEAMS.zip',
        scopePath: 'Case1/Search1',
        scopeLabel: 'Case1 / Search1',
        size: 1024,
        modifiedAt: null
      })
    ])

    const started = await startApp(pstDir, undefined, undefined, undefined, undefined, undefined, {
      searchIndexStore: fingerprintStore,
      skipInitialRefresh: true
    })
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const mailboxSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=signature&mailOnly=1&pageSize=20`
    )
    expect(mailboxSearch.scope).toBe('all')
    expect(mailboxSearch.page.total).toBe(1)
    expect(mailboxSearch.page.sourceCounts.mailbox).toBe(1)
    expect(mailboxSearch.page.sourceCounts.teams).toBe(0)
    expect(mailboxSearch.page.items[0].messageId).toBe('message:mail-1')
    expect(mailboxSearch.page.items[0].scopePath).toBe('Case1/Search1')
    expect(mailboxSearch.page.items[0].previewText).toBeUndefined()

    const teamsSearch = await requestJson(
      `${started.baseUrl}/api/search?scope=search&scopePath=${encodeURIComponent(
        'Case1/Search1'
      )}&sourceType=teams&query=launch&pageSize=20`
    )
    expect(teamsSearch.scope).toBe('search')
    expect(teamsSearch.page.scopePath).toBe('Case1/Search1')
    expect(teamsSearch.page.scopeLabel).toBe('Case1 / Search1')
    expect(teamsSearch.page.total).toBe(1)
    expect(teamsSearch.page.sourceCounts.teams).toBe(1)
    expect(teamsSearch.page.items[0].sourceType).toBe('teams')
    expect(teamsSearch.page.items[0].messageId).toBe('message:teams-1')
    expect(teamsSearch.page.items[0].previewHtml).toBeUndefined()
  })

  it('builds invite links from the request origin when no public base url is configured', async () => {
    rootDir = makeTempDir('pst-review-api-auth-invite-origin-')
    const pstDir = path.join(rootDir, 'PST')
    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' }
    ])
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    }, authUserStore)
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)

    const inviteResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'bob',
        recipientEmail: 'bob@example.com'
      })
    })
    const invitePayload = await readJson(inviteResponse)
    expect(inviteResponse.status).toBe(200)
    expect(new URL(invitePayload.inviteUrl).origin).toBe(started.baseUrl)
  })

  it('serves SMTP settings, preserves passwords, and sends test emails', async () => {
    rootDir = makeTempDir('pst-review-api-smtp-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const authUserStore = createMemoryAuthUserStore([
      { username: 'admin', password: 'pst-extractor' },
      { username: 'bob', password: 'password123' }
    ])
    const settingsStore = createMemoryAppSettingsStore({
      enabled: true,
      host: 'smtp.initial.local',
      port: 2525,
      secure: false,
      username: 'smtp-user',
      password: 'smtp-secret',
      fromName: 'Initial Sender',
      fromAddress: 'notify@example.com',
      replyTo: 'reply@example.com'
    })
    const sentMessages: Array<{
      settings: Record<string, unknown>
      message: Record<string, unknown>
    }> = []
    const smtpTransportFactory = (settings: Record<string, unknown>) => ({
      async sendMail(message: Record<string, unknown>) {
        sentMessages.push({
          settings,
          message
        })

        if (String(message.to || '').includes('fail@example.com')) {
          throw new Error('simulated transport failure')
        }

        return {
          messageId: 'smtp-test-message-id',
          accepted: [String(message.to || '')],
          rejected: []
        }
      },
      async close() {}
    })

    const started = await startApp(
      pstDir,
      undefined,
      {
        username: 'admin',
        password: 'pst-extractor',
        sessionTtlMinutes: 180
      },
      authUserStore,
      settingsStore,
      smtpTransportFactory
    )
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)

    const smtpSettingsResponse = await fetch(`${started.baseUrl}/api/settings/smtp`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const smtpSettingsPayload = await readJson(smtpSettingsResponse)
    expect(smtpSettingsResponse.status).toBe(200)
    expect(smtpSettingsPayload.settings.enabled).toBe(true)
    expect(smtpSettingsPayload.settings.host).toBe('smtp.initial.local')
    expect(smtpSettingsPayload.settings.hasPassword).toBe(true)
    expect(smtpSettingsPayload.settings.password).toBeUndefined()

    const updateResponse = await fetch(`${started.baseUrl}/api/settings/smtp`, {
      method: 'PUT',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        enabled: true,
        host: 'smtp.updated.local',
        port: 587,
        secure: true,
        username: 'smtp-user-updated',
        password: '',
        fromName: 'Updated Sender',
        fromAddress: 'updated@example.com',
        replyTo: 'reply-updated@example.com'
      })
    })
    const updatePayload = await readJson(updateResponse)
    expect(updateResponse.status).toBe(200)
    expect(updatePayload.settings.host).toBe('smtp.updated.local')
    expect(updatePayload.settings.hasPassword).toBe(true)
    expect(updatePayload.settings.password).toBeUndefined()

    const storedSettings = await settingsStore.getSmtpSettings()
    expect(storedSettings.password).toBe('smtp-secret')
    expect(storedSettings.host).toBe('smtp.updated.local')

    const testSendResponse = await fetch(`${started.baseUrl}/api/settings/smtp/test`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        enabled: true,
        host: 'smtp.updated.local',
        port: 587,
        secure: true,
        username: 'smtp-user-updated',
        password: '',
        fromName: 'Updated Sender',
        fromAddress: 'updated@example.com',
        replyTo: 'reply-updated@example.com',
        recipient: 'recipient@example.com'
      })
    })
    const testSendPayload = await readJson(testSendResponse)
    expect(testSendResponse.status).toBe(200)
    expect(testSendPayload.success).toBe(true)
    expect(testSendPayload.recipient).toBe('recipient@example.com')
    expect(sentMessages).toHaveLength(1)
    expect(sentMessages[0].settings.host).toBe('smtp.updated.local')
    expect(sentMessages[0].settings.password).toBe('smtp-secret')
    expect(String(sentMessages[0].message.from || '')).toContain('Updated Sender')
    expect(String(sentMessages[0].message.to || '')).toBe('recipient@example.com')

    const failedTestSendResponse = await fetch(`${started.baseUrl}/api/settings/smtp/test`, {
      method: 'POST',
      headers: {
        Cookie: cookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        recipient: 'fail@example.com'
      })
    })
    const failedTestSendPayload = await readJson(failedTestSendResponse)
    expect(failedTestSendResponse.status).toBe(502)
    expect(failedTestSendPayload.error).toContain('Unable to send test email')

    const bobLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'bob',
        password: 'password123'
      })
    })
    const bobCookiePair = getCookiePair(getSetCookieHeader(bobLoginResponse))
    expect(bobLoginResponse.status).toBe(200)

    const bobSettingsResponse = await fetch(`${started.baseUrl}/api/settings/smtp`, {
      headers: {
        Cookie: bobCookiePair
      }
    })
    const bobSettingsPayload = await readJson(bobSettingsResponse)
    expect(bobSettingsResponse.status).toBe(403)
    expect(bobSettingsPayload.error).toBe('Admin access required')

    const bobSettingsUpdateResponse = await fetch(`${started.baseUrl}/api/settings/smtp`, {
      method: 'PUT',
      headers: {
        Cookie: bobCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        host: 'smtp.blocked.local'
      })
    })
    const bobSettingsUpdatePayload = await readJson(bobSettingsUpdateResponse)
    expect(bobSettingsUpdateResponse.status).toBe(403)
    expect(bobSettingsUpdatePayload.error).toBe('Admin access required')

    const bobTestSendResponse = await fetch(`${started.baseUrl}/api/settings/smtp/test`, {
      method: 'POST',
      headers: {
        Cookie: bobCookiePair,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        recipient: 'recipient@example.com'
      })
    })
    const bobTestSendPayload = await readJson(bobTestSendResponse)
    expect(bobTestSendResponse.status).toBe(403)
    expect(bobTestSendPayload.error).toBe('Admin access required')

    await new Promise<void>((resolveClose) => {
      started.server.close(() => resolveClose())
    })
    server = null
    await started.reviewStore.close()
    reviewStore = null
    await started.searchIndexStore.close()
    searchIndexStore = null

    const restarted = await startApp(
      pstDir,
      undefined,
      {
        username: 'admin',
        password: 'pst-extractor',
        sessionTtlMinutes: 180
      },
      authUserStore,
      settingsStore,
      smtpTransportFactory
    )
    server = restarted.server
    reviewStore = restarted.reviewStore
    searchIndexStore = restarted.searchIndexStore

    const restartedLoginResponse = await fetch(`${restarted.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const restartedCookiePair = getCookiePair(getSetCookieHeader(restartedLoginResponse))
    expect(restartedLoginResponse.status).toBe(200)

    const restartedSettingsResponse = await fetch(`${restarted.baseUrl}/api/settings/smtp`, {
      headers: {
        Cookie: restartedCookiePair
      }
    })
    const restartedSettingsPayload = await readJson(restartedSettingsResponse)
    expect(restartedSettingsResponse.status).toBe(200)
    expect(restartedSettingsPayload.settings.host).toBe('smtp.updated.local')
    expect(restartedSettingsPayload.settings.hasPassword).toBe(true)
  })

  it('writes activity log entries to disk and replays them after restart', async () => {
    rootDir = makeTempDir('pst-review-api-activity-log-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    })
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const failedLoginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'wrong-password'
      })
    })
    const failedLoginPayload = await readJson(failedLoginResponse)
    expect(failedLoginResponse.status).toBe(401)
    expect(failedLoginPayload.error).toBe('Invalid username or password')
    expect(fs.existsSync(started.auditLogPath)).toBe(true)
    expect(readAuditLogEntries(started.auditLogPath)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          action: 'auth.login',
          outcome: 'denied'
        })
      ])
    )

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const adminCookie = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)

    const usersResponse = await fetch(`${started.baseUrl}/api/auth/users`, {
      headers: {
        Cookie: adminCookie
      }
    })
    const usersPayload = await readJson(usersResponse)
    expect(usersResponse.status).toBe(200)
    expect(usersPayload.users.map((user: { username: string }) => user.username)).toEqual(
      expect.arrayContaining(['admin'])
    )

    const opened = await requestJson(`${started.baseUrl}/api/psts/open`, {
      method: 'POST',
      headers: {
        Cookie: adminCookie,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        scopePath: '',
        fileName: 'sample.pst'
      })
    })
    const folder = findMailFolder(opened.tree)
    expect(folder).toBeTruthy()

    const folderPage = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/folders/${encodeURIComponent(
        folder!.id
      )}/messages?page=1&pageSize=10`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(folderPage.page.items.length).toBeGreaterThan(0)

    const message = folderPage.page.items.find((item: { isMailLike: boolean }) => item.isMailLike) ||
      folderPage.page.items[0]
    expect(message).toBeTruthy()

    const detail = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(detail.detail).toBeTruthy()

    const searchResults = await requestJson(
      `${started.baseUrl}/api/search?scope=all&query=sample&pageSize=5`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(searchResults.page).toBeTruthy()

    const reviewResponse = await requestJson(
      `${started.baseUrl}/api/sessions/${opened.sessionId}/messages/${encodeURIComponent(
        message.id
      )}/review`,
      {
        method: 'PATCH',
        headers: {
          Cookie: adminCookie,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          flagged: true,
          tags: ['audit']
        })
      }
    )
    expect(reviewResponse.review.flagged).toBe(true)

    const refreshResponse = await fetch(`${started.baseUrl}/api/search/index/refresh?source=mailboxes`, {
      method: 'POST',
      headers: {
        Cookie: adminCookie
      }
    })
    const refreshPayload = await readJson(refreshResponse)
    expect(refreshResponse.status).toBe(202)
    expect(refreshPayload.status.status).toBe('running')

    await waitForRefreshStatus(started.baseUrl, adminCookie)

    const activityLogResponse = await requestJson(
      `${started.baseUrl}/api/activity-log?limit=20`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    expect(activityLogResponse.entries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ action: 'auth.login', outcome: 'denied' }),
        expect.objectContaining({ action: 'auth.login', outcome: 'success' }),
        expect.objectContaining({ action: 'auth.users.list', outcome: 'success' }),
        expect.objectContaining({ action: 'mailbox.open', outcome: 'success' }),
        expect.objectContaining({ action: 'folder.view', outcome: 'success' }),
        expect.objectContaining({ action: 'message.view', outcome: 'success' }),
        expect.objectContaining({ action: 'search.execute', outcome: 'success' }),
        expect.objectContaining({ action: 'message.review.update', outcome: 'success' }),
        expect.objectContaining({ action: 'search.index.refresh.mailboxes', outcome: 'success' })
      ])
    )
    expect(activityLogResponse.entries[0].action).toBe('search.index.refresh.mailboxes')
    expect(activityLogResponse.entries.some((entry: any) =>
      entry.action === 'message.review.update' && entry.metadata?.flagged === true
    )).toBe(true)

    const activityLogCsvResponse = await fetch(
      `${started.baseUrl}/api/activity-log.csv?username=admin`,
      {
        headers: {
          Cookie: adminCookie
        }
      }
    )
    const activityLogCsvText = await activityLogCsvResponse.text()
    expect(activityLogCsvResponse.status).toBe(200)
    expect(activityLogCsvResponse.headers.get('content-type')).toContain('text/csv')
    expect(activityLogCsvResponse.headers.get('content-disposition')).toContain('activity-log-admin.csv')
    expect(activityLogCsvText).toContain('timestamp,actorUsername,actorAuthenticated,actorAdmin')
    expect(activityLogCsvText).toContain('auth.login')
    expect(activityLogCsvText).toContain('search.index.refresh.mailboxes')
    expect(readAuditLogEntries(started.auditLogPath)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          action: 'activity.log.export',
          outcome: 'success'
        })
      ])
    )

    await new Promise<void>((resolveClose) => {
      started.server.close(() => resolveClose())
    })
    server = null
    await started.reviewStore.close()
    reviewStore = null
    await started.searchIndexStore.close()
    searchIndexStore = null

    const restarted = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 180
    })
    server = restarted.server
    reviewStore = restarted.reviewStore
    searchIndexStore = restarted.searchIndexStore

    const restartedLoginResponse = await fetch(`${restarted.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const restartedCookie = getCookiePair(getSetCookieHeader(restartedLoginResponse))
    expect(restartedLoginResponse.status).toBe(200)

    const restartedActivityLogResponse = await requestJson(
      `${restarted.baseUrl}/api/activity-log?limit=20`,
      {
        headers: {
          Cookie: restartedCookie
        }
      }
    )
    expect(restartedActivityLogResponse.entries).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ action: 'mailbox.open', outcome: 'success' }),
        expect.objectContaining({ action: 'search.index.refresh.mailboxes', outcome: 'success' })
      ])
    )
    expect(readAuditLogEntries(started.auditLogPath).length).toBeGreaterThanOrEqual(
      restartedActivityLogResponse.entries.length
    )
  })

  it('expires viewer sessions after the configured ttl', async () => {
    rootDir = makeTempDir('pst-review-api-auth-expiry-')
    const pstDir = path.join(rootDir, 'PST')
    fs.mkdirSync(pstDir)
    stageFixture(enronPath, path.join(pstDir, 'sample.pst'))

    const started = await startApp(pstDir, undefined, {
      username: 'admin',
      password: 'pst-extractor',
      sessionTtlMinutes: 0.02
    })
    server = started.server
    reviewStore = started.reviewStore
    searchIndexStore = started.searchIndexStore

    const loginResponse = await fetch(`${started.baseUrl}/api/auth/login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        username: 'admin',
        password: 'pst-extractor'
      })
    })
    const loginPayload = await readJson(loginResponse)
    const cookiePair = getCookiePair(getSetCookieHeader(loginResponse))
    expect(loginResponse.status).toBe(200)
    expect(loginPayload.authenticated).toBe(true)

    await new Promise((resolve) => setTimeout(resolve, 1500))

    const expiredMeResponse = await fetch(`${started.baseUrl}/api/auth/me`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const expiredMePayload = await readJson(expiredMeResponse)
    expect(expiredMeResponse.status).toBe(401)
    expect(expiredMePayload.error).toBe('Authentication required')

    const expiredCatalogResponse = await fetch(`${started.baseUrl}/api/psts`, {
      headers: {
        Cookie: cookiePair
      }
    })
    const expiredCatalogPayload = await readJson(expiredCatalogResponse)
    expect(expiredCatalogResponse.status).toBe(401)
    expect(expiredCatalogPayload.error).toBe('Authentication required')
  })
})
