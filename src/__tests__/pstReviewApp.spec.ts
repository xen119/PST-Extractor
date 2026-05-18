import * as fs from 'fs'
import * as http from 'http'
import * as os from 'os'
import * as path from 'path'
import { AddressInfo } from 'net'
import { buildOpenApiDocument } from '../openApi'
import { createPstReviewApp, type ApiSecurityConfig } from '../pstReviewApp'
import { MemoryReviewStore } from '../reviewStore'
import { MemorySearchIndexStore, refreshSearchIndexFromCatalog } from '../searchIndex'

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

async function startApp(pstRootDir: string, apiSecurity?: ApiSecurityConfig) {
  const reviewStore = new MemoryReviewStore()
  const searchIndexStore = new MemorySearchIndexStore()
  await refreshSearchIndexFromCatalog(pstRootDir, reviewStore, searchIndexStore)
  const app = createPstReviewApp({
    publicDir,
    pstRootDir,
    reviewStore,
    searchIndexStore,
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
    searchIndexStore,
    server
  }
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

    const refreshedIndex = await requestJson(`${started.baseUrl}/api/search/index/refresh`, {
      method: 'POST'
    })
    expect(refreshedIndex.summary.mailboxCount).toBe(2)
    expect(Array.isArray(refreshedIndex.hiddenRules)).toBe(true)

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
    expect(openApi.paths['/api/exports/flagged.zip']).toBeDefined()

    const docsResponse = await fetch(`${started.baseUrl}/api/docs`)
    const docsHtml = await docsResponse.text()
    expect(docsHtml).toContain('SwaggerUIBundle')
    expect(docsHtml).toContain('/api/openapi.json')
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
})
