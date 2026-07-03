import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import AdmZip from 'adm-zip'
import {
  MemorySearchIndexStore,
  MongoSearchIndexStore,
  refreshSearchIndexFromCatalog,
  type SearchIndexDocument
} from '../searchIndex'
import type { ReviewStore } from '../reviewStore'
import { extractArchiveBundleItems, listArchiveBundleFiles } from '../archiveBundles'

function makeDocument(overrides: Partial<SearchIndexDocument> = {}): SearchIndexDocument {
  const now = new Date('2024-01-01T00:00:00.000Z').toISOString()
  const base: SearchIndexDocument = {
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
    bodySearchText: 'please send the signature today',
    searchText:
      'project alpha re project alpha alice example alice@example.com bob example bob@example.com please send the signature today ipm.note mail',
    searchTokens: ['project', 'alpha', 're', 'alice', 'example', 'signature', 'today'],
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

  return base
}

describe('search index cache', () => {
  it('supports AND/OR search and exact hide rules from the cached documents', async () => {
    const store = new MemorySearchIndexStore()
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/alpha.pst', [
      makeDocument(),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/beta.pst',
        fileName: 'beta.pst',
        mailboxName: 'Beta',
        messageId: 'message:2',
        descriptorId: '2',
        folderId: 'folder:2',
        folderPath: 'Inbox',
        order: 2,
        subject: 'Budget Review',
        originalSubject: 'Budget Review',
        senderName: 'Carol Example',
        senderEmailAddress: 'carol@example.com',
        recipientText: 'Dan Example <dan@example.com>',
        displayTo: 'Dan Example <dan@example.com>',
        resolvedDisplayTo: 'Dan Example <dan@example.com>',
        bodySearchText: 'quarterly numbers attached',
        searchText:
          'budget review carol example carol@example.com dan example dan@example.com quarterly numbers attached ipm.note mail',
        searchTokens: ['budget', 'review', 'carol', 'example', 'quarterly', 'numbers', 'attached'],
        addressValues: ['carol@example.com', 'dan@example.com'],
        subjectValues: ['budget review'],
        sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
        updatedAt: new Date('2024-01-02T00:00:00.000Z').toISOString()
      })
    ])

    const andResults = await store.search({
      scope: 'all',
      query: '"Project Alpha" signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(andResults.total).toBe(1)
    expect(andResults.items[0].messageId).toBe('message:1')

    const orResults = await store.search({
      scope: 'all',
      query: 'missing signature',
      mode: 'or',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(orResults.total).toBe(1)
    expect(orResults.items[0].messageId).toBe('message:1')

    const andSyntaxResults = await store.search({
      scope: 'all',
      query: 'Project Alpha + signature',
      mode: 'or',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(andSyntaxResults.total).toBe(1)
    expect(andSyntaxResults.items[0].messageId).toBe('message:1')

    const orSyntaxResults = await store.search({
      scope: 'all',
      query: 'missing | signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(orSyntaxResults.total).toBe(1)
    expect(orSyntaxResults.items[0].messageId).toBe('message:1')

    const selectedScope = await store.search({
      scope: 'search',
      scopePath: 'Case1/Search1',
      query: 'budget',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(selectedScope.total).toBe(1)
    expect(selectedScope.scopeLabel).toBe('Case1 / Search1')

    const selectedMailbox = await store.search({
      scope: 'pst',
      mailboxKey: 'C:/PST/Case1/Search1/alpha.pst',
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(selectedMailbox.total).toBe(1)
    expect(selectedMailbox.scopeLabel).toBe('Selected PST')

    const activeMailboxOnly = await store.search({
      scope: 'all',
      allowedMailboxKeys: ['C:/PST/Case1/Search1/beta.pst'],
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(activeMailboxOnly.total).toBe(0)

    const activeMailboxMatch = await store.search({
      scope: 'all',
      allowedMailboxKeys: ['C:/PST/Case1/Search1/alpha.pst'],
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(activeMailboxMatch.total).toBe(1)
    expect(activeMailboxMatch.items[0].mailboxKey).toBe('C:/PST/Case1/Search1/alpha.pst')

    const hiddenRule = await store.upsertHiddenRule({
      kind: 'subject',
      value: 'Project Alpha',
      label: 'Project Alpha'
    })
    expect(hiddenRule.value).toBe('project alpha')

    const hiddenResults = await store.search({
      scope: 'all',
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(hiddenResults.total).toBe(0)
    expect(hiddenResults.hiddenRules).toHaveLength(1)

    const addressRule = await store.upsertHiddenRule({
      kind: 'address',
      value: 'bob@example.com',
      label: 'bob@example.com'
    })
    expect(addressRule.value).toBe('bob@example.com')

    const addressHidden = await store.search({
      scope: 'all',
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(addressHidden.total).toBe(0)

    expect(await store.deleteHiddenRule(hiddenRule.filterId)).toBe(true)
    expect(await store.deleteHiddenRule(addressRule.filterId)).toBe(true)

    const restored = await store.search({
      scope: 'all',
      query: 'signature',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(restored.total).toBe(1)
    expect(restored.items[0].messageId).toBe('message:1')
  })

  it('resolves review filters per reviewer username', async () => {
    const store = new MemorySearchIndexStore()
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/alpha.pst', [
      makeDocument({
        review: {
          flagged: false,
          tags: [],
          createdAt: '',
          updatedAt: ''
        },
        reviewStates: [
          {
            reviewerUsername: 'admin',
            review: {
              flagged: true,
              tags: ['Admin'],
              createdAt: '2024-01-03T00:00:00.000Z',
              updatedAt: '2024-01-03T00:00:00.000Z'
            }
          },
          {
            reviewerUsername: 'bob',
            review: {
              flagged: false,
              tags: ['Bob'],
              createdAt: '2024-01-04T00:00:00.000Z',
              updatedAt: '2024-01-04T00:00:00.000Z'
            }
          }
        ]
      })
    ])

    const adminResults = await store.search({
      scope: 'all',
      query: 'project',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewerUsername: 'admin',
      reviewFlaggedOnly: true,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(adminResults.total).toBe(1)
    expect(adminResults.items[0].review.flagged).toBe(true)
    expect(adminResults.items[0].review.tags).toEqual(['Admin'])

    const bobResults = await store.search({
      scope: 'all',
      query: 'project',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewerUsername: 'bob',
      reviewFlaggedOnly: true,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(bobResults.total).toBe(0)
    expect(bobResults.items).toHaveLength(0)
  })

  it('discovers and indexes top-level items bundles as searchable archive corpora', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-index-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(
        scopeDir,
        'Items.1.001.BONUS_AND_COMMISSION_DECISION_MAKING.zip'
      )

      const nestedZip = new AdmZip()
      nestedZip.addFile(
        'Exchange/Thread/TeamsMessagesData/chat.txt',
        Buffer.from('Teams chat about the launch plan', 'utf8')
      )

      const bundleZip = new AdmZip()
      bundleZip.addFile(
        'SharePoint/Docs/report.txt',
        Buffer.from('SharePoint quarterly report', 'utf8')
      )
      bundleZip.addFile('Exchange/Team/nested.zip', nestedZip.toBuffer())
      bundleZip.writeZip(bundlePath)

      const catalog = listArchiveBundleFiles(rootDir)
      expect(catalog.scopes).toHaveLength(1)
      expect(catalog.scopes[0].fileCount).toBe(1)

      const items = await extractArchiveBundleItems(bundlePath, 'Case1/Search1')
      expect(items).toHaveLength(2)
      expect(items.some((item) => item.sourceType === 'teams')).toBe(true)
      expect(items.some((item) => item.sourceType === 'sharepoint')).toBe(true)

      const reviewStore = {
        kind: 'memory' as const,
        isPersistent: false,
        async listReviews() {
          return []
        }
      } as ReviewStore

      const store = new MemorySearchIndexStore()
      const summary = await refreshSearchIndexFromCatalog(rootDir, reviewStore, store)
      expect(summary.messageCount).toBe(2)

      const teamsResults = await store.search({
        scope: 'all',
        sourceType: 'teams',
        query: 'launch',
        mode: 'and',
        mailOnly: false,
        sort: 'date-desc',
        page: 1,
        pageSize: 20,
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      })
      expect(teamsResults.total).toBe(1)
      expect(teamsResults.items[0].sourceType).toBe('teams')
      expect(teamsResults.sourceCounts.teams).toBe(1)

      const sharepointResults = await store.search({
        scope: 'all',
        sourceType: 'sharepoint',
        query: 'quarterly',
        mode: 'and',
        mailOnly: false,
        sort: 'date-desc',
        page: 1,
        pageSize: 20,
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      })
      expect(sharepointResults.total).toBe(1)
      expect(sharepointResults.items[0].sourceType).toBe('sharepoint')
      expect(sharepointResults.sourceCounts.sharepoint).toBe(1)
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('pretty prints JSON preview text from Teams archives', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-json-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(scopeDir, 'Items.1.001.TEAMS.zip')
      const bundleZip = new AdmZip()
      bundleZip.addFile(
        'Exchange/Thread/TeamsMessagesData/chat.json',
        Buffer.from(JSON.stringify({ subject: 'Launch', body: 'Teams chat' }), 'utf8')
      )
      bundleZip.writeZip(bundlePath)

      const items = await extractArchiveBundleItems(bundlePath, 'Case1/Search1')
      expect(items).toHaveLength(1)
      expect(items[0].sourceType).toBe('teams')
      expect(items[0].contentType).toBe('application/json; charset=utf-8')
      expect(items[0].previewText).toBe('{\n  "subject": "Launch",\n  "body": "Teams chat"\n}')
      expect(items[0].searchText).toBe(items[0].previewText)
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('classifies bmp archives as images for preview delivery', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-bmp-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(scopeDir, 'Items.1.001.IMAGES.zip')
      const bundleZip = new AdmZip()
      bundleZip.addFile('SharePoint/Images/diagram.bmp', Buffer.from('bmp-bytes', 'utf8'))
      bundleZip.writeZip(bundlePath)

      const items = await extractArchiveBundleItems(bundlePath, 'Case1/Search1')
      expect(items).toHaveLength(1)
      expect(items[0].contentType).toBe('image/bmp')
      expect(items[0].previewKind).toBe('binary')
      expect(items[0].downloadFilename).toBe('diagram.bmp')
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('falls back to tolerant archive reading when trailing bytes break strict zip parsing', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-fallback-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(scopeDir, 'Items.1.001.BROKEN.zip')
      const bundleZip = new AdmZip()
      bundleZip.addFile('SharePoint/Docs/good.txt', Buffer.from('good content', 'utf8'))
      bundleZip.writeZip(bundlePath)
      fs.appendFileSync(bundlePath, Buffer.from('JUNKJUNKJUNK'))

      const items = await extractArchiveBundleItems(bundlePath, 'Case1/Search1')
      expect(items).toHaveLength(1)
      expect(items[0].entryName).toBe('good.txt')
      expect(items[0].previewText).toBe('good content')
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('skips unreadable archive entries instead of failing the whole bundle', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-skip-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(scopeDir, 'Items.1.001.FAILSAFE.zip')
      const bundleZip = new AdmZip()
      bundleZip.addFile('SharePoint/Docs/good.txt', Buffer.from('good content', 'utf8'))
      bundleZip.addFile('SharePoint/Docs/bad.zip', Buffer.from('not a zip archive', 'utf8'))
      bundleZip.writeZip(bundlePath)

      const items = await extractArchiveBundleItems(bundlePath, 'Case1/Search1')
      expect(items).toHaveLength(1)
      expect(items[0].entryName).toBe('good.txt')
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('uses aggregate counts and clamps paging in the Mongo search store', async () => {
    const now = new Date('2024-01-01T00:00:00.000Z').toISOString()
    const mailboxDocument = makeDocument({
      messageId: 'message:mailbox-1',
      descriptorId: 'mailbox-1',
      folderId: 'folder:mailbox-1',
      fileName: 'mailbox-a.pst',
      mailboxName: 'Mailbox A',
      sortDateMs: Date.parse(now),
      sortDate: now,
      updatedAt: now
    })
    const teamsDocument = makeDocument({
      messageId: 'message:teams-1',
      descriptorId: 'teams-1',
      folderId: 'folder:teams-1',
      fileName: 'items.zip',
      mailboxName: 'Teams Bundle',
      sourceType: 'teams',
      kind: 'other',
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
      sortDateMs: Date.parse('2023-12-31T00:00:00.000Z'),
      sortDate: '2023-12-31T00:00:00.000Z',
      updatedAt: '2023-12-31T00:00:00.000Z'
    })

    const cursorState = { skip: 0, limit: 0 }
    const toArray = jest.fn(async () =>
      [mailboxDocument, teamsDocument].slice(cursorState.skip, cursorState.skip + cursorState.limit)
    )
    const limit = jest.fn((count: number) => {
      cursorState.limit = count
      return { toArray }
    })
    const skip = jest.fn((count: number) => {
      cursorState.skip = count
      return { limit }
    })
    const sort = jest.fn(() => ({ skip }))
    const find = jest.fn(() => ({ sort }))
    const aggregate = jest.fn(() => ({
      toArray: jest.fn().mockResolvedValue([
        {
          total: [{ value: 2 }],
          sourceCounts: [
            { _id: 'mailbox', count: 1 },
            { _id: 'teams', count: 1 }
          ]
        }
      ])
    }))
    const countDocuments = jest.fn().mockResolvedValue(0)

    const documentsCollection = {
      aggregate,
      countDocuments,
      find,
      findOne: jest.fn().mockResolvedValue(null),
      updateOne: jest.fn().mockResolvedValue(undefined),
      insertMany: jest.fn().mockResolvedValue(undefined),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any
    const rulesCollection = {
      find: jest.fn(() => ({
        sort: jest.fn(() => ({
          toArray: jest.fn().mockResolvedValue([])
        }))
      })),
      findOne: jest.fn().mockResolvedValue(null),
      updateOne: jest.fn().mockResolvedValue(undefined),
      deleteOne: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any
    const fingerprintsCollection = {
      find: jest.fn(() => ({
        sort: jest.fn(() => ({
          toArray: jest.fn().mockResolvedValue([])
        }))
      })),
      findOne: jest.fn().mockResolvedValue(null),
      updateOne: jest.fn().mockResolvedValue(undefined),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any

    const store = new MongoSearchIndexStore(documentsCollection, rulesCollection, fingerprintsCollection)
    const results = await store.search({
      scope: 'all',
      sourceType: 'all',
      query: '',
      mode: 'and',
      mailOnly: false,
      sort: 'date-desc',
      page: 5,
      pageSize: 1,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })

    expect(aggregate).toHaveBeenCalledTimes(1)
    expect(countDocuments).not.toHaveBeenCalled()
    expect(find).toHaveBeenCalledTimes(1)
    expect(skip).toHaveBeenCalledWith(1)
    expect(limit).toHaveBeenCalledWith(1)
    expect(results.total).toBe(2)
    expect(results.page).toBe(2)
    expect(results.totalPages).toBe(2)
    expect(results.items).toHaveLength(1)
    expect(results.items[0].messageId).toBe('message:teams-1')
    expect(results.sourceCounts).toEqual({
      mailbox: 1,
      teams: 1,
      sharepoint: 0
    })
  })
})
