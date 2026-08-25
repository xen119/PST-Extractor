import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import AdmZip from 'adm-zip'
import {
  MemorySearchIndexStore,
  MongoSearchIndexStore,
  buildMailboxThreadMetadata,
  refreshSearchIndexFromCatalog,
  refreshSearchIndexSourceFromCatalog,
  type SearchIndexDocument
} from '../searchIndex'
import type { ReviewStore } from '../reviewStore'
import { extractArchiveBundleItems, listArchiveBundleFiles, readArchiveBundleItemContent } from '../archiveBundles'

const enronPath = path.resolve('./src/__tests__/testdata/enron.pst')

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
      size: 1024 * 1024,
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

function stageFixture(sourcePath: string, targetPath: string): void {
  try {
    fs.linkSync(sourcePath, targetPath)
  } catch {
    fs.copyFileSync(sourcePath, targetPath)
  }
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
        folderPath: 'Archive',
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
    await store.replaceMailboxDocuments('C:/PST/Case10/Search1/gamma.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case10/Search1/gamma.pst',
        scopePath: 'Case10/Search1',
        scopeLabel: 'Case10 / Search1',
        fileName: 'gamma.pst',
        mailboxName: 'Gamma',
        messageId: 'message:10',
        descriptorId: '10',
        folderId: 'folder:10',
        folderPath: 'Inbox',
        order: 3,
        subject: 'Case Ten',
        originalSubject: 'Case Ten',
        senderName: 'Zed Example',
        senderEmailAddress: 'zed@example.com',
        recipientText: 'Zed Example <zed@example.com>',
        displayTo: 'Zed Example <zed@example.com>',
        resolvedDisplayTo: 'Zed Example <zed@example.com>',
        bodySearchText: 'case ten',
        searchText: 'case ten zed example zed@example.com ipm.note mail',
        searchTokens: ['case', 'ten', 'zed', 'example'],
        addressValues: ['zed@example.com'],
        subjectValues: ['case ten'],
        sortDateMs: Date.parse('2024-01-03T00:00:00.000Z'),
        sortDate: '2024-01-03T00:00:00.000Z',
        updatedAt: new Date('2024-01-03T00:00:00.000Z').toISOString()
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

    const caseScoped = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(caseScoped.total).toBe(2)
    expect(caseScoped.items.every((item) => item.scopePath.startsWith('Case1/'))).toBe(true)
    expect(caseScoped.items.some((item) => item.scopePath.startsWith('Case10/'))).toBe(false)

    const subjectAsc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'subject-asc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(subjectAsc.items.map((item) => item.messageId)).toEqual(['message:2', 'message:1'])

    const subjectDesc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'subject-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(subjectDesc.items.map((item) => item.messageId)).toEqual(['message:1', 'message:2'])

    const senderAsc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'sender-asc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(senderAsc.items.map((item) => item.messageId)).toEqual(['message:1', 'message:2'])

    const senderDesc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'sender-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(senderDesc.items.map((item) => item.messageId)).toEqual(['message:2', 'message:1'])

    const locationAsc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'location-asc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(locationAsc.items.map((item) => item.messageId)).toEqual(['message:2', 'message:1'])

    const locationDesc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'location-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(locationDesc.items.map((item) => item.messageId)).toEqual(['message:1', 'message:2'])

    const dateAsc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'date-asc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(dateAsc.items.map((item) => item.messageId)).toEqual(['message:1', 'message:2'])

    const dateDesc = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(dateDesc.items.map((item) => item.messageId)).toEqual(['message:2', 'message:1'])

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

  it('collapses mailbox conversation threads to the newest matching item', async () => {
    const store = new MemorySearchIndexStore()

    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/alpha.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/alpha.pst',
        fileName: 'alpha.pst',
        mailboxName: 'Alpha',
        messageId: 'message:older',
        descriptorId: 'alpha-1',
        folderId: 'folder:alpha',
        folderPath: 'Inbox',
        subject: 'Shared contract',
        originalSubject: 'Shared contract',
        senderName: 'Alice Example',
        senderEmailAddress: 'alice@example.com',
        recipientText: 'Bob Example <bob@example.com>',
        displayTo: 'Bob Example <bob@example.com>',
        resolvedDisplayTo: 'Bob Example <bob@example.com>',
        bodySearchText: 'legacy-term contract body text',
        searchText: 'shared contract legacy-term alice example alice@example.com bob example bob@example.com contract body text ipm.note mail',
        searchTokens: ['shared', 'contract', 'legacy-term', 'alice', 'example', 'bob'],
        addressValues: ['alice@example.com', 'bob@example.com'],
        subjectValues: ['shared contract'],
        sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
        sortDate: '2024-01-01T00:00:00.000Z',
        updatedAt: new Date('2024-01-01T00:00:00.000Z').toISOString(),
        threadMetadata: {
          messageId: '<project-alpha-root@example.com>',
          inReplyToId: '',
          referenceIds: [],
          conversationId: '',
          isForward: false
        }
      })
    ])
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/beta.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/beta.pst',
        fileName: 'beta.pst',
        mailboxName: 'Beta',
        messageId: 'message:newer',
        descriptorId: 'beta-1',
        folderId: 'folder:beta',
        folderPath: 'Inbox',
        subject: 'Shared contract',
        originalSubject: 'Shared contract',
        senderName: 'Alice Example',
        senderEmailAddress: 'alice@example.com',
        recipientText: 'Bob Example <bob@example.com>',
        displayTo: 'Bob Example <bob@example.com>',
        resolvedDisplayTo: 'Bob Example <bob@example.com>',
        bodySearchText: 'contract body text',
        searchText: 'shared contract alice example alice@example.com bob example bob@example.com contract body text ipm.note mail',
        searchTokens: ['shared', 'contract', 'alice', 'example', 'bob'],
        addressValues: ['alice@example.com', 'bob@example.com'],
        subjectValues: ['shared contract'],
        sortDateMs: Date.parse('2024-01-03T00:00:00.000Z'),
        sortDate: '2024-01-03T00:00:00.000Z',
        updatedAt: new Date('2024-01-03T00:00:00.000Z').toISOString(),
        threadMetadata: {
          messageId: '<project-alpha-reply@example.com>',
          inReplyToId: '<project-alpha-root@example.com>',
          referenceIds: ['project-alpha-root@example.com'],
          conversationId: '',
          isForward: false
        }
      })
    ])

    const uncollapsed = await store.search({
      scope: 'all',
      query: 'legacy-term',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: ''
    })
    expect(uncollapsed.total).toBe(1)
    expect(uncollapsed.items[0].messageId).toBe('message:older')

    const collapsed = await store.search({
      scope: 'all',
      query: 'legacy-term',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(collapsed.total).toBe(1)
    expect(collapsed.totalPages).toBe(1)
    expect(collapsed.page).toBe(1)
    expect(collapsed.items).toHaveLength(1)
    expect(collapsed.items[0].messageId).toBe('message:newer')
    expect(collapsed.sourceCounts).toEqual({ mailbox: 1, teams: 0, sharepoint: 0 })
  })

  it('collapses appointment threads and leaves untitled items independent', async () => {
    const store = new MemorySearchIndexStore()
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/calendar.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/calendar.pst',
        messageId: 'appointment:older',
        kind: 'appointment',
        isMailLike: false,
        subject: 'Project Alpha meeting',
        searchText: 'project alpha meeting appointment',
        searchTokens: ['project', 'alpha', 'meeting', 'appointment'],
        sortDate: '2024-01-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<appointment-root@example.com>',
          inReplyToId: '',
          referenceIds: [],
          conversationId: 'appointment-conversation-a',
          isForward: false
        }
      }),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/calendar.pst',
        messageId: 'appointment:newer',
        kind: 'appointment',
        isMailLike: false,
        subject: 'Project Alpha meeting updated',
        searchText: 'project alpha meeting appointment',
        searchTokens: ['project', 'alpha', 'meeting', 'appointment'],
        sortDate: '2024-01-02T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<appointment-update@example.com>',
          inReplyToId: '',
          referenceIds: [],
          conversationId: 'appointment-conversation-a',
          isForward: false
        }
      }),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/calendar.pst',
        messageId: 'appointment:untitled-a',
        kind: 'appointment',
        isMailLike: false,
        subject: 'Untitled appointment',
        searchText: 'untitled appointment',
        searchTokens: ['untitled', 'appointment'],
        mailboxDetail: { conversationTopic: '' } as any
      }),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/calendar.pst',
        messageId: 'appointment:untitled-b',
        kind: 'appointment',
        isMailLike: false,
        subject: 'Untitled appointment',
        searchText: 'untitled appointment',
        searchTokens: ['untitled', 'appointment'],
        mailboxDetail: undefined
      })
    ])

    const collapsed = await store.search({
      scope: 'all',
      query: 'appointment',
      mode: 'and',
      mailOnly: false,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(collapsed.total).toBe(3)
    expect(collapsed.items.map((item) => item.messageId)).toEqual(
      expect.arrayContaining(['appointment:newer', 'appointment:untitled-a', 'appointment:untitled-b'])
    )
    expect(collapsed.items).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ messageId: 'appointment:older' })])
    )
  })

  it('does not collapse same-topic forwards without explicit relationship metadata', async () => {
    const store = new MemorySearchIndexStore()
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/forwarded.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/forwarded.pst',
        messageId: 'message:forward-1',
        subject: 'Indicative offer',
        searchText: 'indicative offer forwarded message',
        searchTokens: ['indicative', 'offer', 'forwarded', 'message'],
        sortDate: '2024-01-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
        mailboxDetail: { conversationTopic: 'Indicative offer' } as any
      }),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/forwarded.pst',
        messageId: 'message:forward-2',
        subject: 'Indicative offer',
        searchText: 'indicative offer forwarded message',
        searchTokens: ['indicative', 'offer', 'forwarded', 'message'],
        sortDate: '2024-01-02T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
        mailboxDetail: { conversationTopic: 'Indicative offer' } as any
      })
    ])

    const collapsed = await store.search({
      scope: 'all',
      query: 'indicative',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(collapsed.total).toBe(2)
    expect(collapsed.items.map((item) => item.messageId)).toEqual(
      expect.arrayContaining(['message:forward-1', 'message:forward-2'])
    )
  })

  it('keeps copied-reference forwards in separate branches and collapses replies to each forward', async () => {
    const store = new MemorySearchIndexStore()
    const mailboxKey = 'C:/PST/Case1/Search1/branches.pst'
    const makeThreadDocument = (overrides: Partial<SearchIndexDocument>): SearchIndexDocument =>
      makeDocument({
        mailboxKey,
        fileName: 'branches.pst',
        mailboxName: 'Branches',
        subject: 'Indicative offer',
        originalSubject: 'Indicative offer',
        searchText: 'indicative offer',
        searchTokens: ['indicative', 'offer'],
        ...overrides
      })

    await store.replaceMailboxDocuments(mailboxKey, [
      makeThreadDocument({
        messageId: 'message:root',
        sortDate: '2024-01-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<root@example.com>',
          inReplyToId: '',
          referenceIds: [],
          conversationId: '',
          isForward: false
        }
      }),
      makeThreadDocument({
        messageId: 'message:original-reply',
        sortDate: '2024-01-02T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<original-reply@example.com>',
          inReplyToId: '<root@example.com>',
          referenceIds: ['root@example.com'],
          conversationId: '',
          isForward: false
        }
      }),
      makeThreadDocument({
        messageId: 'message:march-forward',
        subject: 'Fw: Indicative offer',
        originalSubject: 'Fw: Indicative offer',
        sortDate: '2024-03-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-03-01T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<march-forward@example.com>',
          inReplyToId: '<root@example.com>',
          referenceIds: ['root@example.com'],
          conversationId: '',
          isForward: true
        }
      }),
      makeThreadDocument({
        messageId: 'message:march-reply',
        subject: 'Re: Fw: Indicative offer',
        originalSubject: 'Re: Fw: Indicative offer',
        sortDate: '2024-03-02T00:00:00.000Z',
        sortDateMs: Date.parse('2024-03-02T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<march-reply@example.com>',
          inReplyToId: '<march-forward@example.com>',
          referenceIds: ['root@example.com', 'march-forward@example.com'],
          conversationId: '',
          isForward: false
        }
      }),
      makeThreadDocument({
        messageId: 'message:december-forward',
        subject: 'Fwd: Indicative offer',
        originalSubject: 'Fwd: Indicative offer',
        sortDate: '2024-12-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-12-01T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<december-forward@example.com>',
          inReplyToId: '<root@example.com>',
          referenceIds: ['root@example.com'],
          conversationId: '',
          isForward: true
        }
      })
    ])

    const collapsed = await store.search({
      scope: 'all',
      query: 'indicative',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(collapsed.total).toBe(3)
    expect(collapsed.items.map((item) => item.messageId)).toEqual([
      'message:december-forward',
      'message:march-reply',
      'message:original-reply'
    ])
  })

  it('retains the newest representative for each reply branch', async () => {
    const store = new MemorySearchIndexStore()
    const mailboxKey = 'C:/PST/Case1/Search1/branching.pst'
    const makeBranchDocument = (overrides: Partial<SearchIndexDocument>): SearchIndexDocument =>
      makeDocument({
        mailboxKey,
        fileName: 'branching.pst',
        subject: 'Branching thread',
        searchText: 'branching thread root',
        searchTokens: ['branching', 'thread', 'root'],
        ...overrides
      })

    await store.replaceMailboxDocuments(mailboxKey, [
      makeBranchDocument({
        messageId: 'message:branch-root',
        sortDate: '2024-01-01T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<branch-root@example.com>',
          inReplyToId: '',
          referenceIds: [],
          conversationId: '',
          isForward: false
        }
      }),
      makeBranchDocument({
        messageId: 'message:branch-a',
        searchText: 'branching branch-a',
        searchTokens: ['branching', 'branch-a'],
        sortDate: '2024-01-02T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<branch-a@example.com>',
          inReplyToId: '<branch-root@example.com>',
          referenceIds: ['branch-root@example.com'],
          conversationId: '',
          isForward: false
        }
      }),
      makeBranchDocument({
        messageId: 'message:branch-a-latest',
        searchText: 'branching branch-a latest',
        searchTokens: ['branching', 'branch-a', 'latest'],
        sortDate: '2024-01-03T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-03T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<branch-a-latest@example.com>',
          inReplyToId: '<branch-a@example.com>',
          referenceIds: ['branch-root@example.com', 'branch-a@example.com'],
          conversationId: '',
          isForward: false
        }
      }),
      makeBranchDocument({
        messageId: 'message:branch-b',
        searchText: 'branching branch-b',
        searchTokens: ['branching', 'branch-b'],
        sortDate: '2024-01-04T00:00:00.000Z',
        sortDateMs: Date.parse('2024-01-04T00:00:00.000Z'),
        threadMetadata: {
          messageId: '<branch-b@example.com>',
          inReplyToId: '<branch-root@example.com>',
          referenceIds: ['branch-root@example.com'],
          conversationId: '',
          isForward: false
        }
      })
    ])

    const collapsed = await store.search({
      scope: 'all',
      query: 'branching',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(collapsed.items.map((item) => item.messageId)).toEqual([
      'message:branch-b',
      'message:branch-a-latest'
    ])
    expect(collapsed.items.every((item) => item.threadInfo?.branchCount === 2)).toBe(true)
    expect(collapsed.items.find((item) => item.messageId === 'message:branch-a-latest')?.threadInfo).toMatchObject({
      branchItemCount: 3,
      threadItemCount: 4,
      isRepresentative: true
    })

    const oneBranch = await store.search({
      scope: 'all',
      query: 'branch-a',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })
    expect(oneBranch.items.map((item) => item.messageId)).toEqual(['message:branch-a-latest'])
  })

  it('parses folded reply and references headers into relationship metadata', () => {
    expect(
      buildMailboxThreadMetadata(
        {
          internetMessageId: '<reply@example.com>',
          inReplyToId: '',
          conversationId: '',
          transportMessageHeaders:
            'In-Reply-To: <root@example.com>\r\nReferences: <root@example.com>\r\n  <prior@example.com>'
        },
        'mail'
      )
    ).toEqual({
      messageId: 'reply@example.com',
      inReplyToId: 'root@example.com',
      referenceIds: ['root@example.com', 'prior@example.com'],
      conversationId: '',
      isForward: false
    })
    expect(
      buildMailboxThreadMetadata(
        {
          internetMessageId: 'not-a-message-id',
          inReplyToId: 'also-invalid',
          conversationId: '',
          transportMessageHeaders: 'References: malformed-value'
        },
        'mail'
      )
    ).toBeUndefined()
  })

  it('classifies forwards from subject prefixes and unambiguous body markers only', () => {
    const baseDetail = {
      internetMessageId: '<message@example.com>',
      inReplyToId: '',
      conversationId: '',
      transportMessageHeaders: ''
    }

    expect(
      buildMailboxThreadMetadata({ ...baseDetail, subject: 'Fw: Indicative offer' }, 'mail')?.isForward
    ).toBe(true)
    expect(
      buildMailboxThreadMetadata({ ...baseDetail, subject: 'Re: Fwd: Indicative offer' }, 'mail')?.isForward
    ).toBe(true)
    expect(
      buildMailboxThreadMetadata({ ...baseDetail, bodyText: '-----Original Message-----\nFrom: Alice' }, 'mail')?.isForward
    ).toBe(true)
    expect(
      buildMailboxThreadMetadata({ ...baseDetail, bodyText: 'Begin forwarded message:\nFrom: Alice' }, 'mail')?.isForward
    ).toBe(true)
    expect(
      buildMailboxThreadMetadata({ ...baseDetail, subject: 'Re: Indicative offer', bodyText: 'From: Alice\nSent: today' }, 'mail')?.isForward
    ).toBe(false)
    expect(
      buildMailboxThreadMetadata({ ...baseDetail, subject: 'Fw: Indicative offer' }, 'appointment')?.isForward
    ).toBe(false)
  })

  it('includes item sizes and keeps flagged-size totals stable across paging', async () => {
    const store = new MemorySearchIndexStore()
    const flaggedAt = new Date('2024-01-03T00:00:00.000Z').toISOString()
    const unflaggedAt = new Date('2024-01-02T00:00:00.000Z').toISOString()
    const olderAt = new Date('2024-01-01T00:00:00.000Z').toISOString()

    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/mailbox.pst', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/mailbox.pst',
        fileName: 'mailbox.pst',
        mailboxName: 'Mailbox',
        messageId: 'message:mail-1',
        descriptorId: 'mail-1',
        folderId: 'folder:mail-1',
        folderPath: 'Inbox',
        sourceType: 'mailbox',
        size: 1024 * 1024,
        sortDateMs: Date.parse(flaggedAt),
        sortDate: flaggedAt,
        reviewStates: [
          {
            reviewerUsername: 'admin',
            review: {
              flagged: true,
              tags: [],
              createdAt: flaggedAt,
              updatedAt: flaggedAt
            }
          }
        ]
      })
    ])
    await store.replaceMailboxDocuments('C:/PST/Case1/Search1/items.zip', [
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/items.zip',
        fileName: 'items.zip',
        mailboxName: 'Teams bundle',
        messageId: 'message:teams-1',
        descriptorId: 'teams-1',
        folderId: 'folder:teams-1',
        folderPath: 'Teams',
        sourceType: 'teams',
        kind: 'other',
        messageClass: 'IPM.Note',
        isMailLike: false,
        size: 2 * 1024 * 1024,
        senderName: 'Teams',
        senderEmailAddress: '',
        recipientText: '',
        displayTo: '',
        displayCC: '',
        displayBCC: '',
        resolvedDisplayTo: '',
        resolvedDisplayCC: '',
        resolvedDisplayBCC: '',
        subject: 'Teams update',
        originalSubject: 'Teams update',
        searchText: 'teams update',
        searchTokens: ['teams', 'update'],
        addressValues: [],
        subjectValues: ['teams update'],
        sortDateMs: Date.parse(unflaggedAt),
        sortDate: unflaggedAt,
        reviewStates: [
          {
            reviewerUsername: 'admin',
            review: {
              flagged: false,
              tags: [],
              createdAt: unflaggedAt,
              updatedAt: unflaggedAt
            }
          }
        ]
      }),
      makeDocument({
        mailboxKey: 'C:/PST/Case1/Search1/items.zip',
        fileName: 'items.zip',
        mailboxName: 'SharePoint bundle',
        messageId: 'message:sharepoint-1',
        descriptorId: 'sharepoint-1',
        folderId: 'folder:sharepoint-1',
        folderPath: 'SharePoint',
        sourceType: 'sharepoint',
        kind: 'other',
        messageClass: 'IPM.Document',
        isMailLike: false,
        size: 3 * 1024 * 1024,
        senderName: 'SharePoint',
        senderEmailAddress: '',
        recipientText: '',
        displayTo: '',
        displayCC: '',
        displayBCC: '',
        resolvedDisplayTo: '',
        resolvedDisplayCC: '',
        resolvedDisplayBCC: '',
        subject: 'SharePoint file',
        originalSubject: 'SharePoint file',
        searchText: 'sharepoint file',
        searchTokens: ['sharepoint', 'file'],
        addressValues: [],
        subjectValues: ['sharepoint file'],
        sortDateMs: Date.parse(olderAt),
        sortDate: olderAt,
        reviewStates: [
          {
            reviewerUsername: 'admin',
            review: {
              flagged: true,
              tags: [],
              createdAt: olderAt,
              updatedAt: olderAt
            }
          }
        ]
      })
    ])

    const allResults = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: false,
      sort: 'date-desc',
      page: 1,
      pageSize: 20,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      reviewerUsername: 'admin'
    })

    expect(allResults.total).toBe(3)
    expect(allResults.items.find((item) => item.sourceType === 'mailbox')?.size).toBe(1024 * 1024)
    expect(allResults.items.find((item) => item.sourceType === 'teams')?.size).toBe(2 * 1024 * 1024)
    expect(allResults.items.find((item) => item.sourceType === 'sharepoint')?.size).toBe(3 * 1024 * 1024)
    expect(allResults.flaggedSizeBytes).toBe(4 * 1024 * 1024)

    const firstPage = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: false,
      sort: 'date-desc',
      page: 1,
      pageSize: 1,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      reviewerUsername: 'admin'
    })
    const secondPage = await store.search({
      scope: 'all',
      casePath: 'Case1',
      query: '',
      mode: 'and',
      mailOnly: false,
      sort: 'date-desc',
      page: 2,
      pageSize: 1,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      reviewerUsername: 'admin'
    })

    expect(firstPage.flaggedSizeBytes).toBe(4 * 1024 * 1024)
    expect(secondPage.flaggedSizeBytes).toBe(4 * 1024 * 1024)
    expect(firstPage.items).toHaveLength(1)
    expect(secondPage.items).toHaveLength(1)
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

  it('reads archive content even when the stored chain uses a prefixed or differently cased path', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-archive-chain-fallback-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const bundlePath = path.join(scopeDir, 'Items.1.001.CHAIN.zip')
      const bundleZip = new AdmZip()
      bundleZip.addFile('SharePoint/Docs/report.txt', Buffer.from('sharepoint report', 'utf8'))
      bundleZip.writeZip(bundlePath)

      const content = await readArchiveBundleItemContent(bundlePath, ['./sharepoint/Docs/report.txt'])
      expect(content.fileName).toBe('report.txt')
      expect(content.buffer.toString('utf8')).toBe('sharepoint report')
      expect(content.contentType).toBe('text/plain; charset=utf-8')
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

  it('stores mailbox preview payloads on search documents during refresh', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-mailbox-snapshot-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const mailboxPath = path.join(scopeDir, 'sample.pst')
      stageFixture(enronPath, mailboxPath)

      const reviewStore = {
        kind: 'memory' as const,
        isPersistent: false,
        async listReviews() {
          return []
        }
      } as ReviewStore

      class TrackingSearchIndexStore extends MemorySearchIndexStore {
        mailboxDeleteCount = 0

        override async deleteMailboxDocuments(mailboxKey: string): Promise<void> {
          this.mailboxDeleteCount += 1
          await super.deleteMailboxDocuments(mailboxKey)
        }
      }

      const store = new TrackingSearchIndexStore()
      const firstPlan = await refreshSearchIndexSourceFromCatalog(
        rootDir,
        'mailboxes',
        reviewStore,
        store
      )
      expect(firstPlan.mailboxCount).toBeGreaterThan(0)

      const initialSearch = await store.search({
        scope: 'all',
        query: '',
        mode: 'and',
        mailOnly: true,
        sort: 'date-desc',
        page: 1,
        pageSize: 20,
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      })
      const mailboxItem = initialSearch.items.find((item) => item.sourceType === 'mailbox')
      expect(mailboxItem).toBeTruthy()
      if (!mailboxItem) {
        throw new Error('Expected a mailbox search result')
      }

      expect(mailboxItem.mailboxDetail).toBeTruthy()
      expect(mailboxItem.mailboxDetail?.id).toBe(mailboxItem.messageId)
      expect(
        (mailboxItem.mailboxDetail?.bodyText || mailboxItem.mailboxDetail?.bodyHtml || '').length
      ).toBeGreaterThan(0)
      expect(Array.isArray(mailboxItem.mailboxDetail?.attachments)).toBe(true)

      const refreshedAt = new Date(Date.now() + 60_000)
      fs.utimesSync(mailboxPath, refreshedAt, refreshedAt)

      const secondPlan = await refreshSearchIndexSourceFromCatalog(
        rootDir,
        'mailboxes',
        reviewStore,
        store
      )
      expect(secondPlan.changedCount).toBeGreaterThan(0)
      expect(await store.findDocumentById(mailboxItem.id || mailboxItem.messageId)).toBeTruthy()

      fs.unlinkSync(mailboxPath)
      const thirdPlan = await refreshSearchIndexSourceFromCatalog(
        rootDir,
        'mailboxes',
        reviewStore,
        store
      )
      expect(thirdPlan.removedCount).toBeGreaterThan(0)
      expect(store.mailboxDeleteCount).toBeGreaterThan(0)
      expect(await store.findDocumentById(mailboxItem.id || mailboxItem.messageId)).toBeNull()
    } finally {
      fs.rmSync(rootDir, { recursive: true, force: true })
    }
  })

  it('persists each refreshed file before the next file finishes', async () => {
    const rootDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-refresh-live-commit-'))
    try {
      const scopeDir = path.join(rootDir, 'Case1', 'Search1')
      fs.mkdirSync(scopeDir, { recursive: true })

      const firstFilePath = path.join(scopeDir, 'first.pst')
      const secondFilePath = path.join(scopeDir, 'second.pst')
      stageFixture(enronPath, firstFilePath)
      stageFixture(enronPath, secondFilePath)

      const reviewStore = {
        kind: 'memory' as const,
        isPersistent: false,
        async listReviews() {
          return []
        }
      } as ReviewStore

      const secondReplaceStarted = createDeferred<void>()
      const releaseSecondReplace = createDeferred<void>()

      class BlockingSearchIndexStore extends MemorySearchIndexStore {
        private replaceCount = 0

        override async replaceMailboxDocuments(
          mailboxKey: string,
          documents: Parameters<MemorySearchIndexStore['replaceMailboxDocuments']>[1]
        ): Promise<void> {
          this.replaceCount += 1
          if (this.replaceCount === 2) {
            secondReplaceStarted.resolve()
            await releaseSecondReplace.promise
            throw new Error('synthetic second-file failure')
          }
          return super.replaceMailboxDocuments(mailboxKey, documents)
        }
      }

      const store = new BlockingSearchIndexStore()
      const refreshPromise = refreshSearchIndexSourceFromCatalog(
        rootDir,
        'mailboxes',
        reviewStore,
        store
      )

      await secondReplaceStarted.promise

      const fingerprintsDuringRefresh = await store.listFileFingerprints('mailboxes')
      expect(fingerprintsDuringRefresh).toHaveLength(1)
      expect(fingerprintsDuringRefresh[0].fileName).toBe('first.pst')

      const interimSearch = await store.search({
        scope: 'all',
        query: '',
        mode: 'and',
        mailOnly: true,
        sort: 'date-desc',
        page: 1,
        pageSize: 50,
        reviewFlaggedOnly: false,
        reviewTaggedOnly: false,
        reviewTag: ''
      })
      expect(interimSearch.total).toBeGreaterThan(0)
      expect(interimSearch.items.some((item) => item.fileName === 'first.pst')).toBe(true)
      expect(interimSearch.items.some((item) => item.fileName === 'second.pst')).toBe(false)

      releaseSecondReplace.resolve()
      const plan = await refreshPromise

      expect(plan.changedCount).toBe(1)
      expect(plan.failedCount).toBe(1)
      expect(plan.mailboxCount).toBe(1)
      expect(plan.fingerprints).toHaveLength(1)
      expect(plan.fingerprints[0].fileName).toBe('first.pst')
      expect(await store.findDocumentById(interimSearch.items[0].id || interimSearch.items[0].messageId)).toBeTruthy()
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

  it('collapses Mongo mailbox threads after matching and clamps the representative page', async () => {
    const older = makeDocument({
      messageId: 'mongo:older',
      searchText: 'legacy-term thread',
      searchTokens: ['legacy-term', 'thread'],
      sortDate: '2024-01-01T00:00:00.000Z',
      sortDateMs: Date.parse('2024-01-01T00:00:00.000Z'),
      threadMetadata: {
        messageId: '<mongo-root@example.com>',
        inReplyToId: '',
        referenceIds: [],
        conversationId: '',
        isForward: false
      }
    })
    const newer = makeDocument({
      messageId: 'mongo:newer',
      searchText: 'thread',
      searchTokens: ['thread'],
      sortDate: '2024-01-02T00:00:00.000Z',
      sortDateMs: Date.parse('2024-01-02T00:00:00.000Z'),
      threadMetadata: {
        messageId: '<mongo-reply@example.com>',
        inReplyToId: '<mongo-root@example.com>',
        referenceIds: ['mongo-root@example.com'],
        conversationId: '',
        isForward: false
      }
    })
    const find = jest.fn(() => {
      let skipCount = 0
      let limitCount = 0
      const toArray = jest.fn(async () => {
        if (limitCount === 0) {
          return [older, newer]
        }
        return [older, newer].slice(skipCount, skipCount + limitCount)
      })
      const limit = jest.fn((count: number) => {
        limitCount = count
        return { toArray }
      })
      const skip = jest.fn((count: number) => {
        skipCount = count
        return { limit }
      })
      const sort = jest.fn(() => ({ skip }))
      return { sort }
    })
    const documentsCollection = {
      find,
      findOne: jest.fn().mockResolvedValue(null),
      aggregate: jest.fn(),
      countDocuments: jest.fn(),
      updateOne: jest.fn().mockResolvedValue(undefined),
      insertMany: jest.fn().mockResolvedValue(undefined),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any
    const rulesCollection = {
      find: jest.fn(() => ({
        sort: jest.fn(() => ({ toArray: jest.fn().mockResolvedValue([]) }))
      })),
      findOne: jest.fn().mockResolvedValue(null),
      updateOne: jest.fn().mockResolvedValue(undefined),
      deleteOne: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any
    const fingerprintsCollection = {
      find: jest.fn(() => ({
        sort: jest.fn(() => ({ toArray: jest.fn().mockResolvedValue([]) }))
      })),
      findOne: jest.fn().mockResolvedValue(null),
      updateOne: jest.fn().mockResolvedValue(undefined),
      deleteMany: jest.fn().mockResolvedValue({ deletedCount: 0 }),
      createIndex: jest.fn().mockResolvedValue(undefined)
    } as any

    const store = new MongoSearchIndexStore(documentsCollection, rulesCollection, fingerprintsCollection)
    const results = await store.search({
      scope: 'all',
      sourceType: 'mailbox',
      query: 'legacy-term',
      mode: 'and',
      mailOnly: true,
      sort: 'date-desc',
      page: 9,
      pageSize: 1,
      reviewFlaggedOnly: false,
      reviewTaggedOnly: false,
      reviewTag: '',
      collapseDuplicates: true
    })

    expect(find).toHaveBeenCalledTimes(2)
    expect(results.total).toBe(1)
    expect(results.page).toBe(1)
    expect(results.totalPages).toBe(1)
    expect(results.items[0].messageId).toBe('mongo:newer')
    expect(results.sourceCounts).toEqual({ mailbox: 1, teams: 0, sharepoint: 0 })
  })
})
