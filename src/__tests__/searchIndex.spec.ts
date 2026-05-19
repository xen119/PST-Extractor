import { MemorySearchIndexStore, type SearchIndexDocument } from '../searchIndex'

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
    reviewTagValues: [],
    updatedAt: now,
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
})
