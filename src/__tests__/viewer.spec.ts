import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import {
  buildEmptyMessageDetail,
  buildFolderTree,
  buildMessageDetail,
  buildMessageDetailFromSession,
  createViewerSession,
  htmlToText,
  exportMessageAsEml,
  exportMessageAsEmlFromSession,
  exportMessageAsJson,
  getAttachmentDownloadBuffer,
  getMessageDetail,
  listSessionMessages,
  messageMatchesQuery,
  listFolderMessages,
  sanitizeFileNameForDownload,
  withSessionMessage,
  type MessageDetail,
  type MessageSummary
} from '../viewer'

const resolve = path.resolve

const enronPath = resolve('./src/__tests__/testdata/enron.pst')
const outlookPath = resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')

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

function walkFolderTree(
  node: ReturnType<typeof buildFolderTree>,
  callback: (folderId: string) => void
): void {
  callback(node.id)
  node.messageIds.forEach((messageId) => callback(messageId))
  node.children.forEach((child) => walkFolderTree(child, callback))
}

function makeSummary(overrides: Partial<MessageSummary> = {}): MessageSummary {
  const now = new Date('2024-01-01T00:00:00.000Z').toISOString()
  return {
    id: 'message:summary',
    descriptorId: '1234',
    folderId: 'folder:1',
    folderPath: 'Mailbox/Inbox',
    order: 0,
    messageClass: 'IPM.Note',
    kind: 'mail',
    subject: 'Summary subject',
    senderName: 'Sender',
    senderEmailAddress: 'sender@example.com',
    recipientText: 'Recipient',
    displayTo: 'Recipient',
    displayCC: '',
    displayBCC: '',
    resolvedDisplayTo: 'Recipient <recipient@example.com>',
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
    ...overrides
  }
}

function makeDetail(overrides: Partial<MessageDetail> = {}): MessageDetail {
  return {
    ...buildEmptyMessageDetail(makeSummary()),
    ...overrides
  }
}

describe('viewer integration', () => {
  let enronSession = createViewerSession(enronPath, 'enron.pst')
  let outlookSession = createViewerSession(outlookPath, 'mtnman1965@outlook.com.ost')

  afterEach(() => {
    jest.restoreAllMocks()
  })

  it('indexes the bundled PST and keeps malformed folders as warnings', () => {
    expect(enronSession.stats.folderCount).toBe(11)
    expect(enronSession.stats.messageCount).toBe(71)
    expect(enronSession.stats.warningCount).toBeGreaterThan(0)
    expect(enronSession.warnings[0]).toContain('Unable to load subfolders')

    const tree = buildFolderTree(enronSession)
    expect(tree.displayName).toBe('Personal folders')
    expect(tree.children.map((child) => child.displayName)).toEqual([
      'Top of Personal Folders',
      'Search Root',
      'SPAM Search Folder 2'
    ])

    const seen = new Set<string>()
    walkFolderTree(tree, (itemId) => {
      expect(seen.has(itemId)).toBe(false)
      seen.add(itemId)
      expect(itemId.startsWith('folder:') || itemId.startsWith('message:')).toBe(true)
    })
    expect(seen.size).toBeGreaterThan(0)
  })

  it('filters and pages folder messages by metadata', () => {
    const targetFolder = [...enronSession.folders.values()].find(
      (folder) => folder.displayName === 'TW-Commercial Group'
    )
    expect(targetFolder).toBeTruthy()

    const filtered = listFolderMessages(enronSession, targetFolder!.id, {
      query: 'Lindberg',
      pageSize: 10
    })
    expect(filtered.total).toBeGreaterThan(0)
    expect(filtered.items.some((item) => item.subject === "New OBA's")).toBe(true)

    const ordered = listFolderMessages(enronSession, targetFolder!.id, {
      pageSize: 5,
      sort: 'order'
    })
    expect(ordered.pageSize).toBe(5)
    expect(ordered.items[0].order).toBeLessThan(ordered.items[1].order)
  })

  it('matches metadata, plain body text, and normalized html body text', () => {
    const summary = makeSummary({
      subject: 'Alpha Notice',
      senderName: 'Sender One'
    })

    expect(messageMatchesQuery(summary, 'alpha')).toBe(true)
    expect(
      messageMatchesQuery(
        summary,
        'signature',
        "The attached OBA's have been submitted to the customers for their signature."
      )
    ).toBe(true)

    const htmlBody = htmlToText('<div>Project <strong>update</strong></div>')
    expect(htmlBody).toBe('Project update')
    expect(messageMatchesQuery(summary, 'update', htmlBody)).toBe(true)
  })

  it('searches body text across the whole session', () => {
    const matches = listSessionMessages(enronSession, {
      query: 'signature',
      mailOnly: true
    })

    expect(matches.some((item) => item.id === 'message:2097188')).toBe(true)
  })

  it('resolves recipient names to email addresses when loading details', () => {
    const detail = buildMessageDetailFromSession(outlookSession, 'message:2110308')

    expect(detail.resolvedDisplayTo || detail.displayTo || '').toContain('@')
    expect(detail.resolvedDisplayCC || detail.displayCC || '').toEqual(expect.any(String))
  })

  it('finds messages by body keyword within a real folder', () => {
    const targetFolder = [...enronSession.folders.values()].find(
      (folder) => folder.displayName === 'TW-Commercial Group'
    )
    expect(targetFolder).toBeTruthy()

    const filtered = listFolderMessages(enronSession, targetFolder!.id, {
      query: 'signature',
      pageSize: 20
    })

    expect(filtered.total).toBeGreaterThan(0)
    expect(filtered.items.some((item) => item.id === 'message:2097188')).toBe(true)
    expect(filtered.items.find((item) => item.id === 'message:2097188')?.subject).toBe(
      "New OBA's"
    )
  })

  it('builds attachment download URLs and JSON exports', () => {
    const detail = withSessionMessage(outlookSession, 'message:2110308', (message, summary) =>
      buildMessageDetail(message, summary, {
        messageId: summary.id,
        attachmentBaseUrl: `/api/sessions/demo/messages/${encodeURIComponent(
          summary.id
        )}/attachments/`
      })
    )

    expect(detail.attachments).toHaveLength(1)
    expect(detail.attachments[0].downloadUrl).toContain('/attachments/0')

    const exported = JSON.parse(exportMessageAsJson(detail))
    expect(exported.attachments[0].downloadUrl).toContain('/attachments/0')
  })

  it('exports raw attachment bytes and EML for real messages', () => {
    const attachment = getAttachmentDownloadBuffer(
      outlookSession,
      'message:2110308',
      0
    )
    expect(attachment.filename).toBe('OBA_2760.doc')
    expect(attachment.contentType).toBe('application/msword')
    expect(attachment.buffer.length).toBeGreaterThan(1000)

    const eml = exportMessageAsEmlFromSession(outlookSession, 'message:2110308')
    expect(eml).toContain('MIME-Version: 1.0')
    expect(eml).toContain('Subject: word attachment')
    expect(eml).toContain('multipart/mixed')
    expect(eml).toContain('OBA_2760.doc')
  })

  it('renders nested embedded messages in best-effort EML output', () => {
    const nested = makeDetail({
      id: 'message:nested',
      subject: 'Nested child',
      senderName: 'Nested Sender',
      attachments: []
    })
    const parent = makeDetail({
      id: 'message:parent',
      subject: 'Parent message',
      hasAttachments: true,
      attachments: [
        {
          attachmentId: 'message:parent:attachment:0',
          index: 0,
          filename: 'nested.eml',
          longFilename: 'nested.eml',
          downloadFilename: 'nested.eml',
          mimeTag: 'message/rfc822',
          size: 0,
          attachMethod: 1,
          contentId: '',
          pathname: '',
          longPathname: '',
          isEmbeddedMessage: true,
          embeddedMessage: nested,
          downloadUrl: ''
        }
      ]
    })

    const eml = exportMessageAsEml(parent)
    expect(eml).toContain('Subject: Parent message')
    expect(eml).toContain('Subject: Nested child')
    expect(eml).toContain('message/rfc822')
  })

  it('sanitizes file names for browser downloads', () => {
    expect(sanitizeFileNameForDownload(' bad<>name?.pst ', 'fallback.pst')).toBe(
      'bad_name_.pst'
    )
  })

  it('handles invalid headers and encrypted PST files', () => {
    const invalidDir = makeTempDir('pst-explorer-invalid-')
    const invalidPath = path.join(invalidDir, 'invalid.pst')
    fs.writeFileSync(invalidPath, Buffer.from('not a pst file', 'utf8'))

    expect(() => createViewerSession(invalidPath, 'invalid.pst')).toThrow(
      /Invalid file header/
    )

    const encryptedDir = makeTempDir('pst-explorer-encrypted-')
    const encryptedPath = path.join(encryptedDir, 'encrypted.pst')
    const source = Buffer.alloc(514, 0)
    source[0] = '!'.charCodeAt(0)
    source[1] = 'B'.charCodeAt(0)
    source[2] = 'D'.charCodeAt(0)
    source[3] = 'N'.charCodeAt(0)
    source[10] = 23
    source[513] = 0x02
    fs.writeFileSync(encryptedPath, source)

    expect(() => createViewerSession(encryptedPath, 'encrypted.pst')).toThrow(
      /encrypted/
    )

    fs.rmSync(invalidDir, { recursive: true, force: true })
    fs.rmSync(encryptedDir, { recursive: true, force: true })
  })

  it('survives per-message and per-attachment failures without killing the session', () => {
    const { PSTUtil } = require('../PSTUtil.class')
    const loadSpy = jest.spyOn(PSTUtil, 'detectAndLoadPSTObject').mockImplementation(() => {
      throw new Error('forced message load failure')
    })

    const failedDetail = buildMessageDetailFromSession(
      outlookSession,
      'message:2110308'
    )
    expect(failedDetail.parseError).toContain('forced message load failure')
    expect(failedDetail.attachments).toHaveLength(0)
    loadSpy.mockRestore()

    const { PSTMessage } = require('../PSTMessage.class')
    const originalGetAttachment = PSTMessage.prototype.getAttachment
    const attachmentSpy = jest
      .spyOn(PSTMessage.prototype, 'getAttachment')
      .mockImplementation(function (this: PSTMessage, index: number) {
        if (this.subject === 'word attachment' && index === 0) {
          throw new Error('forced attachment failure')
        }
        return originalGetAttachment.call(this, index)
      })

    const attachmentDetail = getMessageDetail(outlookSession, 'message:2110308')
    expect(attachmentDetail.attachments[0].parseError).toContain(
      'forced attachment failure'
    )
    attachmentSpy.mockRestore()
  })
})
