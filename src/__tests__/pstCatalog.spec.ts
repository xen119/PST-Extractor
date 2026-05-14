import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import {
  listPstMailboxFiles,
  openPstMailbox,
  resolvePstMailboxPath
} from '../pstCatalog'

const resolve = path.resolve

const enronPath = resolve('./src/__tests__/testdata/enron.pst')
const outlookPath = resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')

function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix))
}

describe('pst catalog helpers', () => {
  it('lists only PST and OST files from a mailbox folder', () => {
    const rootDir = makeTempDir('pst-catalog-list-')
    fs.copyFileSync(enronPath, path.join(rootDir, 'enron.pst'))
    fs.copyFileSync(outlookPath, path.join(rootDir, 'mail.ost'))
    fs.writeFileSync(path.join(rootDir, 'notes.txt'), 'ignore me')
    fs.mkdirSync(path.join(rootDir, 'nested'))

    const result = listPstMailboxFiles(rootDir)

    expect(result.rootExists).toBe(true)
    expect(result.files.map((file) => file.fileName)).toEqual(['enron.pst', 'mail.ost'])
    expect(result.files[0].size).toBeGreaterThan(0)
    expect(result.files[0].modifiedAt).toMatch(/T/)
    expect(result.message).toContain('2 mailbox files')
  })

  it('reports an empty or missing PST folder without throwing', () => {
    const emptyRoot = makeTempDir('pst-catalog-empty-')
    const emptyResult = listPstMailboxFiles(emptyRoot)
    expect(emptyResult.rootExists).toBe(true)
    expect(emptyResult.files).toHaveLength(0)
    expect(emptyResult.message).toContain('No PST or OST files')

    const missingRoot = path.join(makeTempDir('pst-catalog-missing-'), 'PST')
    const missingResult = listPstMailboxFiles(missingRoot)
    expect(missingResult.rootExists).toBe(false)
    expect(missingResult.files).toHaveLength(0)
    expect(missingResult.message).toContain('Create a PST folder')
  })

  it('rejects traversal attempts and opens mailbox files within the folder', () => {
    const rootDir = makeTempDir('pst-catalog-open-')
    fs.copyFileSync(enronPath, path.join(rootDir, 'enron.pst'))

    expect(() => resolvePstMailboxPath(rootDir, '../enron.pst')).toThrow(
      'Mailbox file name must not include a path'
    )
    expect(() => resolvePstMailboxPath(rootDir, 'nested/enron.pst')).toThrow(
      'Mailbox file name must not include a path'
    )
    expect(() => resolvePstMailboxPath(rootDir, 'notes.txt')).toThrow(
      'Only .pst and .ost files are supported'
    )

    const session = openPstMailbox(rootDir, 'enron.pst')
    expect(session.fileName).toBe('enron.pst')
    expect(session.mailboxName).toBe('Personal folders')
    expect(session.stats.folderCount).toBe(11)
    expect(session.stats.messageCount).toBe(71)
  })
})
