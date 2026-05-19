import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import {
  listPstMailboxFiles,
  listRemovedPstMailboxFiles,
  movePstMailboxToRemoved,
  openPstMailbox,
  resolvePstMailboxPath,
  restorePstMailboxFromRemoved
} from '../pstCatalog'

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

describe('pst catalog helpers', () => {
  it('discovers recursive case/search scopes and filters empty folders', () => {
    const rootDir = makeTempDir('pst-catalog-list-')
    fs.mkdirSync(path.join(rootDir, 'Case1', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(rootDir, 'Case1', 'Search2'), { recursive: true })
    fs.mkdirSync(path.join(rootDir, 'Case2', 'Search1'), { recursive: true })
    fs.mkdirSync(path.join(rootDir, '_removed', 'Case9', 'Search9'), { recursive: true })
    stageFixture(enronPath, path.join(rootDir, 'Case1', 'Search1', 'enron.pst'))
    stageFixture(outlookPath, path.join(rootDir, 'Case2', 'Search1', 'mail.ost'))
    stageFixture(enronPath, path.join(rootDir, '_removed', 'Case9', 'Search9', 'removed.pst'))
    fs.writeFileSync(path.join(rootDir, 'Case1', 'Search2', 'notes.txt'), 'ignore me')

    const result = listPstMailboxFiles(rootDir)

    expect(result.rootExists).toBe(true)
    expect(result.scopes.map((scope) => scope.scopeLabel)).toEqual([
      'Case1 / Search1',
      'Case2 / Search1'
    ])
    expect(result.scopePath).toBe('Case1/Search1')
    expect(result.scopeLabel).toBe('Case1 / Search1')
    expect(result.files.map((file) => file.fileName)).toEqual(['enron.pst'])
    expect(result.scopes[0].files.map((file) => file.fileName)).toEqual(['enron.pst'])
    expect(result.scopes[1].files.map((file) => file.fileName)).toEqual(['mail.ost'])
    expect(result.message).toContain('mailbox file')

    const removed = listRemovedPstMailboxFiles(rootDir)
    expect(removed.rootExists).toBe(true)
    expect(removed.scopes.map((scope) => scope.scopeLabel)).toEqual(['Case9 / Search9'])
    expect(removed.files.map((file) => file.fileName)).toEqual(['removed.pst'])
  })

  it('includes PST root when direct mailbox files exist there', () => {
    const rootDir = makeTempDir('pst-catalog-root-')
    fs.mkdirSync(path.join(rootDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(rootDir, 'Case1', 'Search1', 'enron.pst'))
    stageFixture(outlookPath, path.join(rootDir, 'root.ost'))

    const result = listPstMailboxFiles(rootDir)

    expect(result.scopes[0].scopePath).toBe('')
    expect(result.scopes[0].scopeLabel).toBe('PST root')
    expect(result.scopePath).toBe('')
    expect(result.files.map((file) => file.fileName)).toEqual(['root.ost'])
  })

  it('reports an empty or missing PST folder without throwing', () => {
    const emptyRoot = makeTempDir('pst-catalog-empty-')
    const emptyResult = listPstMailboxFiles(emptyRoot)
    expect(emptyResult.rootExists).toBe(true)
    expect(emptyResult.scopes).toHaveLength(0)
    expect(emptyResult.files).toHaveLength(0)
    expect(emptyResult.message).toContain('No PST or OST files')

    const missingRoot = path.join(makeTempDir('pst-catalog-missing-'), 'PST')
    const missingResult = listPstMailboxFiles(missingRoot)
    expect(missingResult.rootExists).toBe(false)
    expect(missingResult.scopes).toHaveLength(0)
    expect(missingResult.files).toHaveLength(0)
    expect(missingResult.message).toContain('Create a PST folder')
  })

  it('rejects traversal attempts and opens mailbox files within a selected scope', () => {
    const rootDir = makeTempDir('pst-catalog-open-')
    fs.mkdirSync(path.join(rootDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(rootDir, 'Case1', 'Search1', 'enron.pst'))

    expect(() => resolvePstMailboxPath(rootDir, '../Case1/Search1', 'enron.pst')).toThrow(
      'Scope path must stay within the PST folder'
    )
    expect(() => resolvePstMailboxPath(rootDir, '_removed/Case1/Search1', 'enron.pst')).toThrow(
      'Scope path must stay within the PST folder'
    )
    expect(() => resolvePstMailboxPath(rootDir, 'Case1/Search1', '../enron.pst')).toThrow(
      'Mailbox file name must not include a path'
    )
    expect(() => resolvePstMailboxPath(rootDir, 'Case1/Search1', 'notes.txt')).toThrow(
      'Only .pst and .ost files are supported'
    )

    const session = openPstMailbox(rootDir, 'Case1/Search1', 'enron.pst')
    expect(session.fileName).toBe('enron.pst')
    expect(session.mailboxName).toBe('Personal folders')
    expect(session.stats.folderCount).toBe(11)
    expect(session.stats.messageCount).toBe(71)
  })

  it('moves mailbox files into and out of the removed archive without deleting them', () => {
    const rootDir = makeTempDir('pst-catalog-archive-')
    fs.mkdirSync(path.join(rootDir, 'Case1', 'Search1'), { recursive: true })
    stageFixture(enronPath, path.join(rootDir, 'Case1', 'Search1', 'enron.pst'))

    const removal = movePstMailboxToRemoved(rootDir, 'Case1/Search1', 'enron.pst')
    expect(fs.existsSync(removal.sourcePath)).toBe(false)
    expect(fs.existsSync(removal.destinationPath)).toBe(true)
    expect(
      listPstMailboxFiles(rootDir).scopes.some((scope) => scope.scopePath === 'Case1/Search1')
    ).toBe(false)
    expect(listRemovedPstMailboxFiles(rootDir).scopes[0].files.map((file) => file.fileName)).toEqual([
      'enron.pst'
    ])

    const restore = restorePstMailboxFromRemoved(rootDir, 'Case1/Search1', 'enron.pst')
    expect(fs.existsSync(restore.sourcePath)).toBe(false)
    expect(fs.existsSync(restore.destinationPath)).toBe(true)
    expect(listPstMailboxFiles(rootDir).scopes[0].files.map((file) => file.fileName)).toEqual([
      'enron.pst'
    ])
    expect(listRemovedPstMailboxFiles(rootDir).files).toHaveLength(0)
  })
})
