import * as fs from 'fs'
import * as path from 'path'

const resolve = path.resolve

describe('viewer UI shell', () => {
  it('shows a PST list and no upload controls', () => {
    const html = fs.readFileSync(resolve('./example/public/index.html'), 'utf8')
    const script = fs.readFileSync(resolve('./example/public/app.js'), 'utf8')

    expect(html).toContain('id="refresh-catalog"')
    expect(html).toContain('id="case-select"')
    expect(html).toContain('id="case-name-label"')
    expect(html).toContain('id="scope-select"')
    expect(html).toContain('id="search-scope-select"')
    expect(html).toContain('id="scope-count-badge"')
    expect(html).toContain('id="pst-list"')
    expect(html).toContain('id="review-flagged-toggle"')
    expect(html).toContain('id="review-tagged-toggle"')
    expect(html).toContain('placeholder="Keyword, body, recipients, subject"')
    expect(html).toContain('Select a mailbox to load folders with content.')
    expect(html).not.toContain('navigator-header')
    expect(html).not.toContain('PST Browser')
    expect(script).toContain('scopeSelect')
    expect(script).toContain('caseSelect')
    expect(script).toContain('scopePath')
    expect(script).toContain('searchScope')
    expect(script).toContain('const label = `${getSearchLabel(entry.scopePath)} (${entry.fileCount})`')
    expect(script).toContain("return parts[parts.length - 1] || 'PST root'")
    expect(script).toContain('collectFoldersWithContent')
    expect(script).toContain('flat-folder-list')
    expect(script).toContain('/api/psts')
    expect(script).toContain('/api/psts/open')
    expect(script).toContain('/api/search')
    expect(script).toContain('/review')
    expect(script).toContain('toggle-review-flag')
    expect(script).toContain('indexedMessageCount')

    expect(html).not.toContain('upload-file')
    expect(html).not.toContain('upload-button')
    expect(html).not.toContain('restore-button')
    expect(html).not.toContain('drop-zone')
    expect(script).not.toContain('uploadMailbox')
    expect(script).not.toContain('restoreSession')
  })
})
