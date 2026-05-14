import * as fs from 'fs'
import * as path from 'path'

const resolve = path.resolve

describe('viewer UI shell', () => {
  it('shows a PST list and no upload controls', () => {
    const html = fs.readFileSync(resolve('./example/public/index.html'), 'utf8')
    const script = fs.readFileSync(resolve('./example/public/app.js'), 'utf8')

    expect(html).toContain('id="pst-list"')
    expect(html).toContain('id="refresh-catalog"')
    expect(html).toContain('id="review-flagged-toggle"')
    expect(html).toContain('id="review-tagged-toggle"')
    expect(script).toContain('/api/psts')
    expect(script).toContain('/api/psts/open')
    expect(script).toContain('/review')
    expect(script).toContain('toggle-review-flag')

    expect(html).not.toContain('upload-file')
    expect(html).not.toContain('upload-button')
    expect(html).not.toContain('restore-button')
    expect(html).not.toContain('drop-zone')
    expect(script).not.toContain('uploadMailbox')
    expect(script).not.toContain('restoreSession')
  })
})
