import * as fs from 'fs'
import * as path from 'path'

const resolve = path.resolve

describe('viewer UI shell', () => {
  it('builds the React shell and omits hidden-filters controls', () => {
    const html = fs.readFileSync(resolve('./example/public/index.html'), 'utf8')
    const script = fs.readFileSync(resolve('./example/public/app.js'), 'utf8')

    expect(html).toContain('content="light dark"')
    expect(html).toContain('<div id="root"></div>')
    expect(html).toContain('src="/app.js"')
    expect(html).toContain('href="/styles.css"')

    expect(script).toContain('PST Mail Explorer')
    expect(script).toContain('Sign in to continue')
    expect(script).not.toContain('hiddenFiltersDropdown')
    expect(script).not.toContain('hiddenFiltersToggle')
    expect(script).not.toContain('hiddenFiltersCount')
    expect(script).not.toContain('hiddenFiltersPanel')
    expect(script).not.toContain('data-action="close-hidden-filters"')
  })
})
