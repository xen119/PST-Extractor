const {
  normalizeMessageId,
  normalizeSearchResultItem,
  normalizeSearchResultsPage
} = require('../../example/public/search-utils.js')

describe('search result normalization', () => {
  it('uses messageId as the canonical id when id is missing', () => {
    const item = normalizeSearchResultItem({
      messageId: 'message-123',
      fileName: 'mailbox.pst'
    })

    expect(item.id).toBe('message-123')
    expect(item.messageId).toBe('message-123')
    expect(normalizeMessageId(item)).toBe('message-123')
  })

  it('normalizes search result pages so items are clickable', () => {
    const page = normalizeSearchResultsPage({
      items: [
        {
          messageId: 'message-1',
          subject: 'Example 1'
        },
        {
          id: 'message-2',
          subject: 'Example 2'
        }
      ]
    })

    expect(page.items[0].id).toBe('message-1')
    expect(page.items[0].messageId).toBe('message-1')
    expect(page.items[1].id).toBe('message-2')
    expect(page.items[1].messageId).toBe('message-2')
  })
})
