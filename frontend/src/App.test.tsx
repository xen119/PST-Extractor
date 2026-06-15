import * as React from 'react'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { App, getCasePathFromScopePath } from '@/App'
import { api } from '@/api'
import { AppShell, EmailPreview, MessageList, Sidebar, TagManagerDialog } from '@/components/layout'
import { AuthScreen, MfaReminderDialog } from '@/components/auth'
import type {
  AttachmentDetail,
  AuthStatus,
  FolderNode,
  HiddenRulesResponse,
  InviteLookupResponse,
  MessageDetail,
  MessageSummary,
  PageResponse,
  PstCatalogResponse
} from '@/types'

afterEach(() => {
  vi.restoreAllMocks()
  window.history.replaceState({}, '', '/')
})

function createDeferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void
  let reject!: (reason?: unknown) => void
  const promise = new Promise<T>((promiseResolve, promiseReject) => {
    resolve = promiseResolve
    reject = promiseReject
  })

  return { promise, resolve, reject }
}

describe('auth shell', () => {
  it('shows a minimal login screen', () => {
    render(
      <AuthScreen
        view="login"
        busy={false}
        message=""
        error=""
        invite={null}
        inviteStep="password"
        inviteMfaAvailable={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
        inviteMfaEnforced={false}
      />
    )

    expect(screen.getByRole('heading', { name: 'Sign in to continue' })).toBeInTheDocument()
    expect(screen.getByLabelText('Username')).toBeInTheDocument()
    expect(screen.getByLabelText('Password')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Sign in' })).toBeInTheDocument()
    expect(screen.getByText('PST Mail Explorer')).toBeInTheDocument()
    expect(
      screen.getByText(
        (_, element) => element?.tagName === 'P' && Boolean(element.textContent?.includes('Developed by DigiVectra DevOps'))
      )
    ).toBeInTheDocument()
    expect(screen.queryByText('Secure access')).not.toBeInTheDocument()
    expect(screen.queryByText('Local credentials')).not.toBeInTheDocument()
    expect(screen.queryByText('Username and password sign-in')).not.toBeInTheDocument()
    expect(screen.queryByText('Invite links and recovery codes are supported')).not.toBeInTheDocument()
  })

  it('shows the invite password form', () => {
    render(
      <AuthScreen
        view="invite"
        busy={false}
        message=""
        error=""
        invite={{
          username: 'jane.doe',
          createdAt: new Date().toISOString(),
          recipientEmail: 'jane@example.com',
          inviteStatus: 'pending',
          inviteSentAt: new Date().toISOString(),
          inviteExpiresAt: new Date().toISOString(),
          inviteAcceptedAt: '',
          inviteRevokedAt: '',
          mfaEnabled: false,
          mfaEnforced: false,
          mfaEnrolledAt: ''
        }}
        inviteStep="password"
        inviteMfaAvailable={true}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
        inviteMfaEnforced={false}
      />
    )

    expect(screen.getByText('Set your password')).toBeInTheDocument()
    expect(screen.getByLabelText('Password')).toBeInTheDocument()
    expect(screen.getByLabelText('Confirm password')).toBeInTheDocument()
  })

  it('shows a minimal mfa verification screen', () => {
    render(
      <AuthScreen
        view="mfa"
        busy={false}
        message="Enter the verification code for admin."
        error=""
        invite={null}
        inviteStep="password"
        inviteMfaAvailable={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
        inviteMfaEnforced={false}
      />
    )

    expect(screen.getByRole('heading', { name: 'Verify your sign in' })).toBeInTheDocument()
    expect(screen.getByLabelText('Verification code')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Verify code' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Back to sign in' })).toBeInTheDocument()
    expect(screen.queryByText('Enter the verification code for admin.')).not.toBeInTheDocument()
    expect(screen.queryByText('Use your authenticator app or a recovery code.')).not.toBeInTheDocument()
    expect(screen.queryByText('PST Mail Explorer')).not.toBeInTheDocument()
  })

  it('clears invite loading after the invite lookup succeeds', async () => {
    const inviteLookupResponse: InviteLookupResponse = {
      invite: {
        username: 'jane.doe',
        createdAt: new Date().toISOString(),
        recipientEmail: 'jane@example.com',
        inviteStatus: 'pending',
        inviteSentAt: new Date().toISOString(),
        inviteExpiresAt: new Date().toISOString(),
        inviteAcceptedAt: '',
        inviteRevokedAt: '',
        mfaEnabled: false,
        mfaEnforced: false,
        mfaEnrolledAt: ''
      }
    }
    const anonymousStatus: AuthStatus = {
      authenticated: false,
      enabled: true,
      canManageUsers: false,
      mfaEnabled: false,
      mfaEnforced: false,
      user: null,
      expiresAt: null
    }

    vi.spyOn(api.auth, 'inviteLookup').mockResolvedValueOnce(inviteLookupResponse)
    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(anonymousStatus)
    window.history.pushState({}, '', '/invite/test-token')

    render(<App />)

    expect(await screen.findByRole('heading', { name: 'Set your password' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Set password' })).toBeEnabled()
  })

  it('blocks the workspace while the search index refresh runs', async () => {
    const user = userEvent.setup()
    const refreshDeferred = createDeferred<unknown>()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      user: { username: 'admin' },
      expiresAt: null
    }
    const emptyCatalog: PstCatalogResponse = {
      rootPath: '',
      rootExists: true,
      message: '',
      scopePath: '',
      scopeLabel: 'PST root',
      scopes: [],
      files: []
    }
    const hiddenRulesResponse: HiddenRulesResponse = { items: [] }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(emptyCatalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.pst, 'refreshSearchIndex').mockReturnValueOnce(refreshDeferred.promise)

    render(<App />)

    const refreshButton = await screen.findByRole('button', { name: 'Refresh search index' })
    await user.click(refreshButton)

    expect(await screen.findByRole('heading', { name: 'Rebuilding search index' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Refresh search index' })).not.toBeInTheDocument()

    refreshDeferred.resolve({})

    await screen.findByRole('button', { name: 'Refresh search index' })
    expect(screen.queryByRole('heading', { name: 'Rebuilding search index' })).not.toBeInTheDocument()
  })

  it('shows an error state when search index refresh fails and allows dismissal', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      user: { username: 'admin' },
      expiresAt: null
    }
    const emptyCatalog: PstCatalogResponse = {
      rootPath: '',
      rootExists: true,
      message: '',
      scopePath: '',
      scopeLabel: 'PST root',
      scopes: [],
      files: []
    }
    const hiddenRulesResponse: HiddenRulesResponse = { items: [] }
    const refreshError = new Error('Refresh failed')

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(emptyCatalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.pst, 'refreshSearchIndex').mockRejectedValueOnce(refreshError)

    render(<App />)

    await user.click(await screen.findByRole('button', { name: 'Refresh search index' }))

    expect(await screen.findByRole('heading', { name: 'Reindex failed' })).toBeInTheDocument()
    expect(screen.getByText('Refresh failed')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Dismiss' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Retry' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Refresh search index' })).not.toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Dismiss' }))

    await screen.findByRole('button', { name: 'Refresh search index' })
    expect(screen.queryByRole('heading', { name: 'Reindex failed' })).not.toBeInTheDocument()
  })

  it('shows a blocking MFA reminder without a close button', () => {
    render(<MfaReminderDialog open username="admin" allowSkip onSetup={vi.fn()} onSkip={vi.fn()} />)

    expect(screen.getByRole('heading', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Skip for now' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Close' })).not.toBeInTheDocument()
  })

  it('hides skip when MFA is enforced', () => {
    render(<MfaReminderDialog open username="admin" allowSkip={false} onSetup={vi.fn()} onSkip={vi.fn()} />)

    expect(screen.getByRole('heading', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Skip for now' })).not.toBeInTheDocument()
  })
})

describe('shell and preview', () => {
  it('keeps the top-level case path when a search folder is selected', () => {
    expect(getCasePathFromScopePath('ED - Investigation regarding DR et al/Emails_to_from_davidradford5_hotmail_co_uk')).toBe(
      'ED - Investigation regarding DR et al'
    )
  })

  const sampleRow: MessageSummary = {
    id: 'message-1',
    subject: 'Quarterly update',
    senderName: 'Alice Example',
    senderEmailAddress: 'alice@example.com',
    recipientText: 'To: team@example.com',
    sortDate: new Date().toISOString(),
    hasAttachments: true,
    isRead: false,
    isMailLike: true,
    review: {
      flagged: true,
      tags: ['Legal'],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  }

  const sampleDetail: MessageDetail = {
    subject: 'Quarterly update',
    senderName: 'Alice Example',
    senderEmailAddress: 'alice@example.com',
    displayTo: 'Team',
    sortDate: new Date().toISOString(),
    bodyText: 'Hello world',
    attachments: [
      {
        attachmentId: 'attach-1',
        index: 0,
        filename: 'invoice.pdf',
        mimeTag: 'application/pdf',
        size: 1024,
        isDownloadable: false
      } satisfies AttachmentDetail
    ],
    review: {
      flagged: true,
      tags: ['Legal'],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  }

  it('shows icon controls and settings items in the shell', async () => {
    const user = userEvent.setup()
    render(
      <AppShell
        userName="admin"
        authenticated
        settingsMenu={
          <div>
            <button type="button">Manage users</button>
            <button type="button">Switch to dark mode</button>
          </div>
        }
        breadcrumbs={[{ label: 'Case A' }, { label: 'Mailbox' }]}
        sidebarCollapsed={false}
        previewCollapsed={false}
        onLogout={vi.fn()}
        sidebar={<div>Sidebar</div>}
        messagePanel={<div>Messages</div>}
        preview={<div>Preview</div>}
      />
    )

    expect(screen.queryByLabelText('Hide sidebar')).not.toBeInTheDocument()
    expect(screen.queryByLabelText('Hide preview')).not.toBeInTheDocument()
    expect(screen.getByLabelText('Logout')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'admin' })).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'admin' }))
    expect(await screen.findByText('Manage users')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Switch to dark mode' })).toBeInTheDocument()
  })

  it('shows label-only case and search selectors with the mailbox dropdown', () => {
    render(
      <div style={{ width: 420, height: 800 }}>
        <Sidebar
          caseOptions={[
            { label: 'Case Alpha', value: 'Case Alpha', count: 3 },
            { label: 'Case Beta', value: 'Case Beta', count: 2 }
          ]}
          selectedCasePath="Case Alpha"
          selectedScopePath="Case Alpha/Search One"
          searchOptions={[
            { label: 'Search One', value: 'Case Alpha/Search One', count: 2 },
            { label: 'Search Two', value: 'Case Alpha/Search Two', count: 1 }
          ]}
          catalogFiles={[
            {
              fileName: 'mailbox-a.pst',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One'
            }
          ]}
          selectedPstFileName="mailbox-a.pst"
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          canRefreshSearchIndex
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    expect(screen.getByRole('option', { name: 'Case Alpha' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'Search One' })).toBeInTheDocument()
    expect(screen.queryByText('Case Alpha (3)')).not.toBeInTheDocument()
    expect(screen.queryByText('Search One (2)')).not.toBeInTheDocument()
    expect(screen.getByRole('combobox', { name: 'Mailbox selector' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Refresh search index' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'mailbox-a.pst' })).toBeInTheDocument()
    expect(screen.queryByText('Found 2 mailbox files in Case Alpha / Search One.')).not.toBeInTheDocument()
  })

  it('hides the search index refresh button for non-admin users', () => {
    render(
      <div style={{ width: 420, height: 800 }}>
        <Sidebar
          caseOptions={[{ label: 'Case Alpha', value: 'Case Alpha', count: 1 }]}
          selectedCasePath="Case Alpha"
          selectedScopePath="Case Alpha/Search One"
          searchOptions={[{ label: 'Search One', value: 'Case Alpha/Search One', count: 1 }]}
          catalogFiles={[]}
          selectedPstFileName=""
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          canRefreshSearchIndex={false}
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    expect(screen.queryByRole('button', { name: 'Refresh search index' })).not.toBeInTheDocument()
  })

  it('hides empty folders and keeps the mailbox selector compact', () => {
    const tree: FolderNode = {
      id: 'root',
      displayName: 'Root',
      path: 'root',
      indexedMessageCount: 4,
      mailMessageCount: 4,
      children: [
        {
          id: 'empty',
          displayName: 'Empty',
          path: 'root/empty',
          indexedMessageCount: 0,
          mailMessageCount: 0,
          children: []
        },
        {
          id: 'inbox',
          displayName: 'Inbox',
          path: 'root/inbox',
          indexedMessageCount: 2,
          mailMessageCount: 2,
          children: []
        }
      ]
    }

    render(
      <div style={{ width: 420, height: 900 }}>
        <Sidebar
          caseOptions={[{ label: 'Case Alpha', value: 'Case Alpha', count: 1 }]}
          selectedCasePath="Case Alpha"
          selectedScopePath="Case Alpha/Search One"
          searchOptions={[{ label: 'Search One', value: 'Case Alpha/Search One', count: 1 }]}
          catalogFiles={[
            {
              fileName: 'mailbox.pst',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One',
              displayPath: 'PST root'
            }
          ]}
          selectedPstFileName="mailbox.pst"
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          canRefreshSearchIndex
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          folderTree={tree}
          currentFolderId="inbox"
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    expect(screen.queryByText('Selected')).not.toBeInTheDocument()
    expect(screen.queryByText('Open')).not.toBeInTheDocument()
    expect(screen.queryByText('PST root')).not.toBeInTheDocument()
    expect(screen.getByRole('combobox', { name: 'Mailbox selector' })).toBeInTheDocument()
    expect(screen.queryByText('Empty')).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: /Inbox - 2/ })).toBeInTheDocument()
    expect(screen.queryByText(/- 0/)).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Refresh search index' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Show removed PSTs' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Remove PST' })).not.toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Restore PST' })).not.toBeInTheDocument()
  })

  it('shows the message list and preview details', async () => {
    const user = userEvent.setup()
    const onOpenFullView = vi.fn()
    render(
      <div style={{ width: 1600, height: 1000 }}>
        <MessageList
        page={{
          items: [sampleRow],
          total: 1,
          page: 1,
          totalPages: 1,
          query: ''
        }}
          loading={false}
          query=""
          searchScope="pst"
          mailOnly={false}
          sort="date-desc"
          reviewFlaggedOnly={false}
          reviewTaggedOnly={false}
          onQueryChange={vi.fn()}
          onSearch={vi.fn()}
          onSearchScopeChange={vi.fn()}
          onMailOnlyChange={vi.fn()}
          onSortChange={vi.fn()}
          onReviewFlaggedChange={vi.fn()}
          onReviewTaggedChange={vi.fn()}
          onSelectMessage={vi.fn()}
          onPrevPage={vi.fn()}
          onNextPage={vi.fn()}
          onOpenBundle={vi.fn()}
          selectedMessageId="message-1"
          sessionId="session-1"
        />
        <EmailPreview
          detail={sampleDetail}
          theme="light"
          onDownloadJson={vi.fn()}
          onDownloadEml={vi.fn()}
          onToggleFlag={vi.fn()}
          onClearReview={vi.fn()}
          onOpenTags={vi.fn()}
          onOpenFullView={onOpenFullView}
          tagCount={sampleDetail.review?.tags?.length || 0}
          onOpenAttachment={vi.fn()}
          onOpenPrev={vi.fn()}
          onOpenNext={vi.fn()}
          canNavigatePrev={false}
          canNavigateNext={false}
        />
      </div>
    )

    expect(screen.getByText('Quarterly update')).toBeInTheDocument()
    expect(screen.getByText('Alice Example')).toBeInTheDocument()
    expect(screen.getByText('Flagged')).toBeInTheDocument()
    expect(screen.getByText('Unavailable')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Manage tags' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Open full view' })).toBeInTheDocument()
    expect(screen.getByLabelText('Run search')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Open full view' }))
    expect(onOpenFullView).toHaveBeenCalled()
    await user.click(screen.getByLabelText('Download JSON'))
    expect(screen.getByLabelText('Recipients')).toBeInTheDocument()
  })

  it('opens a search result from a different mailbox in the reading pane', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      user: { username: 'admin' },
      expiresAt: null
    }
    const catalog: PstCatalogResponse = {
      rootPath: '',
      rootExists: true,
      message: '',
      scopePath: 'Case Alpha/Search One',
      scopeLabel: 'Search One',
      scopes: [
        {
          scopePath: 'Case Alpha/Search One',
          scopeLabel: 'Search One',
          fileCount: 1,
          files: [
            {
              fileName: 'mailbox-a.pst',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One'
            }
          ]
        },
        {
          scopePath: 'Case Beta/Search Two',
          scopeLabel: 'Search Two',
          fileCount: 1,
          files: [
            {
              fileName: 'mailbox-b.pst',
              size: 2048,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Beta/Search Two'
            }
          ]
        }
      ],
      files: []
    }
    const currentDetail: MessageDetail = {
      subject: 'Current mailbox message',
      senderName: 'Alice Example',
      senderEmailAddress: 'alice@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'Current mailbox body'
    }
    const targetDetail: MessageDetail = {
      subject: 'Target search result',
      senderName: 'Bob Example',
      senderEmailAddress: 'bob@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'Target mailbox body'
    }
    const searchResultsPage: PageResponse<MessageSummary> = {
      items: [
        {
          id: 'search-hit',
          messageId: 'search-hit',
          subject: 'Target search result',
          senderName: 'Bob Example',
          senderEmailAddress: 'bob@example.com',
          sortDate: new Date().toISOString(),
          fileName: 'mailbox-b.pst',
          scopePath: 'Case Beta/Search Two',
          isMailLike: true,
          review: {
            flagged: false,
            tags: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }
        }
      ],
      total: 1,
      page: 1,
      pageSize: 50,
      totalPages: 1,
      query: 'target',
      mailOnly: false,
      sort: 'date-desc'
    }
    const searchResponse = {
      scope: 'search' as const,
      scopePath: 'Case Alpha/Search One',
      scopeLabel: 'Search One',
      page: searchResultsPage
    }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(catalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue({ items: [] })
    vi.spyOn(api.pst, 'open').mockImplementation(async (_scopePath, fileName) => {
      if (fileName === 'mailbox-b.pst') {
        return {
          sessionId: 'session-b',
          scopePath: 'Case Beta/Search Two',
          scopeLabel: 'Search Two',
          fileName: 'mailbox-b.pst',
          summary: {
            fileName: 'mailbox-b.pst',
            mailboxName: 'mailbox-b.pst'
          },
          tree: {
            id: 'root-b',
            displayName: 'Inbox',
            path: 'root-b',
            children: []
          }
        }
      }

      return {
        sessionId: 'session-a',
        scopePath: 'Case Alpha/Search One',
        scopeLabel: 'Search One',
        fileName: 'mailbox-a.pst',
        summary: {
          fileName: 'mailbox-a.pst',
          mailboxName: 'mailbox-a.pst'
        },
        tree: {
          id: 'root-a',
          displayName: 'Inbox',
          path: 'root-a',
          children: []
        }
      }
    })
    vi.spyOn(api.session, 'folderMessages').mockResolvedValue({
      sessionId: 'session-a',
      page: {
        items: [
          {
            id: 'current-message',
            messageId: 'current-message',
            subject: 'Current mailbox message',
            senderName: 'Alice Example',
            senderEmailAddress: 'alice@example.com',
            sortDate: new Date().toISOString(),
            fileName: 'mailbox-a.pst',
            scopePath: 'Case Alpha/Search One',
            isMailLike: true,
            review: {
              flagged: false,
              tags: [],
              createdAt: new Date().toISOString(),
              updatedAt: new Date().toISOString()
            }
          } satisfies MessageSummary
        ],
        total: 1,
        page: 1,
        pageSize: 50,
        totalPages: 1,
        query: '',
        mailOnly: false,
        sort: 'date-desc'
      }
    })
    vi.spyOn(api, 'search').mockResolvedValue(searchResponse)
    vi.spyOn(api.session, 'messageDetail').mockImplementation(async (sessionId, messageId) => {
      if (sessionId === 'session-b' && messageId === 'search-hit') {
        return { sessionId, detail: targetDetail }
      }
      return { sessionId, detail: currentDetail }
    })

    window.history.replaceState({}, '', '/')
    render(
      <div style={{ width: 1800, height: 1600 }}>
        <App />
      </div>
    )

    expect(await screen.findByText('Current mailbox message')).toBeInTheDocument()

    await user.type(screen.getByPlaceholderText('Keywords, "phrases", + AND, | OR'), 'target')
    await user.selectOptions(screen.getByRole('combobox', { name: 'Scope' }), 'all')
    await user.click(screen.getByRole('button', { name: 'Run search' }))

    expect(await screen.findByText('Target mailbox body')).toBeInTheDocument()
    expect(api.pst.open).toHaveBeenCalledWith('Case Beta/Search Two', 'mailbox-b.pst')
    expect(screen.getByText('Target mailbox body')).toBeInTheDocument()
  })

  it('shows the tag manager modal', async () => {
    const user = userEvent.setup()
    const onOpenChange = vi.fn()
    render(
      <TagManagerDialog
        open
        tags={['Legal', 'Review']}
        subject="Quarterly update"
        onOpenChange={onOpenChange}
        onAddTag={vi.fn()}
        onRemoveTag={vi.fn()}
      />
    )

    expect(screen.getByRole('heading', { name: 'Tags' })).toBeInTheDocument()
    expect(screen.getByText('Quarterly update')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Remove Legal' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Remove Review' })).toBeInTheDocument()
    expect(screen.getByPlaceholderText('Add tag')).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: /close/i }))
    expect(onOpenChange).toHaveBeenCalledWith(false)
  })
})
