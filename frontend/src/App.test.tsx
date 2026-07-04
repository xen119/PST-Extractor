import * as React from 'react'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { App, getCasePathFromScopePath } from '@/App'
import { api } from '@/api'
import { AppShell, EmailPreview, MessageList, Sidebar, TagManagerDialog } from '@/components/layout'
import { AuthScreen, MfaReminderDialog } from '@/components/auth'
import type {
  ActivityLogResponse,
  AttachmentDetail,
  AuthStatus,
  FolderNode,
  HiddenRulesResponse,
  InviteLookupResponse,
  MessageDetail,
  MessageSummary,
  PageResponse,
  PasswordResetConfirmResponse,
  PasswordResetLookupResponse,
  PstCatalogResponse,
  SearchIndexRefreshStatus,
  UsersResponse
} from '@/types'

vi.mock('@tanstack/react-virtual', () => ({
  useVirtualizer: ({ count }: { count: number }) => ({
    getTotalSize: () => count * 96,
    getVirtualItems: () =>
      Array.from({ length: count }, (_, index) => ({
        index,
        start: index * 96
      })),
    measureElement: vi.fn()
  })
}))

afterEach(() => {
  vi.restoreAllMocks()
  try {
    window.localStorage?.clear()
  } catch {
    // Ignore storage access errors in the test environment.
  }
  try {
    window.sessionStorage?.clear()
  } catch {
    // Ignore storage access errors in the test environment.
  }
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
        passwordResetAvailable={false}
        invite={null}
        inviteStep="password"
        inviteMfaAvailable={false}
        inviteMfaEnforced={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        resetLookup={null}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onPasswordResetRequest={async () => undefined}
        onPasswordResetConfirm={async () => undefined}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
      />
    )

    expect(screen.getByRole('heading', { name: 'Sign in to continue' })).toBeInTheDocument()
    expect(screen.getByLabelText('Username')).toBeInTheDocument()
    expect(screen.getByLabelText('Username')).toHaveValue('')
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
    expect(screen.queryByRole('button', { name: 'Forgot password?' })).not.toBeInTheDocument()
  })

  it('shows the invite password form', () => {
    render(
      <AuthScreen
        view="invite"
        busy={false}
        message=""
        error=""
        passwordResetAvailable={false}
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
        inviteMfaEnforced={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        resetLookup={null}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onPasswordResetRequest={async () => undefined}
        onPasswordResetConfirm={async () => undefined}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
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
        passwordResetAvailable={false}
        invite={null}
        inviteStep="password"
        inviteMfaAvailable={false}
        inviteMfaEnforced={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        resetLookup={null}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onPasswordResetRequest={async () => undefined}
        onPasswordResetConfirm={async () => undefined}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
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

  it('opens the password reset request dialog from the login screen', async () => {
    const user = userEvent.setup()
    const onPasswordResetRequest = vi.fn().mockResolvedValue(undefined)

    render(
      <AuthScreen
        view="login"
        busy={false}
        message=""
        error=""
        passwordResetAvailable={true}
        invite={null}
        inviteStep="password"
        inviteMfaAvailable={false}
        inviteMfaEnforced={false}
        inviteSetup={null}
        inviteRecoveryCodes={[]}
        resetLookup={null}
        onLogin={vi.fn()}
        onMfaChallenge={vi.fn()}
        onPasswordResetRequest={onPasswordResetRequest}
        onPasswordResetConfirm={async () => undefined}
        onInviteAccept={vi.fn()}
        onInviteMfaStart={vi.fn()}
        onInviteMfaSkip={vi.fn()}
        onInviteMfaSubmit={vi.fn()}
        onInviteFinish={vi.fn()}
        onOpenLogin={vi.fn()}
      />
    )

    await user.click(screen.getByRole('button', { name: 'Forgot password?' }))
    expect(await screen.findByRole('heading', { name: 'Reset password' })).toBeInTheDocument()

    const requestInput = screen.getByLabelText('Username or email')
    await user.clear(requestInput)
    await user.type(requestInput, ' jane.doe@example.com ')
    await user.click(screen.getByRole('button', { name: 'Send reset link' }))

    expect(onPasswordResetRequest).toHaveBeenCalledWith('jane.doe@example.com')
    await waitFor(() => {
      expect(screen.queryByRole('heading', { name: 'Reset password' })).not.toBeInTheDocument()
    })
  })

  it('loads the password reset screen from a reset token and completes the reset', async () => {
    const user = userEvent.setup()
    const resetLookupResponse: PasswordResetLookupResponse = {
      reset: {
        username: 'jane.doe',
        recipientEmail: 'jane@example.com'
      }
    }
    const resetConfirmResponse: PasswordResetConfirmResponse = {
      user: {
        username: 'jane.doe',
        assignedCasePaths: []
      },
      message: 'Password updated'
    }

    vi.spyOn(api.auth, 'passwordResetLookup').mockResolvedValueOnce(resetLookupResponse)
    vi.spyOn(api.auth, 'passwordResetConfirm').mockResolvedValueOnce(resetConfirmResponse)
    window.history.pushState({}, '', '/reset/reset-token')

    render(<App />)

    expect(await screen.findByRole('heading', { name: 'Set a new password' })).toBeInTheDocument()
    expect(screen.getByText('jane.doe')).toBeInTheDocument()
    expect(screen.getByText('jane@example.com')).toBeInTheDocument()

    await user.type(screen.getByLabelText('New password'), 'NewPass123!')
    await user.type(screen.getByLabelText('Confirm password'), 'NewPass123!')
    await user.click(screen.getByRole('button', { name: 'Update password' }))

    expect(await screen.findByRole('heading', { name: 'Sign in to continue' })).toBeInTheDocument()
    expect(api.auth.passwordResetConfirm).toHaveBeenCalledWith('reset-token', 'NewPass123!', 'NewPass123!')
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
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
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

  it('shows a background refresh status without blocking the workspace', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
    const idleStatus: SearchIndexRefreshStatus = {
      jobId: null,
      status: 'idle',
      trigger: null,
      startedAt: null,
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }
    const runningStatus: SearchIndexRefreshStatus = {
      jobId: 'job-1',
      status: 'running',
      trigger: 'manual',
      startedAt: new Date().toISOString(),
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }
    const succeededStatus: SearchIndexRefreshStatus = {
      jobId: 'job-1',
      status: 'succeeded',
      trigger: 'manual',
      startedAt: runningStatus.startedAt,
      completedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      summary: { mailboxCount: 1, messageCount: 1 },
      error: null
    }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(emptyCatalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.pst, 'refreshSearchIndexStatus')
      .mockResolvedValueOnce({ status: idleStatus })
      .mockResolvedValueOnce({ status: idleStatus })
      .mockResolvedValueOnce({ status: runningStatus })
      .mockResolvedValueOnce({ status: succeededStatus })
    vi.spyOn(api.pst, 'refreshSearchIndex').mockResolvedValue({ status: runningStatus })

    render(<App />)

    const refreshButton = await screen.findByRole('button', { name: 'Reindex mailboxes' })
    await user.click(refreshButton)

    expect(await screen.findByText('Reindexing mailboxes')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Reindex mailboxes' })).toBeDisabled()

    await waitFor(
      () => {
        expect(screen.queryByText('Reindexing mailboxes')).not.toBeInTheDocument()
      },
      { timeout: 5000 }
    )
    expect(screen.getByRole('button', { name: 'Reindex mailboxes' })).toBeEnabled()
  })

  it('shows an error state when search index refresh fails and keeps the workspace usable', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
    const idleStatus: SearchIndexRefreshStatus = {
      jobId: null,
      status: 'idle',
      trigger: null,
      startedAt: null,
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }
    const runningStatus: SearchIndexRefreshStatus = {
      jobId: 'job-2',
      status: 'running',
      trigger: 'manual',
      startedAt: new Date().toISOString(),
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }
    const failedStatus: SearchIndexRefreshStatus = {
      jobId: 'job-2',
      status: 'failed',
      trigger: 'manual',
      startedAt: runningStatus.startedAt,
      completedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      summary: null,
      error: 'Refresh failed'
    }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(emptyCatalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.pst, 'refreshSearchIndexStatus')
      .mockResolvedValueOnce({ status: idleStatus })
      .mockResolvedValueOnce({ status: idleStatus })
      .mockResolvedValueOnce({ status: runningStatus })
      .mockResolvedValueOnce({ status: failedStatus })
    vi.spyOn(api.pst, 'refreshSearchIndex').mockResolvedValue({ status: runningStatus })

    render(<App />)

    await user.click(await screen.findByRole('button', { name: 'Reindex mailboxes' }))

    expect(await screen.findByText('Reindexing mailboxes')).toBeInTheDocument()
    await waitFor(
      () => {
        expect(screen.getByText('Mailboxes reindex failed')).toBeInTheDocument()
      },
      { timeout: 5000 }
    )
    expect(screen.getByText('Refresh failed')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Reindex mailboxes' })).toBeEnabled()
  })

  it('opens case access in a per-user modal and defaults to no cases', async () => {
    const user = userEvent.setup()
    const now = new Date().toISOString()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
    const usersResponse: UsersResponse = {
      users: [
        {
          username: 'admin',
          createdAt: now,
          recipientEmail: 'admin@example.com',
          inviteStatus: 'active',
          inviteSentAt: now,
          inviteExpiresAt: '',
          inviteAcceptedAt: now,
          inviteRevokedAt: '',
          mfaEnabled: true,
          mfaEnforced: false,
          mfaEnrolledAt: now,
          assignedCasePaths: ['Case1']
        },
        {
          username: 'bob',
          createdAt: now,
          recipientEmail: 'bob@example.com',
          inviteStatus: 'active',
          inviteSentAt: now,
          inviteExpiresAt: '',
          inviteAcceptedAt: now,
          inviteRevokedAt: '',
          mfaEnabled: false,
          mfaEnforced: false,
          mfaEnrolledAt: '',
          assignedCasePaths: []
        }
      ]
    }
    const activityResponse: ActivityLogResponse = { entries: [] }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(emptyCatalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.auth, 'users').mockResolvedValue(usersResponse)
    vi.spyOn(api.activityLog, 'list').mockResolvedValue(activityResponse)

    render(<App />)

    await user.click(await screen.findByRole('button', { name: 'admin' }))
    expect(await screen.findByRole('button', { name: 'Download items CSV' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Clear all flags' })).toBeInTheDocument()
    await user.click(await screen.findByRole('button', { name: 'Manage users' }))

    expect(await screen.findByRole('heading', { name: 'User management' })).toBeInTheDocument()
    expect(screen.getByText('No cases assigned')).toBeInTheDocument()

    await user.click(screen.getByRole('button', { name: 'Manage case access for bob' }))

    expect(await screen.findByRole('heading', { name: 'Case access' })).toBeInTheDocument()
    expect(screen.getByText('Users start with no cases assigned. Leave everything unchecked for no access.')).toBeInTheDocument()
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
          sourceType="mailbox"
          sourceCounts={{ mailbox: 1, teams: 0, sharepoint: 0 }}
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
          searchIndexRefreshStatuses={{ mailboxes: null, items: null }}
          searchIndexRefreshBusyBySource={{ mailboxes: false, items: false }}
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          onSourceTypeChange={vi.fn()}
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
    expect(screen.getByRole('button', { name: 'Reindex mailboxes' })).toBeInTheDocument()
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
          sourceType="mailbox"
          sourceCounts={{ mailbox: 0, teams: 0, sharepoint: 0 }}
          searchOptions={[{ label: 'Search One', value: 'Case Alpha/Search One', count: 1 }]}
          catalogFiles={[]}
          selectedPstFileName=""
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          canRefreshSearchIndex={false}
          searchIndexRefreshStatuses={{ mailboxes: null, items: null }}
          searchIndexRefreshBusyBySource={{ mailboxes: false, items: false }}
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          onSourceTypeChange={vi.fn()}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    expect(screen.queryByRole('button', { name: 'Reindex mailboxes' })).not.toBeInTheDocument()
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
          sourceType="mailbox"
          sourceCounts={{ mailbox: 1, teams: 0, sharepoint: 0 }}
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
          searchIndexRefreshStatuses={{ mailboxes: null, items: null }}
          searchIndexRefreshBusyBySource={{ mailboxes: false, items: false }}
          onRefreshSearchIndex={vi.fn()}
          onOpenMailbox={vi.fn()}
          onSourceTypeChange={vi.fn()}
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
    expect(screen.getByRole('button', { name: 'Reindex mailboxes' })).toBeInTheDocument()
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
          sourceType="mailbox"
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

    expect(screen.getByRole('button', { name: /Quarterly update/ })).toBeInTheDocument()
    expect(screen.getByText('Hello world')).toBeInTheDocument()
    expect(screen.getAllByText('Flagged')).toHaveLength(2)
    expect(screen.getByText('Unavailable')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Manage tags' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Open full view' })).toBeInTheDocument()
    expect(screen.getByLabelText('Run search')).toBeInTheDocument()
    await user.click(screen.getByRole('button', { name: 'Open full view' }))
    expect(onOpenFullView).toHaveBeenCalled()
    await user.click(screen.getByLabelText('Download JSON'))
    expect(screen.getByLabelText('Recipients')).toBeInTheDocument()
  })

  it('shows a loading placeholder and ignores stale preview responses when selection changes quickly', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
              fileName: 'items.zip',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One'
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
    const firstDetail = createDeferred<{ detail: MessageDetail }>()
    const secondDetail = createDeferred<{ detail: MessageDetail }>()
    const firstPreview: MessageDetail = {
      subject: 'Archive item A',
      senderName: 'Team A',
      senderEmailAddress: 'team-a@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'First item body',
      sourceType: 'sharepoint',
      archivePath: 'Case Alpha/Search One/items.zip',
      archiveEntryName: 'archive-a.json',
      downloadUrl: '/api/items/archive-a/content'
    }
    const secondPreview: MessageDetail = {
      subject: 'Archive item B',
      senderName: 'Team B',
      senderEmailAddress: 'team-b@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'Second item body',
      sourceType: 'sharepoint',
      archivePath: 'Case Alpha/Search One/items.zip',
      archiveEntryName: 'archive-b.json',
      downloadUrl: '/api/items/archive-b/content'
    }
    const hiddenRulesResponse: HiddenRulesResponse = { items: [] }
    const idleStatus: SearchIndexRefreshStatus = {
      jobId: null,
      status: 'idle',
      trigger: null,
      startedAt: null,
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(catalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue(hiddenRulesResponse)
    vi.spyOn(api.pst, 'refreshSearchIndexStatus').mockResolvedValue({ status: idleStatus })
    vi.spyOn(api.pst, 'open').mockResolvedValue({
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
          } satisfies MessageSummary,
          {
            id: 'archive-a',
            messageId: 'archive-a',
            subject: 'Archive item A',
            senderName: 'Team A',
            senderEmailAddress: 'team-a@example.com',
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
          } satisfies MessageSummary,
          {
            id: 'archive-b',
            messageId: 'archive-b',
            subject: 'Archive item B',
            senderName: 'Team B',
            senderEmailAddress: 'team-b@example.com',
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
        total: 3,
        page: 1,
        pageSize: 50,
        totalPages: 1,
        query: '',
        mailOnly: false,
        sort: 'date-desc'
      }
    })
    vi.spyOn(api.session, 'messageDetail').mockImplementation(async (sessionId, messageId) => {
      if (sessionId !== 'session-a') {
        throw new Error(`Unexpected session: ${sessionId}`)
      }
      if (messageId === 'current-message') {
        return { sessionId, detail: currentDetail }
      }
      if (messageId === 'archive-a') {
        const response = await firstDetail.promise
        return { sessionId, detail: response.detail }
      }
      if (messageId === 'archive-b') {
        const response = await secondDetail.promise
        return { sessionId, detail: response.detail }
      }
      throw new Error(`Unexpected message detail request: ${sessionId}/${messageId}`)
    })

    window.localStorage.setItem('pst-mail-explorer.casePath::admin', 'Case Alpha')
    window.localStorage.setItem('pst-mail-explorer.scopePath::admin', 'Case Alpha/Search One')
    window.localStorage.setItem('pst-mail-explorer.pstFileName::admin', 'mailbox-a.pst')
    window.localStorage.setItem('pst-mail-explorer.folderId::admin', 'root-a')
    window.localStorage.setItem('pst-mail-explorer.messageId::admin', 'current-message')
    window.localStorage.setItem('pst-mail-explorer.sourceType::admin', 'mailbox')

    render(
      <div style={{ width: 1800, height: 5000 }}>
        <App />
      </div>
    )

    expect(await screen.findByText('Current mailbox body')).toBeInTheDocument()
    const firstItem = await screen.findByRole('button', { name: /Archive item A/ })
    await user.click(firstItem)
    expect(await screen.findByText('Loading preview...')).toBeInTheDocument()

    const secondItem = await screen.findByRole('button', { name: /Archive item B/ })
    await user.click(secondItem)
    expect(screen.getByText('Loading preview...')).toBeInTheDocument()

    secondDetail.resolve({ detail: secondPreview })
    expect(await screen.findByText('Second item body')).toBeInTheDocument()

    firstDetail.resolve({ detail: firstPreview })
    await waitFor(() => {
      expect(screen.getByText('Second item body')).toBeInTheDocument()
      expect(screen.queryByText('First item body')).not.toBeInTheDocument()
    })
  })

  it('hydrates restored mailbox selections and reuses warmed preview requests for next navigation', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: false,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
        }
      ],
      files: []
    }
    const idleStatus: SearchIndexRefreshStatus = {
      jobId: null,
      status: 'idle',
      trigger: null,
      startedAt: null,
      completedAt: null,
      updatedAt: new Date().toISOString(),
      summary: null,
      error: null
    }
    const firstDetail = createDeferred<{ sessionId: string; detail: MessageDetail }>()
    const secondDetail = createDeferred<{ sessionId: string; detail: MessageDetail }>()
    const thirdDetail = createDeferred<{ sessionId: string; detail: MessageDetail }>()
    const detailRequests: Record<string, typeof firstDetail> = {
      'message-1': firstDetail,
      'message-2': secondDetail,
      'message-3': thirdDetail
    }
    const firstPreview: MessageDetail = {
      subject: 'First mailbox result',
      senderName: 'Alice Example',
      senderEmailAddress: 'alice@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'First mailbox body'
    }
    const secondPreview: MessageDetail = {
      subject: 'Second mailbox result',
      senderName: 'Bob Example',
      senderEmailAddress: 'bob@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'Second mailbox body'
    }
    const thirdPreview: MessageDetail = {
      subject: 'Third mailbox result',
      senderName: 'Carol Example',
      senderEmailAddress: 'carol@example.com',
      sortDate: new Date().toISOString(),
      bodyText: 'Third mailbox body'
    }
    const folderPage: PageResponse<MessageSummary> = {
      items: [
        {
          id: 'message-1',
          messageId: 'message-1',
          subject: 'First mailbox result',
          senderName: 'Alice Example',
          senderEmailAddress: 'alice@example.com',
          sortDate: new Date().toISOString(),
          fileName: 'mailbox-a.pst',
          scopePath: 'Case Alpha/Search One',
          sourceType: 'mailbox',
          isMailLike: true,
          review: {
            flagged: false,
            tags: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }
        } satisfies MessageSummary,
        {
          id: 'message-2',
          messageId: 'message-2',
          subject: 'Second mailbox result',
          senderName: 'Bob Example',
          senderEmailAddress: 'bob@example.com',
          sortDate: new Date().toISOString(),
          fileName: 'mailbox-a.pst',
          scopePath: 'Case Alpha/Search One',
          sourceType: 'mailbox',
          isMailLike: true,
          review: {
            flagged: false,
            tags: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }
        } satisfies MessageSummary,
        {
          id: 'message-3',
          messageId: 'message-3',
          subject: 'Third mailbox result',
          senderName: 'Carol Example',
          senderEmailAddress: 'carol@example.com',
          sortDate: new Date().toISOString(),
          fileName: 'mailbox-a.pst',
          scopePath: 'Case Alpha/Search One',
          sourceType: 'mailbox',
          isMailLike: true,
          review: {
            flagged: false,
            tags: [],
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }
        } satisfies MessageSummary
      ],
      total: 3,
      page: 1,
      pageSize: 50,
      totalPages: 1,
      query: '',
      mailOnly: false,
      sort: 'date-desc'
    }

    vi.spyOn(api.auth, 'me').mockResolvedValueOnce(authenticatedStatus)
    vi.spyOn(api.pst, 'catalog').mockResolvedValueOnce(catalog)
    vi.spyOn(api.hiddenFilters, 'list').mockResolvedValue({ items: [] })
    vi.spyOn(api.pst, 'open').mockResolvedValue({
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
    })
    vi.spyOn(api.session, 'folderMessages').mockResolvedValue({
      sessionId: 'session-a',
      page: folderPage
    })
    const messageDetailMock = vi.spyOn(api.session, 'messageDetail').mockImplementation(async (_sessionId, messageId) => {
      const request = detailRequests[messageId]
      if (!request) {
        throw new Error(`Unexpected mailbox message detail request: ${messageId}`)
      }
      return request.promise
    })

    window.localStorage.setItem('pst-mail-explorer.casePath::admin', 'Case Alpha')
    window.localStorage.setItem('pst-mail-explorer.scopePath::admin', 'Case Alpha/Search One')
    window.localStorage.setItem('pst-mail-explorer.pstFileName::admin', 'mailbox-a.pst')
    window.localStorage.setItem('pst-mail-explorer.folderId::admin', 'root-a')
    window.localStorage.setItem('pst-mail-explorer.messageId::admin', 'message-1')
    window.localStorage.setItem('pst-mail-explorer.sourceType::admin', 'mailbox')

    render(
      <div style={{ width: 1800, height: 1600 }}>
        <App />
      </div>
    )

    expect(await screen.findByRole('button', { name: /First mailbox result/ })).toBeInTheDocument()
    await waitFor(() => expect(messageDetailMock).toHaveBeenCalledTimes(3))
    expect(screen.getByText('Loading preview...')).toBeInTheDocument()
    expect(screen.queryByText('No message selected')).not.toBeInTheDocument()

    firstDetail.resolve({ sessionId: 'session-a', detail: firstPreview })
    expect(await screen.findByText('First mailbox body')).toBeInTheDocument()

    const secondItem = await screen.findByRole('button', { name: /Second mailbox result/ })
    await user.click(secondItem)
    expect(messageDetailMock).toHaveBeenCalledTimes(3)
    expect(screen.getByText('Loading preview...')).toBeInTheDocument()
    expect(screen.queryByText('No message selected')).not.toBeInTheDocument()

    secondDetail.resolve({ sessionId: 'session-a', detail: secondPreview })
    expect(await screen.findByText('Second mailbox body')).toBeInTheDocument()
    expect(messageDetailMock).toHaveBeenCalledTimes(3)

    thirdDetail.resolve({ sessionId: 'session-a', detail: thirdPreview })
  })

  it('shows archive office document previews through the preview url', () => {
    render(
      <div style={{ width: 1200, height: 900 }}>
        <EmailPreview
          detail={{
            ...sampleDetail,
            subject: 'Quarterly report',
            bodyText: '',
            archivePath: 'Case1/Search1/Items.1.001.BONUS_AND_COMMISSION_DECISION_MAKING.zip',
            archiveEntryPath: 'SharePoint/Docs/report.docx',
            archiveEntryChain: ['SharePoint/Docs/report.zip', 'SharePoint/Docs/report.docx'],
            archiveEntryName: 'report.docx',
            contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            downloadFilename: 'report.docx',
            previewKind: 'text',
            previewText: 'Quarterly report',
            previewHtml: '',
            previewUrl: '/api/items/archive-1/preview',
            downloadUrl: '/api/items/archive-1/content'
          }}
          theme="light"
          onDownloadJson={vi.fn()}
          onDownloadEml={vi.fn()}
          onDownloadItem={vi.fn()}
          onToggleFlag={vi.fn()}
          onClearReview={vi.fn()}
          onOpenTags={vi.fn()}
          onOpenFullView={vi.fn()}
          tagCount={0}
          onOpenAttachment={vi.fn()}
          onOpenPrev={vi.fn()}
          onOpenNext={vi.fn()}
          canNavigatePrev={false}
          canNavigateNext={false}
        />
      </div>
    )

    expect(screen.getByText('Document preview')).toBeInTheDocument()
    expect(screen.getByText('Teams or document preview')).toBeInTheDocument()
    expect(screen.getByTitle('Document preview')).toHaveAttribute('src', '/api/items/archive-1/preview')
    expect(screen.queryByText('No inline preview is available for this file type.')).not.toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Clear flag' })).toBeInTheDocument()
    expect(screen.getAllByRole('button', { name: 'Download file' })).toHaveLength(1)
  })

  it('opens a search result from a different mailbox in the reading pane', async () => {
    const user = userEvent.setup()
    const authenticatedStatus: AuthStatus = {
      authenticated: true,
      enabled: true,
      canManageUsers: true,
      mfaEnabled: true,
      mfaEnforced: false,
      lockedUntil: null,
      loginFailedCount: 0,
      passwordResetAvailable: false,
      user: { username: 'admin', assignedCasePaths: [] },
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
          id: 'search-hit-current',
          messageId: 'search-hit-current',
          subject: 'Current mailbox search result',
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
        },
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
      if (sessionId === 'session-a' && messageId === 'search-hit-current') {
        return { sessionId, detail: currentDetail }
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

    expect(await screen.findByText('Current mailbox body')).toBeInTheDocument()
    const callsBeforeSecondSelection = vi.mocked(api.search).mock.calls.length
    await user.click(screen.getByRole('button', { name: /Target search result/ }))

    expect(await screen.findByText('Target mailbox body')).toBeInTheDocument()
    expect(api.pst.open).toHaveBeenCalledWith('Case Beta/Search Two', 'mailbox-b.pst')
    await waitFor(() => {
      expect(vi.mocked(api.search).mock.calls.length).toBe(callsBeforeSecondSelection)
    })
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
