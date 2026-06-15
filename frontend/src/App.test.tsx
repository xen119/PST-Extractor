import * as React from 'react'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { describe, expect, it, vi } from 'vitest'
import { AppShell, EmailPreview, MessageList, Sidebar, TagManagerDialog } from '@/components/layout'
import { AuthScreen, MfaReminderDialog } from '@/components/auth'
import { getCasePathFromScopePath } from '@/App'
import type { AttachmentDetail, FolderNode, MessageDetail, MessageSummary } from '@/types'

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
      />
    )

    expect(screen.getByText('Set your password')).toBeInTheDocument()
    expect(screen.getByLabelText('Password')).toBeInTheDocument()
    expect(screen.getByLabelText('Confirm password')).toBeInTheDocument()
  })

  it('shows a blocking MFA reminder without a close button', () => {
    render(<MfaReminderDialog open username="admin" onSetup={vi.fn()} onSkip={vi.fn()} />)

    expect(screen.getByRole('heading', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Set up MFA' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Skip for now' })).toBeInTheDocument()
    expect(screen.queryByRole('button', { name: 'Close' })).not.toBeInTheDocument()
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
          isRemovedCatalog={false}
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
          onCatalogModeToggle={vi.fn()}
          onRefreshCatalog={vi.fn()}
          onOpenMailbox={vi.fn()}
          onRemoveMailbox={vi.fn()}
          onRestoreMailbox={vi.fn()}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    expect(screen.getByRole('button', { name: 'Show removed PSTs' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'Case Alpha' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'Search One' })).toBeInTheDocument()
    expect(screen.queryByText('Case Alpha (3)')).not.toBeInTheDocument()
    expect(screen.queryByText('Search One (2)')).not.toBeInTheDocument()
    expect(screen.getByRole('combobox', { name: 'Mailbox selector' })).toBeInTheDocument()
    expect(screen.getByRole('option', { name: 'mailbox-a.pst' })).toBeInTheDocument()
    expect(screen.queryByText('Found 2 mailbox files in Case Alpha / Search One.')).not.toBeInTheDocument()
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
          isRemovedCatalog={false}
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
          onCatalogModeToggle={vi.fn()}
          onRefreshCatalog={vi.fn()}
          onOpenMailbox={vi.fn()}
          onRemoveMailbox={vi.fn()}
          onRestoreMailbox={vi.fn()}
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
  })

  it('calls the mailbox remove and restore actions', async () => {
    const user = userEvent.setup()
    const onRemoveMailbox = vi.fn()
    const onRestoreMailbox = vi.fn()

    const { rerender } = render(
      <div style={{ width: 420, height: 900 }}>
        <Sidebar
          isRemovedCatalog={false}
          caseOptions={[{ label: 'Case Alpha', value: 'Case Alpha', count: 1 }]}
          selectedCasePath="Case Alpha"
          selectedScopePath="Case Alpha/Search One"
          searchOptions={[{ label: 'Search One', value: 'Case Alpha/Search One', count: 1 }]}
          catalogFiles={[
            {
              fileName: 'mailbox.pst',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One'
            }
          ]}
          selectedPstFileName="mailbox.pst"
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          onCatalogModeToggle={vi.fn()}
          onRefreshCatalog={vi.fn()}
          onOpenMailbox={vi.fn()}
          onRemoveMailbox={onRemoveMailbox}
          onRestoreMailbox={onRestoreMailbox}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    await user.click(screen.getByRole('button', { name: 'Remove PST' }))
    expect(onRemoveMailbox).toHaveBeenCalledWith('mailbox.pst', 'Case Alpha/Search One')

    rerender(
      <div style={{ width: 420, height: 900 }}>
        <Sidebar
          isRemovedCatalog
          caseOptions={[{ label: 'Case Alpha', value: 'Case Alpha', count: 1 }]}
          selectedCasePath="Case Alpha"
          selectedScopePath="Case Alpha/Search One"
          searchOptions={[{ label: 'Search One', value: 'Case Alpha/Search One', count: 1 }]}
          catalogFiles={[
            {
              fileName: 'mailbox.pst',
              size: 1024,
              modifiedAt: new Date().toISOString(),
              scopePath: 'Case Alpha/Search One'
            }
          ]}
          selectedPstFileName="mailbox.pst"
          onCaseChange={vi.fn()}
          onScopeChange={vi.fn()}
          onCatalogModeToggle={vi.fn()}
          onRefreshCatalog={vi.fn()}
          onOpenMailbox={vi.fn()}
          onRemoveMailbox={onRemoveMailbox}
          onRestoreMailbox={onRestoreMailbox}
          folderTree={null}
          currentFolderId=""
          onSelectFolder={vi.fn()}
        />
      </div>
    )

    await user.click(screen.getByRole('button', { name: 'Restore PST' }))
    expect(onRestoreMailbox).toHaveBeenCalledWith('mailbox.pst', 'Case Alpha/Search One')
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
          hiddenFiltersOpen={false}
          hiddenRules={[]}
          hiddenFiltersCount={0}
          onQueryChange={vi.fn()}
          onSearch={vi.fn()}
          onSearchScopeChange={vi.fn()}
          onMailOnlyChange={vi.fn()}
          onSortChange={vi.fn()}
          onReviewFlaggedChange={vi.fn()}
          onReviewTaggedChange={vi.fn()}
          onToggleHiddenFilters={vi.fn()}
          onRemoveHiddenFilter={vi.fn()}
          onSelectMessage={vi.fn()}
          onPrevPage={vi.fn()}
          onNextPage={vi.fn()}
          onOpenBundle={vi.fn()}
          onRefreshSearchIndex={vi.fn()}
          onCreateHiddenFilter={vi.fn()}
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
