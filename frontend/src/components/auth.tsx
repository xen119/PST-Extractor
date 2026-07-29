import * as React from 'react'
import { Download, KeyRound, LogIn, ScanLine, ShieldAlert, ShieldCheck } from 'lucide-react'
import type {
  MfaEnrollmentStartResponse,
  PasswordResetLookupResponse,
  UserInvite
} from '@/types'
import { Button, Dialog, DialogContent, DialogDescription, DialogTitle, Input } from '@/components/ui'

export type AuthView = 'login' | 'mfa' | 'invite' | 'reset' | 'change'
export type InviteStep = 'password' | 'prompt' | 'setup' | 'complete'

export interface AuthScreenProps {
  view: AuthView
  busy: boolean
  message: string
  error: string
  entraEnabled: boolean
  passwordResetAvailable: boolean
  invite: UserInvite | null
  inviteStep: InviteStep
  inviteMfaAvailable: boolean
  inviteMfaEnforced: boolean
  inviteSetup: MfaEnrollmentStartResponse | null
  inviteRecoveryCodes: string[]
  resetLookup: PasswordResetLookupResponse | null
  passwordChangeUser?: string | null
  onLogin: (username: string, password: string) => void
  onMicrosoftSignIn: () => void
  onMfaChallenge: (code: string) => void
  onPasswordResetRequest: (usernameOrEmail: string) => Promise<void>
  onPasswordResetConfirm: (password: string, confirmPassword: string) => Promise<void>
  onInviteAccept: (password: string) => void
  onInviteMfaStart: () => void
  onInviteMfaSkip: () => void
  onInviteMfaSubmit: (code: string) => void
  onInviteFinish: () => void
  onOpenLogin: () => void
}

export function AuthScreen({
  view,
  busy,
  message,
  error,
  entraEnabled,
  passwordResetAvailable,
  invite,
  inviteStep,
  inviteMfaAvailable,
  inviteMfaEnforced,
  inviteSetup,
  inviteRecoveryCodes,
  resetLookup,
  passwordChangeUser,
  onLogin,
  onMicrosoftSignIn,
  onMfaChallenge,
  onPasswordResetRequest,
  onPasswordResetConfirm,
  onInviteAccept,
  onInviteMfaStart,
  onInviteMfaSkip,
  onInviteMfaSubmit,
  onInviteFinish,
  onOpenLogin
}: AuthScreenProps) {
  const [username, setUsername] = React.useState('')
  const [password, setPassword] = React.useState('')
  const [mfaCode, setMfaCode] = React.useState('')
  const [invitePassword, setInvitePassword] = React.useState('')
  const [inviteConfirmPassword, setInviteConfirmPassword] = React.useState('')
  const [inviteMfaCode, setInviteMfaCode] = React.useState('')
  const [showPasswordResetRequest, setShowPasswordResetRequest] = React.useState(false)
  const [passwordResetUsernameOrEmail, setPasswordResetUsernameOrEmail] = React.useState('')
  const [resetPassword, setResetPassword] = React.useState('')
  const [resetConfirmPassword, setResetConfirmPassword] = React.useState('')

  React.useEffect(() => {
    if (view === 'login') {
      setUsername('')
      setPassword('')
      setMfaCode('')
      setShowPasswordResetRequest(false)
      setPasswordResetUsernameOrEmail('')
    }
    if (view === 'reset') {
      setResetPassword('')
      setResetConfirmPassword('')
    }
    if (view === 'change') {
      setResetPassword('')
      setResetConfirmPassword('')
    }
  }, [view])

  const passwordResetDialog = (
    <Dialog
      open={showPasswordResetRequest}
      onOpenChange={(open) => {
        if (!open) {
          setShowPasswordResetRequest(false)
        }
      }}
    >
      <DialogContent className="w-[min(96vw,560px)]">
        <div className="space-y-4 p-6">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">
              <KeyRound className="h-6 w-6" />
            </div>
            <div>
              <DialogTitle className="text-2xl">Reset password</DialogTitle>
              <DialogDescription>
                Enter your username or email address and we will send a reset link if the account is eligible.
              </DialogDescription>
            </div>
          </div>
          {message ? (
            <div className="rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">
              {message}
            </div>
          ) : null}
          {error ? (
            <div className="rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
              {error}
            </div>
          ) : null}
          <form
            className="space-y-4"
            onSubmit={async (event) => {
              event.preventDefault()
              if (!passwordResetUsernameOrEmail.trim()) {
                return
              }
              try {
                await onPasswordResetRequest(passwordResetUsernameOrEmail.trim())
                setShowPasswordResetRequest(false)
                setPasswordResetUsernameOrEmail('')
              } catch {
                // Surface the error via the parent auth state and keep the dialog open.
              }
            }}
          >
            <label className="block text-sm font-medium text-[color:var(--text)]">
              Username or email
              <Input
                className="mt-2"
                value={passwordResetUsernameOrEmail}
                onChange={(event) => setPasswordResetUsernameOrEmail(event.target.value)}
                autoComplete="username"
                autoFocus
              />
            </label>
            <div className="flex flex-wrap justify-end gap-2">
              <Button
                type="button"
                variant="ghost"
                onClick={() => {
                  setShowPasswordResetRequest(false)
                  setPasswordResetUsernameOrEmail('')
                }}
              >
                Cancel
              </Button>
              <Button type="submit" disabled={busy || !passwordResetUsernameOrEmail.trim()}>
                {busy ? 'Sending...' : 'Send reset link'}
              </Button>
            </div>
          </form>
        </div>
      </DialogContent>
    </Dialog>
  )

  if (view === 'login') {
    return (
      <section className="fixed inset-0 z-20 grid place-items-center overflow-auto px-4 py-8">
        <div className="mx-auto w-full max-w-[560px] panel-surface-strong overflow-hidden">
          <div className="p-8 sm:p-10">
            <div className="flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] text-[color:var(--accent)]">
                <ShieldCheck className="h-7 w-7" aria-hidden="true" />
              </div>
              <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                PST Mail Explorer
              </div>
            </div>

            <div className="mt-8 space-y-4">
              {message ? (
                <div className="rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">
                  {message}
                </div>
              ) : null}

              {error ? (
                <div className="rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
                  {error}
                </div>
              ) : null}

              <h2 className="text-2xl font-semibold tracking-tight text-[color:var(--text)]">Sign in to continue</h2>

              <form
                className="space-y-4"
                onSubmit={(event) => {
                  event.preventDefault()
                  if (!username.trim() || !password) {
                    return
                  }
                  onLogin(username.trim(), password)
                }}
              >
                <label className="block text-sm font-medium text-[color:var(--text)]">
                  Username
                  <Input
                    className="mt-2"
                    value={username}
                    onChange={(event) => setUsername(event.target.value)}
                    autoComplete="username"
                    placeholder="Username or email"
                    autoFocus
                  />
                </label>
                <label className="block text-sm font-medium text-[color:var(--text)]">
                  Password
                  <Input
                    className="mt-2"
                    type="password"
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    autoComplete="current-password"
                  />
                </label>
                <Button type="submit" className="w-full justify-center" disabled={busy}>
                  {busy ? 'Signing in...' : 'Sign in'}
                </Button>
              </form>

              {entraEnabled ? (
                <Button
                  type="button"
                  variant="secondary"
                  className="w-full justify-center gap-2"
                  disabled={busy}
                  onClick={onMicrosoftSignIn}
                >
                  <LogIn className="h-4 w-4" />
                  Sign in with Microsoft
                </Button>
              ) : null}

              {passwordResetAvailable ? (
                <div className="pt-1">
                  <Button
                    type="button"
                    variant="ghost"
                    className="px-0 text-sm"
                    disabled={busy}
                    onClick={() => {
                      setPasswordResetUsernameOrEmail(username.trim())
                      setShowPasswordResetRequest(true)
                    }}
                  >
                    Forgot password?
                  </Button>
                </div>
              ) : null}

              <p className="pt-2 text-xs leading-6 text-[color:var(--muted)]">
                Developed by DigiVectra DevOps
                <br />
                For access or investigation support, contact the Compliance Engineering team.
              </p>
            </div>
          </div>
          {passwordResetDialog}
        </div>
      </section>
    )
  }

  if (view === 'reset' || view === 'change') {
    const resetAccount = resetLookup?.reset || null
    const isPasswordChange = view === 'change'
    return (
      <section className="fixed inset-0 z-20 grid place-items-center overflow-auto px-4 py-8">
        <div className="mx-auto w-full max-w-[560px] panel-surface-strong overflow-hidden">
          <div className="p-8 sm:p-10">
            <div className="flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] text-[color:var(--accent)]">
                <KeyRound className="h-7 w-7" aria-hidden="true" />
              </div>
              <div className="min-w-0">
                <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                  {isPasswordChange ? 'Password change required' : 'Password reset'}
                </div>
                <h1 className="mt-1 text-3xl font-semibold tracking-tight text-[color:var(--text)]">Set a new password</h1>
                <p className="mt-2 max-w-xl text-sm leading-6 text-[color:var(--muted)]">
                  {isPasswordChange
                    ? 'Your administrator provided a temporary password. Change it before continuing.'
                    : 'Use the link from your email to replace the current password and regain access.'}
                </p>
              </div>
            </div>

            {message ? (
              <div className="mt-6 rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">
                {message}
              </div>
            ) : null}

            {error ? (
              <div className="mt-4 rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {error}
              </div>
            ) : null}

            {isPasswordChange ? (
              <>
                <p className="mt-6 text-xs leading-6 text-[color:var(--muted)]">
                  Account: <span className="font-medium text-[color:var(--text)]">{passwordChangeUser || 'Temporary access'}</span>
                </p>
                <div className="mt-6 space-y-5">
                  <form
                    className="space-y-4"
                    onSubmit={(event) => {
                      event.preventDefault()
                      if (!resetPassword || resetPassword !== resetConfirmPassword) {
                        return
                      }
                      void onPasswordResetConfirm(resetPassword, resetConfirmPassword)
                    }}
                  >
                    <label className="block text-sm font-medium text-[color:var(--text)]">
                      New password
                      <Input
                        className="mt-2"
                        type="password"
                        value={resetPassword}
                        onChange={(event) => setResetPassword(event.target.value)}
                        autoComplete="new-password"
                        autoFocus
                      />
                    </label>
                    <label className="block text-sm font-medium text-[color:var(--text)]">
                      Confirm password
                      <Input
                        className="mt-2"
                        type="password"
                        value={resetConfirmPassword}
                        onChange={(event) => setResetConfirmPassword(event.target.value)}
                        autoComplete="new-password"
                      />
                    </label>
                    <Button type="submit" className="w-full justify-center" disabled={busy}>
                      {busy ? 'Updating...' : 'Update password'}
                    </Button>
                    <Button type="button" variant="ghost" className="w-full justify-center" onClick={onOpenLogin}>
                      Back to sign in
                    </Button>
                </form>
              </div>
            </>
            ) : resetAccount ? (
              <>
                <p className="mt-6 text-xs leading-6 text-[color:var(--muted)]">
                  Account: <span className="font-medium text-[color:var(--text)]">{resetAccount.username}</span>
                  {resetAccount.recipientEmail ? (
                    <>
                      {' '}
                      · <span className="font-medium text-[color:var(--text)]">{resetAccount.recipientEmail}</span>
                    </>
                  ) : null}
                </p>

                <div className="mt-6 space-y-5">
                  <form
                    className="space-y-4"
                    onSubmit={(event) => {
                      event.preventDefault()
                      if (!resetPassword || resetPassword !== resetConfirmPassword) {
                        return
                      }
                      void onPasswordResetConfirm(resetPassword, resetConfirmPassword)
                    }}
                  >
                    <label className="block text-sm font-medium text-[color:var(--text)]">
                      New password
                      <Input
                        className="mt-2"
                        type="password"
                        value={resetPassword}
                        onChange={(event) => setResetPassword(event.target.value)}
                        autoComplete="new-password"
                        autoFocus
                      />
                    </label>
                    <label className="block text-sm font-medium text-[color:var(--text)]">
                      Confirm password
                      <Input
                        className="mt-2"
                        type="password"
                        value={resetConfirmPassword}
                        onChange={(event) => setResetConfirmPassword(event.target.value)}
                        autoComplete="new-password"
                      />
                    </label>
                    <Button type="submit" className="w-full justify-center" disabled={busy}>
                      {busy ? 'Updating...' : 'Update password'}
                    </Button>
                    <Button type="button" variant="ghost" className="w-full justify-center" onClick={onOpenLogin}>
                      Back to sign in
                    </Button>
                  </form>
                </div>
              </>
            ) : (
              <div className="mt-6 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4">
                <div className="text-sm font-medium text-[color:var(--text)]">This reset link is not valid anymore.</div>
                <p className="mt-2 text-sm leading-6 text-[color:var(--muted)]">
                  Request a new reset link from the sign-in screen, then open the latest email.
                </p>
                <div className="mt-4 flex flex-wrap gap-2">
                  <Button onClick={onOpenLogin}>Back to sign in</Button>
                </div>
              </div>
            )}
          </div>
        </div>
      </section>
    )
  }

  if (view === 'invite') {
    const inviteMfaOptional = inviteMfaAvailable && !inviteMfaEnforced
    return (
      <section className="fixed inset-0 z-20 grid place-items-center overflow-auto px-4 py-8">
        <div className="mx-auto w-full max-w-[560px] panel-surface-strong overflow-hidden">
          <div className="p-8 sm:p-10">
            <div className="flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] text-[color:var(--accent)]">
                <ShieldCheck className="h-7 w-7" aria-hidden="true" />
              </div>
              <div className="min-w-0">
                <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                  Invite onboarding
                </div>
                <h1 className="mt-1 text-3xl font-semibold tracking-tight text-[color:var(--text)]">
                  Set your password
                </h1>
                <p className="mt-2 max-w-xl text-sm leading-6 text-[color:var(--muted)]">
                  Use this invite to activate your account, then optionally enroll MFA.
                </p>
              </div>
            </div>

            {message ? (
              <div className="mt-6 rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">
                {message}
              </div>
            ) : null}

            {error ? (
              <div className="mt-4 rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {error}
              </div>
            ) : null}

            {invite ? (
              <p className="mt-6 text-xs leading-6 text-[color:var(--muted)]">
                Account: <span className="font-medium text-[color:var(--text)]">{invite.username}</span>
                {invite.recipientEmail ? (
                  <>
                    {' '}
                    · <span className="font-medium text-[color:var(--text)]">{invite.recipientEmail}</span>
                  </>
                ) : null}
              </p>
            ) : null}

            <div className="mt-6 space-y-5">
              {inviteStep === 'password' ? (
                <form
                  className="space-y-4"
                  onSubmit={(event) => {
                    event.preventDefault()
                    if (!invitePassword || invitePassword !== inviteConfirmPassword) {
                      return
                    }
                    onInviteAccept(invitePassword)
                  }}
                >
                  <label className="block text-sm font-medium text-[color:var(--text)]">
                    Password
                    <Input
                      className="mt-2"
                      type="password"
                      value={invitePassword}
                      onChange={(event) => setInvitePassword(event.target.value)}
                      autoComplete="new-password"
                      autoFocus
                    />
                  </label>
                  <label className="block text-sm font-medium text-[color:var(--text)]">
                    Confirm password
                    <Input
                      className="mt-2"
                      type="password"
                      value={inviteConfirmPassword}
                      onChange={(event) => setInviteConfirmPassword(event.target.value)}
                      autoComplete="new-password"
                    />
                  </label>
                  <Button type="submit" className="w-full justify-center" disabled={busy}>
                    {busy ? 'Saving...' : 'Set password'}
                  </Button>
                </form>
              ) : null}

              {inviteStep === 'prompt' ? (
                <div className="rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-4">
                  <div className="flex items-center gap-3">
                    <ShieldAlert className="h-5 w-5 text-[color:var(--warning)]" />
                    <div className="text-sm font-medium text-[color:var(--text)]">Set up MFA</div>
                  </div>
                  <p className="mt-2 text-sm text-[color:var(--muted)]">
                    Optional, but recommended. You can scan a QR code or enter the setup key manually.
                  </p>
                  <div className="mt-4 flex flex-wrap gap-2">
                    <Button onClick={onInviteMfaStart}>Set up MFA</Button>
                    <Button variant="ghost" onClick={onInviteMfaSkip}>Skip for now</Button>
                  </div>
                </div>
              ) : null}

              {inviteStep === 'setup' ? (
                <div className="space-y-4 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-4">
                  <div className="flex items-center gap-3">
                    <ScanLine className="h-5 w-5 text-[color:var(--accent)]" />
                    <div>
                      <div className="text-sm font-semibold text-[color:var(--text)]">Scan the QR code</div>
                      <div className="text-sm text-[color:var(--muted)]">Use an authenticator app or enter the setup key manually.</div>
                    </div>
                  </div>
                  {inviteSetup ? (
                    <div className="grid gap-4 lg:grid-cols-[1fr_1.1fr]">
                      <div className="rounded-2xl border border-[color:var(--line)] bg-white p-3">
                        <img className="w-full rounded-xl" src={inviteSetup.qrCodeDataUrl} alt="MFA setup QR code" />
                      </div>
                      <div className="space-y-3">
                        <div className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-3">
                          <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">Manual setup key</div>
                          <code className="mt-2 block break-all rounded-xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-3 text-sm text-[color:var(--text)]">{inviteSetup.secret}</code>
                          <div className="mt-2 break-all text-xs text-[color:var(--muted)]">{inviteSetup.otpauthUri}</div>
                        </div>
                        <form
                          className="space-y-3"
                          onSubmit={(event) => {
                            event.preventDefault()
                            if (!inviteMfaCode.trim()) {
                              return
                            }
                            onInviteMfaSubmit(inviteMfaCode.trim())
                          }}
                        >
                          <label className="block text-sm font-medium text-[color:var(--text)]">
                            Verification code
                            <Input
                              className="mt-2"
                              value={inviteMfaCode}
                              onChange={(event) => setInviteMfaCode(event.target.value)}
                              autoComplete="one-time-code"
                              inputMode="numeric"
                              autoFocus
                            />
                          </label>
                          <Button type="submit" className="w-full justify-center" disabled={busy}>
                            {busy ? 'Verifying...' : 'Verify code'}
                          </Button>
                        </form>
                      </div>
                    </div>
                  ) : (
                    <div className="empty-state">Loading MFA setup...</div>
                  )}
                </div>
              ) : null}

              {inviteStep === 'complete' ? (
                <div className="space-y-4 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-4">
                  <div className="flex items-center gap-3">
                    <Download className="h-5 w-5 text-[color:var(--accent)]" />
                    <div>
                      <div className="text-sm font-semibold text-[color:var(--text)]">Save your recovery codes</div>
                      <div className="text-sm text-[color:var(--muted)]">Keep these codes somewhere safe. They can be used if you lose access to your authenticator app.</div>
                    </div>
                  </div>
                  <div className="grid max-h-72 gap-2 overflow-auto rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-3 text-sm">
                    {inviteRecoveryCodes.length ? (
                      inviteRecoveryCodes.map((code) => (
                        <code key={code} className="rounded-xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] px-3 py-2">
                          {code}
                        </code>
                      ))
                    ) : (
                      <div className="empty-state">No recovery codes available.</div>
                    )}
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <Button
                      variant="ghost"
                      onClick={() => {
                        if (!inviteRecoveryCodes.length) {
                          return
                        }
                        const contents = inviteRecoveryCodes.join('\n')
                        const blob = new Blob([contents], { type: 'text/plain;charset=utf-8' })
                        const url = URL.createObjectURL(blob)
                        const a = document.createElement('a')
                        a.href = url
                        a.download = 'mfa-recovery-codes.txt'
                        a.click()
                        URL.revokeObjectURL(url)
                      }}
                    >
                      Download recovery codes
                    </Button>
                    <Button onClick={onInviteFinish}>Continue to platform</Button>
                  </div>
                </div>
              ) : null}
            </div>
          </div>
        </div>
      </section>
    )
  }

  if (view === 'mfa') {
    return (
      <section className="fixed inset-0 z-20 grid place-items-center overflow-auto px-4 py-8">
        <div className="mx-auto w-full max-w-[420px] panel-surface-strong overflow-hidden">
          <div className="p-8 sm:p-10">
            <div className="flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] text-[color:var(--accent)]">
                <ShieldCheck className="h-7 w-7" aria-hidden="true" />
              </div>
              <div className="min-w-0">
                <h1 className="text-3xl font-semibold tracking-tight text-[color:var(--text)]">
                  Verify your sign in
                </h1>
              </div>
            </div>

            {error ? (
              <div className="mt-6 rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
                {error}
              </div>
            ) : null}

            <form
              className="mt-6 space-y-4"
              onSubmit={(event) => {
                event.preventDefault()
                if (!mfaCode.trim()) {
                  return
                }
                onMfaChallenge(mfaCode.trim())
              }}
            >
              <label className="block text-sm font-medium text-[color:var(--text)]">
                Verification code
                <Input
                  className="mt-2"
                  value={mfaCode}
                  onChange={(event) => setMfaCode(event.target.value)}
                  autoComplete="one-time-code"
                  inputMode="numeric"
                  autoFocus
                />
              </label>
              <Button type="submit" className="w-full justify-center" disabled={busy}>
                {busy ? 'Verifying...' : 'Verify code'}
              </Button>
              <Button type="button" variant="ghost" className="w-full justify-center" onClick={onOpenLogin}>
                Back to sign in
              </Button>
            </form>
          </div>
        </div>
      </section>
    )
  }

  return null
}

export function MfaReminderDialog({
  open,
  username,
  allowSkip,
  onSetup,
  onSkip
}: {
  open: boolean
  username: string
  allowSkip: boolean
  onSetup: () => void
  onSkip: () => void
}) {
  return (
    <Dialog open={open}>
      <DialogContent
        showCloseButton={false}
        className="w-[min(96vw,560px)]"
        onPointerDownOutside={(event) => event.preventDefault()}
        onEscapeKeyDown={(event) => event.preventDefault()}
      >
        <div className="space-y-4 p-6">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[color:var(--warning-bg)] text-[color:var(--warning)]">
              <ShieldAlert className="h-6 w-6" />
            </div>
            <div>
              <DialogTitle className="text-2xl">Set up MFA</DialogTitle>
              <DialogDescription>Signed in as {username || 'your account'}.</DialogDescription>
            </div>
          </div>
          <p className="text-sm leading-6 text-[color:var(--muted)]">
            {allowSkip
              ? 'This account does not have multi-factor authentication enabled. Add MFA to reduce the risk of unauthorized access.'
              : 'This account requires multi-factor authentication before you can continue.'}
          </p>
          <div className="flex flex-wrap justify-end gap-2">
            <Button onClick={onSetup}>Set up MFA</Button>
            {allowSkip ? <Button variant="ghost" onClick={onSkip}>Skip for now</Button> : null}
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}

export function MfaSetupDialog({
  open,
  loading,
  message,
  error,
  data,
  recoveryCodes,
  onSubmit,
  onClose,
  onDownload,
  onFinish
}: {
  open: boolean
  loading: boolean
  message: string
  error: string
  data: MfaEnrollmentStartResponse | null
  recoveryCodes: string[]
  onSubmit: (code: string) => void
  onClose: () => void
  onDownload: () => void
  onFinish: () => void
}) {
  const [code, setCode] = React.useState('')

  React.useEffect(() => {
    if (!open) {
      setCode('')
    }
  }, [open])

  return (
    <Dialog open={open}>
      <DialogContent className="w-[min(96vw,1040px)]">
        <div className="space-y-4 p-6">
          <div className="flex items-center gap-3">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">
              <ScanLine className="h-6 w-6" />
            </div>
            <div>
              <DialogTitle className="text-2xl">Set up MFA</DialogTitle>
              <DialogDescription>Scan the QR code or enter the setup key manually.</DialogDescription>
            </div>
          </div>
          {message ? (
            <div className="rounded-2xl border border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] px-4 py-3 text-sm text-[color:var(--accent-strong)]">
              {message}
            </div>
          ) : null}
          {error ? (
            <div className="rounded-2xl border border-[color:rgba(220,38,38,0.2)] bg-[color:rgba(220,38,38,0.08)] px-4 py-3 text-sm text-[color:var(--danger)]">
              {error}
            </div>
          ) : null}
          {data ? (
            <div className="grid gap-4 lg:grid-cols-[1fr_1.1fr]">
              <div className="rounded-2xl border border-[color:var(--line)] bg-white p-3">
                <img className="w-full rounded-xl" src={data.qrCodeDataUrl} alt="MFA setup QR code" />
              </div>
              <div className="space-y-3">
                <div className="rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-3">
                  <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[color:var(--muted)]">Manual setup key</div>
                  <code className="mt-2 block break-all rounded-xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-3 text-sm text-[color:var(--text)]">
                    {data.secret}
                  </code>
                  <div className="mt-2 break-all text-xs text-[color:var(--muted)]">{data.otpauthUri}</div>
                </div>
                <form
                  className="space-y-3"
                  onSubmit={(event) => {
                    event.preventDefault()
                    if (!code.trim()) {
                      return
                    }
                    onSubmit(code.trim())
                  }}
                >
                  <label className="block text-sm font-medium text-[color:var(--text)]">
                    Verification code
              <Input
                className="mt-2"
                value={code}
                onChange={(event) => setCode(event.target.value)}
                autoComplete="one-time-code"
                inputMode="numeric"
                autoFocus
              />
                  </label>
                  <Button type="submit" className="w-full justify-center" disabled={loading}>
                    {loading ? 'Verifying...' : 'Verify code'}
                  </Button>
                </form>
              </div>
            </div>
          ) : (
            <div className="empty-state">Loading MFA setup...</div>
          )}

          {recoveryCodes.length ? (
            <div className="space-y-3 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-4">
              <div className="flex items-center gap-3">
                <Download className="h-5 w-5 text-[color:var(--accent)]" />
                <div>
                  <div className="text-sm font-semibold text-[color:var(--text)]">Recovery codes</div>
                  <div className="text-sm text-[color:var(--muted)]">Download them once and store them securely.</div>
                </div>
              </div>
              <div className="grid max-h-72 gap-2 overflow-auto rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-soft)] p-3 text-sm">
                {recoveryCodes.map((item) => (
                  <code key={item} className="rounded-xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] px-3 py-2">
                    {item}
                  </code>
                ))}
              </div>
              <div className="flex flex-wrap justify-end gap-2">
                <Button variant="ghost" onClick={onDownload}>Download recovery codes</Button>
                <Button onClick={onFinish}>Continue to platform</Button>
              </div>
            </div>
          ) : null}
          <div className="flex justify-end">
            <Button variant="ghost" onClick={onClose}>Close</Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  )
}
