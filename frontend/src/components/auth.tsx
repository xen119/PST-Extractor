import * as React from 'react'
import { ShieldCheck, ShieldAlert, ScanLine, Download, LockKeyhole, BadgeInfo } from 'lucide-react'
import type {
  MfaEnrollmentStartResponse,
  UserInvite
} from '@/types'
import { Button, Badge, Dialog, DialogContent, DialogDescription, DialogTitle, Input } from '@/components/ui'

export type AuthView = 'login' | 'mfa' | 'invite'
export type InviteStep = 'password' | 'prompt' | 'setup' | 'complete'

export interface AuthScreenProps {
  view: AuthView
  busy: boolean
  message: string
  error: string
  invite: UserInvite | null
  inviteStep: InviteStep
  inviteMfaAvailable: boolean
  inviteSetup: MfaEnrollmentStartResponse | null
  inviteRecoveryCodes: string[]
  onLogin: (username: string, password: string) => void
  onMfaChallenge: (code: string) => void
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
  invite,
  inviteStep,
  inviteMfaAvailable,
  inviteSetup,
  inviteRecoveryCodes,
  onLogin,
  onMfaChallenge,
  onInviteAccept,
  onInviteMfaStart,
  onInviteMfaSkip,
  onInviteMfaSubmit,
  onInviteFinish,
  onOpenLogin
}: AuthScreenProps) {
  const [username, setUsername] = React.useState('admin')
  const [password, setPassword] = React.useState('')
  const [mfaCode, setMfaCode] = React.useState('')
  const [invitePassword, setInvitePassword] = React.useState('')
  const [inviteConfirmPassword, setInviteConfirmPassword] = React.useState('')
  const [inviteMfaCode, setInviteMfaCode] = React.useState('')

  React.useEffect(() => {
    if (view === 'login') {
      setPassword('')
      setMfaCode('')
    }
  }, [view])

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

              <p className="pt-2 text-xs leading-6 text-[color:var(--muted)]">
                Developed by DigiVectra DevOps
                <br />
                For access or investigation support, contact the Compliance Engineering team.
              </p>
            </div>
          </div>
        </div>
      </section>
    )
  }

  return (
    <section className="fixed inset-0 z-20 grid place-items-center overflow-auto px-4 py-8">
      <div className="mx-auto w-full max-w-[1024px] panel-surface-strong overflow-hidden">
        <div className="grid gap-0 lg:grid-cols-[1.1fr_0.9fr]">
          <div className="border-b border-[color:var(--line)] bg-[color:var(--surface-soft)] p-8 lg:border-b-0 lg:border-r">
            <div className="flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-2xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] text-[color:var(--accent)]">
                <ShieldCheck className="h-7 w-7" aria-hidden="true" />
              </div>
              <div className="min-w-0">
                <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                  PST Mail Explorer
                </div>
                <h1 className="mt-1 text-3xl font-semibold tracking-tight text-[color:var(--text)]">
                  Secure PST review
                </h1>
                <p className="mt-2 max-w-xl text-sm leading-6 text-[color:var(--muted)]">
                  A minimal, adjustable workspace for mailbox review, search, tagging, and audit activity.
                </p>
              </div>
            </div>

            <div className="mt-8 space-y-3 rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-5">
              <div className="flex items-center gap-3">
                <Badge className="border-[color:var(--accent-soft)] bg-[color:var(--accent-soft)] text-[color:var(--accent-strong)]">
                  Secure access
                </Badge>
                <Badge className="border-[color:var(--line)]">Local credentials</Badge>
              </div>
              <ul className="grid gap-2 text-sm text-[color:var(--muted)]">
                <li className="flex items-center gap-2">
                  <LockKeyhole className="h-4 w-4 text-[color:var(--accent)]" />
                  Username and password sign-in
                </li>
                <li className="flex items-center gap-2">
                  <ShieldAlert className="h-4 w-4 text-[color:var(--warning)]" />
                  MFA reminder for accounts that have not enrolled yet
                </li>
                <li className="flex items-center gap-2">
                  <BadgeInfo className="h-4 w-4 text-[color:var(--muted)]" />
                  Invite links and recovery codes are supported
                </li>
              </ul>
            </div>

            <p className="mt-6 text-xs leading-6 text-[color:var(--muted)]">
              Developed by DigiVectra DevOps
              <br />
              For access or investigation support, contact the Compliance Engineering team.
            </p>
          </div>

          <div className="p-8">
            <div className="space-y-4">
              <div>
                <div className="text-sm font-semibold tracking-[0.18em] text-[color:var(--muted)] uppercase">
                  {view === 'invite' ? 'Invite onboarding' : view === 'mfa' ? 'Second factor' : 'Sign in'}
                </div>
                <h2 className="mt-2 text-2xl font-semibold tracking-tight text-[color:var(--text)]">
                  {view === 'invite'
                    ? 'Set your password'
                    : view === 'mfa'
                      ? 'Verify your sign in'
                      : 'Sign in to continue'}
                </h2>
                <p className="mt-2 text-sm leading-6 text-[color:var(--muted)]">
                  {view === 'invite'
                    ? 'Complete the invite to activate the account, then optionally enroll MFA.'
                    : view === 'mfa'
                      ? 'Use your authenticator app or a recovery code.'
                      : 'Use your local username and password to unlock the platform.'}
                </p>
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

              {view === 'login' ? (
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
                        autoFocus={view === 'login'}
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
                        autoFocus={view === 'login'}
                      />
                  </label>
                  <Button type="submit" className="w-full justify-center" disabled={busy}>
                    {busy ? 'Signing in...' : 'Sign in'}
                  </Button>
                </form>
              ) : null}

              {view === 'mfa' ? (
                <form
                  className="space-y-4"
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
              ) : null}

              {view === 'invite' ? (
                <div className="space-y-5">
                  <div className="rounded-3xl border border-[color:var(--line)] bg-[color:var(--surface-strong)] p-4">
                    <div className="text-sm font-semibold text-[color:var(--text)]">Invite details</div>
                    <div className="mt-2 text-sm text-[color:var(--muted)]">
                      {invite ? (
                        <div className="space-y-1">
                          <div><span className="font-medium text-[color:var(--text)]">User:</span> {invite.username}</div>
                          <div><span className="font-medium text-[color:var(--text)]">Email:</span> {invite.recipientEmail}</div>
                          <div><span className="font-medium text-[color:var(--text)]">Status:</span> {invite.inviteStatus}</div>
                        </div>
                      ) : (
                        <div>Loading invite details...</div>
                      )}
                    </div>
                  </div>

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
                        <Button variant="ghost" onClick={() => {
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
                        }}>
                          Download recovery codes
                        </Button>
                        <Button onClick={onInviteFinish}>Continue to platform</Button>
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>
        </div>
      </div>
    </section>
  )
}

export function MfaReminderDialog({
  open,
  username,
  onSetup,
  onSkip
}: {
  open: boolean
  username: string
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
            This account does not have multi-factor authentication enabled. Add MFA to reduce the risk of unauthorized access.
          </p>
          <div className="flex flex-wrap justify-end gap-2">
            <Button onClick={onSetup}>Set up MFA</Button>
            <Button variant="ghost" onClick={onSkip}>Skip for now</Button>
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
