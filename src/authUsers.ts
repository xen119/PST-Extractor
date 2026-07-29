import { createHmac, randomBytes } from 'crypto'
import { MongoClient, type Collection } from 'mongodb'
import {
  buildRecoveryCodeRecords,
  buildTotpOtpauthUri,
  createAuthMasterSecret,
  decryptSecret,
  encryptSecret,
  generateRecoveryCodes,
  generateTotpSecret,
  hashAuthPassword,
  normalizeAuthSecret,
  verifyAuthPassword,
  verifyRecoveryCode,
  verifyTotpCode,
  type RecoveryCodeRecord
} from './authSecurity'
import {
  buildPasswordPolicyView,
  mergePasswordPolicy,
  normalizePasswordPolicyInput,
  validatePasswordAgainstPolicy,
  type PasswordPolicyInput,
  type PasswordPolicyRecord,
  type PasswordPolicyView
} from './passwordPolicy'

export type AuthInviteStatus = 'pending' | 'active' | 'revoked' | 'expired'

export interface AuthUserListItem {
  username: string
  createdAt: string
  recipientEmail: string
  inviteStatus: AuthInviteStatus
  inviteSentAt: string
  inviteExpiresAt: string
  inviteAcceptedAt: string
  inviteRevokedAt: string
  mfaEnabled: boolean
  mfaEnforced: boolean
  mfaEnrolledAt: string
  assignedCasePaths: string[]
  passwordChangeRequired: boolean
}

export interface AuthInviteResult {
  user: AuthUserListItem
  inviteToken: string
  inviteExpiresAt: string
}

export interface AuthLoginResult {
  user: AuthUserListItem | null
  lockedUntil: string | null
  passwordResetAvailable: boolean
  loginFailedCount: number
  passwordChangeRequired: boolean
}

export interface AuthPasswordResetResult {
  user: AuthUserListItem
  resetToken: string
  resetExpiresAt: string
}

export interface AuthMfaSetupResult {
  user: AuthUserListItem
  secret: string
  otpauthUri: string
}

export interface AuthMfaCompletionResult {
  user: AuthUserListItem
  recoveryCodes: string[]
}

export interface AuthUserStore {
  listUsers(): Promise<AuthUserListItem[]>
  getUser(username: string): Promise<AuthUserListItem | null>
  findUsersByLoginIdentifier(identifier: string): Promise<AuthUserListItem[]>
  authenticate(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthLoginResult>
  addUser(username: string, password: string): Promise<AuthUserListItem>
  createInvite(
    username: string,
    recipientEmail: string,
    inviteTtlMinutes: number
  ): Promise<AuthInviteResult>
  resendInvite(username: string, inviteTtlMinutes: number): Promise<AuthInviteResult | null>
  revokeInvite(username: string): Promise<AuthUserListItem | null>
  getInviteByToken(token: string): Promise<AuthUserListItem | null>
  acceptInvite(token: string, password: string, policy?: PasswordPolicyRecord): Promise<AuthUserListItem>
  requestPasswordReset(
    usernameOrEmail: string,
    resetTtlMinutes: number,
    policy?: PasswordPolicyRecord,
    options?: {
      bypassGate?: boolean
      allowMissingRecipient?: boolean
    }
  ): Promise<AuthPasswordResetResult | null>
  getPasswordResetByToken(token: string): Promise<AuthUserListItem | null>
  resetPassword(
    token: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthUserListItem>
  changePassword(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord,
    passwordChangeRequired?: boolean
  ): Promise<AuthUserListItem>
  startMfaEnrollment(username: string, issuer?: string): Promise<AuthMfaSetupResult>
  completeMfaEnrollment(username: string, code: string): Promise<AuthMfaCompletionResult>
  verifyMfaChallenge(username: string, code: string): Promise<AuthUserListItem | null>
  resetMfa(username: string): Promise<AuthUserListItem | null>
  setMfaEnforced(username: string, enforced: boolean): Promise<AuthUserListItem | null>
  setAssignedCasePaths(username: string, assignedCasePaths: string[]): Promise<AuthUserListItem | null>
  deleteUser(username: string): Promise<AuthUserListItem | null>
  close(): Promise<void>
}

interface StoredAuthUserRecord extends AuthUserListItem {
  usernameKey: string
  salt: string
  passwordHash: string
  updatedAt: string
  inviteTokenHash: string
  inviteIssuedAt: string
  mfaSecretEncrypted: string
  mfaSecretNonce: string
  mfaPendingSecretEncrypted: string
  mfaPendingSecretNonce: string
  mfaPendingSecretIssuedAt: string
  mfaRecoveryCodes: RecoveryCodeRecord[]
  assignedCasePaths: string[]
  loginFailedCount: number
  lockedUntil: string
  passwordResetTokenHash: string
  passwordResetIssuedAt: string
  passwordResetExpiresAt: string
}

interface AuthMetaRecord {
  key: 'auth-master-secret'
  value: string
  createdAt: string
  updatedAt: string
}

const DEFAULT_COLLECTION_NAME = 'pst_auth_users'
const DEFAULT_META_COLLECTION_NAME = 'pst_auth_meta'
const DEFAULT_INVITE_TTL_MINUTES = 24 * 60
const DEFAULT_MFA_RECOVERY_CODE_COUNT = 8
const DEFAULT_MFA_ISSUER = 'PST Mail Explorer'

function createAuthStoreError(statusCode: number, message: string): Error & { statusCode: number } {
  const error = new Error(message) as Error & { statusCode: number }
  error.statusCode = statusCode
  return error
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeUsername(value: unknown): string {
  return normalizeText(value)
}

function normalizeUsernameKey(value: unknown): string {
  return normalizeUsername(value).toLowerCase()
}

function normalizeEmail(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

function normalizeLoginIdentifier(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

function normalizeCasePath(value: unknown): string {
  return normalizeText(value)
    .replace(/\\/g, '/')
    .replace(/^\/+|\/+$/g, '')
    .replace(/\/+/g, '/')
}

function normalizeAssignedCasePath(value: unknown): string {
  const normalized = normalizeCasePath(value)
  if (!normalized) {
    return ''
  }

  return normalized.split('/').map((segment) => segment.trim()).filter(Boolean)[0] || ''
}

function normalizeAssignedCasePaths(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return []
  }

  return [...new Set(value.map((entry) => normalizeAssignedCasePath(entry)).filter(Boolean))]
}

function normalizeToken(value: unknown): string {
  return normalizeText(value).toLowerCase()
}

function hashInviteToken(token: string, masterSecret: string): string {
  return createHmac('sha256', normalizeAuthSecret(masterSecret))
    .update(normalizeToken(token), 'utf8')
    .digest('hex')
}

function hashPasswordResetToken(token: string, masterSecret: string): string {
  return createHmac('sha256', normalizeAuthSecret(masterSecret))
    .update(`reset:${normalizeToken(token)}`, 'utf8')
    .digest('hex')
}

function normalizePositiveInt(value: unknown, fallback = 0): number {
  const parsed = Number.parseInt(String(value ?? ''), 10)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback
}

function normalizeRecoveryCodeRecords(value: unknown): RecoveryCodeRecord[] {
  if (!Array.isArray(value)) {
    return []
  }

  return value
    .map((entry) => {
      if (!entry || typeof entry !== 'object') {
        return null
      }

      const salt = normalizeText((entry as { salt?: unknown }).salt)
      const hash = normalizeText((entry as { hash?: unknown }).hash)
      const usedAt = normalizeText((entry as { usedAt?: unknown }).usedAt)
      if (!salt || !hash) {
        return null
      }

      return {
        salt,
        hash,
        usedAt
      }
    })
    .filter((entry): entry is RecoveryCodeRecord => Boolean(entry))
}

function buildStoredAuthUserRecord(input: {
  username: string
  createdAt?: string
  updatedAt?: string
  salt?: string
  passwordHash?: string
  recipientEmail?: string
  inviteStatus?: AuthInviteStatus
  inviteSentAt?: string
  inviteExpiresAt?: string
  inviteAcceptedAt?: string
  inviteRevokedAt?: string
  inviteIssuedAt?: string
  inviteTokenHash?: string
  mfaEnabled?: boolean
  mfaEnforced?: boolean
  mfaEnrolledAt?: string
  mfaSecretEncrypted?: string
  mfaSecretNonce?: string
  mfaPendingSecretEncrypted?: string
  mfaPendingSecretNonce?: string
  mfaPendingSecretIssuedAt?: string
  mfaRecoveryCodes?: RecoveryCodeRecord[]
  assignedCasePaths?: string[]
  loginFailedCount?: number
  lockedUntil?: string
  passwordResetTokenHash?: string
  passwordResetIssuedAt?: string
  passwordResetExpiresAt?: string
  passwordChangeRequired?: boolean
}): StoredAuthUserRecord {
  const username = normalizeUsername(input.username)
  const createdAt = normalizeText(input.createdAt) || new Date().toISOString()
  const updatedAt = normalizeText(input.updatedAt) || createdAt
  const passwordHash = normalizeText(input.passwordHash)
  return {
    username,
    usernameKey: normalizeUsernameKey(username),
    createdAt,
    updatedAt,
    salt: normalizeText(input.salt),
    passwordHash,
    recipientEmail: normalizeEmail(input.recipientEmail),
    inviteStatus: input.inviteStatus || (passwordHash ? 'active' : 'pending'),
    inviteSentAt: normalizeText(input.inviteSentAt),
    inviteExpiresAt: normalizeText(input.inviteExpiresAt),
    inviteAcceptedAt: normalizeText(input.inviteAcceptedAt),
    inviteRevokedAt: normalizeText(input.inviteRevokedAt),
    inviteIssuedAt: normalizeText(input.inviteIssuedAt),
    inviteTokenHash: normalizeText(input.inviteTokenHash),
    mfaEnabled: Boolean(input.mfaEnabled),
    mfaEnforced: Boolean(input.mfaEnforced),
    mfaEnrolledAt: normalizeText(input.mfaEnrolledAt),
    mfaSecretEncrypted: normalizeText(input.mfaSecretEncrypted),
    mfaSecretNonce: normalizeText(input.mfaSecretNonce),
    mfaPendingSecretEncrypted: normalizeText(input.mfaPendingSecretEncrypted),
    mfaPendingSecretNonce: normalizeText(input.mfaPendingSecretNonce),
    mfaPendingSecretIssuedAt: normalizeText(input.mfaPendingSecretIssuedAt),
    mfaRecoveryCodes: Array.isArray(input.mfaRecoveryCodes) ? input.mfaRecoveryCodes : [],
    assignedCasePaths: normalizeAssignedCasePaths(input.assignedCasePaths),
    loginFailedCount: normalizePositiveInt(input.loginFailedCount, 0),
    lockedUntil: normalizeText(input.lockedUntil),
    passwordResetTokenHash: normalizeText(input.passwordResetTokenHash),
    passwordResetIssuedAt: normalizeText(input.passwordResetIssuedAt),
    passwordResetExpiresAt: normalizeText(input.passwordResetExpiresAt),
    passwordChangeRequired: Boolean(input.passwordChangeRequired)
  }
}

function normalizeStoredAuthUserRecord(value: unknown): StoredAuthUserRecord | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const source = value as Partial<StoredAuthUserRecord>
  const username = normalizeUsername(source.username)
  const usernameKey = normalizeUsernameKey(source.usernameKey || username)
  if (!username || !usernameKey) {
    return null
  }

  const passwordHash = normalizeText(source.passwordHash)
  return {
    username,
    usernameKey,
    createdAt: normalizeText(source.createdAt) || new Date().toISOString(),
    updatedAt: normalizeText(source.updatedAt) || normalizeText(source.createdAt) || new Date().toISOString(),
    salt: normalizeText(source.salt),
    passwordHash,
    recipientEmail: normalizeEmail(source.recipientEmail),
    inviteStatus: (source.inviteStatus as AuthInviteStatus) || (passwordHash ? 'active' : 'pending'),
    inviteSentAt: normalizeText(source.inviteSentAt),
    inviteExpiresAt: normalizeText(source.inviteExpiresAt),
    inviteAcceptedAt: normalizeText(source.inviteAcceptedAt),
    inviteRevokedAt: normalizeText(source.inviteRevokedAt),
    inviteIssuedAt: normalizeText(source.inviteIssuedAt),
    inviteTokenHash: normalizeText(source.inviteTokenHash),
    mfaEnabled: Boolean(source.mfaEnabled),
    mfaEnforced: Boolean(source.mfaEnforced),
    mfaEnrolledAt: normalizeText(source.mfaEnrolledAt),
    mfaSecretEncrypted: normalizeText(source.mfaSecretEncrypted),
    mfaSecretNonce: normalizeText(source.mfaSecretNonce),
    mfaPendingSecretEncrypted: normalizeText(source.mfaPendingSecretEncrypted),
    mfaPendingSecretNonce: normalizeText(source.mfaPendingSecretNonce),
    mfaPendingSecretIssuedAt: normalizeText(source.mfaPendingSecretIssuedAt),
    mfaRecoveryCodes: normalizeRecoveryCodeRecords(source.mfaRecoveryCodes),
    assignedCasePaths: normalizeAssignedCasePaths(source.assignedCasePaths),
    loginFailedCount: normalizePositiveInt(source.loginFailedCount, 0),
    lockedUntil: normalizeText(source.lockedUntil),
    passwordResetTokenHash: normalizeText(source.passwordResetTokenHash),
    passwordResetIssuedAt: normalizeText(source.passwordResetIssuedAt),
    passwordResetExpiresAt: normalizeText(source.passwordResetExpiresAt),
    passwordChangeRequired: Boolean(source.passwordChangeRequired)
  }
}

function getInviteStatus(record: StoredAuthUserRecord): AuthInviteStatus {
  if (record.inviteStatus === 'revoked') {
    return 'revoked'
  }
  if (record.inviteStatus === 'active' || (record.passwordHash && record.salt && record.inviteAcceptedAt)) {
    return 'active'
  }
  const expiresAt = Date.parse(record.inviteExpiresAt || '')
  if (record.inviteStatus === 'pending' && Number.isFinite(expiresAt) && expiresAt > 0 && expiresAt <= Date.now()) {
    return 'expired'
  }
  return 'pending'
}

function toAuthUserListItem(record: StoredAuthUserRecord): AuthUserListItem {
  return {
    username: record.username,
    createdAt: record.createdAt,
    recipientEmail: record.recipientEmail,
    inviteStatus: getInviteStatus(record),
    inviteSentAt: record.inviteSentAt,
    inviteExpiresAt: record.inviteExpiresAt,
    inviteAcceptedAt: record.inviteAcceptedAt,
    inviteRevokedAt: record.inviteRevokedAt,
    mfaEnabled: Boolean(record.mfaEnabled),
    mfaEnforced: Boolean(record.mfaEnforced),
    mfaEnrolledAt: record.mfaEnrolledAt,
    assignedCasePaths: [...record.assignedCasePaths],
    passwordChangeRequired: Boolean(record.passwordChangeRequired)
  }
}

function sortAuthUsers(users: AuthUserListItem[]): AuthUserListItem[] {
  return [...users].sort((left, right) => {
    const leftTime = Date.parse(left.createdAt || '')
    const rightTime = Date.parse(right.createdAt || '')
    if (rightTime !== leftTime) {
      return leftTime - rightTime
    }
    return left.username.localeCompare(right.username, undefined, { sensitivity: 'base' })
  })
}

function buildLegacyUserRecord(
  username: string,
  password: string,
  createdAt = new Date().toISOString(),
  recipientEmail = ''
): StoredAuthUserRecord {
  const salt = randomBytes(16).toString('hex')
  return buildStoredAuthUserRecord({
    username,
    createdAt,
    updatedAt: createdAt,
    salt,
    passwordHash: hashAuthPassword(password, salt),
    recipientEmail,
    inviteStatus: 'active',
    inviteAcceptedAt: createdAt,
    inviteTokenHash: '',
    mfaEnabled: false,
    mfaEnforced: false,
    mfaRecoveryCodes: [],
    assignedCasePaths: [],
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired: false
  })
}

function buildInviteRecord(
  existing: StoredAuthUserRecord | null,
  username: string,
  recipientEmail: string,
  inviteTtlMinutes: number,
  masterSecret: string
): { record: StoredAuthUserRecord; inviteToken: string; inviteExpiresAt: string } {
  const now = new Date().toISOString()
  const inviteToken = randomBytes(24).toString('hex')
  const inviteExpiresAt = new Date(Date.now() + Math.max(1, inviteTtlMinutes || DEFAULT_INVITE_TTL_MINUTES) * 60 * 1000).toISOString()
  const tokenHash = hashInviteToken(inviteToken, masterSecret)
  const record = buildStoredAuthUserRecord({
    username,
    createdAt: existing?.createdAt || now,
    updatedAt: now,
    recipientEmail,
    inviteStatus: 'pending',
    inviteSentAt: now,
    inviteExpiresAt,
    inviteAcceptedAt: '',
    inviteRevokedAt: '',
    inviteIssuedAt: now,
    salt: '',
    passwordHash: '',
    mfaEnabled: false,
    mfaEnforced: existing?.mfaEnforced ?? false,
    mfaEnrolledAt: '',
    mfaSecretEncrypted: '',
    mfaSecretNonce: '',
    mfaPendingSecretEncrypted: '',
    mfaPendingSecretNonce: '',
    mfaPendingSecretIssuedAt: '',
    mfaRecoveryCodes: [],
    assignedCasePaths: existing?.assignedCasePaths || [],
    loginFailedCount: existing?.loginFailedCount || 0,
    lockedUntil: existing?.lockedUntil || '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired: false
  })

  return {
    record: {
      ...record,
      inviteTokenHash: tokenHash
    },
    inviteToken,
    inviteExpiresAt
  }
}

function buildDisabledMfaRecord(record: StoredAuthUserRecord): StoredAuthUserRecord {
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    mfaEnabled: false,
    mfaEnrolledAt: '',
    mfaSecretEncrypted: '',
    mfaSecretNonce: '',
    mfaPendingSecretEncrypted: '',
    mfaPendingSecretNonce: '',
    mfaPendingSecretIssuedAt: '',
    mfaRecoveryCodes: [],
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: ''
  })
}

function buildCompletedMfaRecord(
  record: StoredAuthUserRecord,
  secret: string,
  recoveryCodes: string[],
  masterSecret: string
): StoredAuthUserRecord {
  const encrypted = encryptSecret(secret, masterSecret)
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    mfaEnabled: true,
    mfaEnrolledAt: new Date().toISOString(),
    mfaSecretEncrypted: encrypted.ciphertext,
    mfaSecretNonce: encrypted.nonce,
    mfaPendingSecretEncrypted: '',
    mfaPendingSecretNonce: '',
    mfaPendingSecretIssuedAt: '',
    mfaRecoveryCodes: buildRecoveryCodeRecords(recoveryCodes),
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function buildMfaSetupRecord(record: StoredAuthUserRecord, secret: string, masterSecret: string): StoredAuthUserRecord {
  const encrypted = encryptSecret(secret, masterSecret)
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    mfaPendingSecretEncrypted: encrypted.ciphertext,
    mfaPendingSecretNonce: encrypted.nonce,
    mfaPendingSecretIssuedAt: new Date().toISOString(),
    loginFailedCount: record.loginFailedCount,
    lockedUntil: record.lockedUntil,
    passwordResetTokenHash: record.passwordResetTokenHash,
    passwordResetIssuedAt: record.passwordResetIssuedAt,
    passwordResetExpiresAt: record.passwordResetExpiresAt,
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function matchesLoginIdentifier(record: StoredAuthUserRecord, identifier: string): boolean {
  const normalizedIdentifier = normalizeLoginIdentifier(identifier)
  if (!normalizedIdentifier) {
    return false
  }

  return record.usernameKey === normalizedIdentifier || record.recipientEmail === normalizedIdentifier
}

function getLoginFailedCount(record: StoredAuthUserRecord): number {
  return normalizePositiveInt(record.loginFailedCount, 0)
}

function getLockedUntilTime(record: StoredAuthUserRecord): number {
  const timestamp = Date.parse(record.lockedUntil || '')
  return Number.isFinite(timestamp) ? timestamp : 0
}

function isLoginLocked(record: StoredAuthUserRecord): boolean {
  return getLockedUntilTime(record) > Date.now()
}

function buildLoginSuccessRecord(record: StoredAuthUserRecord): StoredAuthUserRecord {
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function buildLoginFailureRecord(record: StoredAuthUserRecord, policy?: PasswordPolicyRecord): StoredAuthUserRecord {
  const nextFailedCount = getLoginFailedCount(record) + 1
  const lockoutThreshold = Math.max(1, policy?.lockoutThreshold || 5)
  const shouldLock = nextFailedCount >= lockoutThreshold
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    loginFailedCount: nextFailedCount,
    lockedUntil: shouldLock
      ? new Date(Date.now() + Math.max(1, policy?.lockoutDurationSeconds || 30) * 1000).toISOString()
      : record.lockedUntil,
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function buildPasswordResetRecord(
  record: StoredAuthUserRecord,
  resetToken: string,
  resetExpiresAt: string,
  masterSecret: string
): StoredAuthUserRecord {
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    passwordResetTokenHash: hashPasswordResetToken(resetToken, masterSecret),
    passwordResetIssuedAt: new Date().toISOString(),
    passwordResetExpiresAt: resetExpiresAt,
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function buildPasswordResetClearedRecord(record: StoredAuthUserRecord): StoredAuthUserRecord {
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired: record.passwordChangeRequired
  })
}

function buildPasswordChangedRecord(
  record: StoredAuthUserRecord,
  password: string,
  passwordChangeRequired: boolean
): StoredAuthUserRecord {
  const salt = randomBytes(16).toString('hex')
  return buildStoredAuthUserRecord({
    ...record,
    updatedAt: new Date().toISOString(),
    salt,
    passwordHash: hashAuthPassword(password, salt),
    loginFailedCount: 0,
    lockedUntil: '',
    passwordResetTokenHash: '',
    passwordResetIssuedAt: '',
    passwordResetExpiresAt: '',
    passwordChangeRequired
  })
}

class MemoryAuthUserStore implements AuthUserStore {
  private readonly users = new Map<string, StoredAuthUserRecord>()
  private readonly masterSecret = createAuthMasterSecret()

  constructor(seedUsers: Array<{ username: string; password: string; recipientEmail?: string }> = []) {
    for (const seed of Array.isArray(seedUsers) ? seedUsers : []) {
      this.seedUser(seed.username, seed.password, seed.recipientEmail)
    }
  }

  private seedUser(username: string, password: string, recipientEmail = ''): void {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    const normalizedEmail = normalizeEmail(recipientEmail)
    if (!normalizedUsername || !normalizedPassword.trim()) {
      return
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = this.users.get(key) || null
    if (!existing) {
      this.users.set(key, buildLegacyUserRecord(normalizedUsername, normalizedPassword, undefined, normalizedEmail))
      return
    }

    if (
      existing.username !== normalizedUsername ||
      !existing.passwordHash ||
      !verifyAuthPassword(normalizedPassword, existing.salt, existing.passwordHash)
    ) {
      this.users.set(
        key,
        buildLegacyUserRecord(normalizedUsername, normalizedPassword, existing.createdAt, normalizedEmail || existing.recipientEmail)
      )
      return
    }

    if (normalizedEmail && existing.recipientEmail !== normalizedEmail) {
      this.users.set(key, {
        ...existing,
        recipientEmail: normalizedEmail,
        updatedAt: new Date().toISOString()
      })
    }
  }

  private getRecord(username: string): StoredAuthUserRecord | null {
    return this.users.get(normalizeUsernameKey(username)) || null
  }

  private getRecordByLoginIdentifier(identifier: string): StoredAuthUserRecord | null {
    const normalizedIdentifier = normalizeLoginIdentifier(identifier)
    if (!normalizedIdentifier) {
      return null
    }

    const byUsername = this.users.get(normalizedIdentifier) || null
    if (byUsername) {
      return byUsername
    }

    for (const record of this.users.values()) {
      if (matchesLoginIdentifier(record, normalizedIdentifier)) {
        return record
      }
    }

    return null
  }

  private setRecord(record: StoredAuthUserRecord): void {
    this.users.set(record.usernameKey, record)
  }

  private removeRecord(username: string): StoredAuthUserRecord | null {
    const key = normalizeUsernameKey(username)
    const record = this.users.get(key) || null
    if (!record) {
      return null
    }
    this.users.delete(key)
    return record
  }

  private findRecordByInviteToken(token: string): StoredAuthUserRecord | null {
    const tokenHash = hashInviteToken(token, this.masterSecret)
    for (const record of this.users.values()) {
      if (record.inviteTokenHash === tokenHash) {
        return record
      }
    }
    return null
  }

  private createInviteInternal(
    username: string,
    recipientEmail: string,
    inviteTtlMinutes: number
  ): AuthInviteResult {
    const normalizedUsername = normalizeUsername(username)
    const normalizedEmail = normalizeEmail(recipientEmail)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!normalizedEmail) {
      throw createAuthStoreError(400, 'Recipient email is required')
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = this.users.get(key) || null
    if (existing && getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'User already exists')
    }

    const { record, inviteToken, inviteExpiresAt } = buildInviteRecord(
      existing,
      normalizedUsername,
      normalizedEmail,
      inviteTtlMinutes,
      this.masterSecret
    )
    this.setRecord(record)
    return {
      user: toAuthUserListItem(record),
      inviteToken,
      inviteExpiresAt
    }
  }

  async listUsers(): Promise<AuthUserListItem[]> {
    return sortAuthUsers([...this.users.values()].map(toAuthUserListItem))
  }

  async getUser(username: string): Promise<AuthUserListItem | null> {
    const record = this.getRecord(username)
    return record ? toAuthUserListItem(record) : null
  }

  async findUsersByLoginIdentifier(identifier: string): Promise<AuthUserListItem[]> {
    const normalizedIdentifier = normalizeLoginIdentifier(identifier)
    if (!normalizedIdentifier) {
      return []
    }

    const matches = [...this.users.values()]
      .filter((record) => matchesLoginIdentifier(record, normalizedIdentifier))
      .map(toAuthUserListItem)

    const unique = new Map<string, AuthUserListItem>()
    for (const match of matches) {
      unique.set(normalizeUsernameKey(match.username), match)
    }
    return sortAuthUsers([...unique.values()])
  }

  async authenticate(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthLoginResult> {
    const record = this.getRecordByLoginIdentifier(username)
    if (!record || getInviteStatus(record) !== 'active') {
      return {
        user: null,
        lockedUntil: null,
        passwordResetAvailable: false,
        loginFailedCount: 0,
        passwordChangeRequired: false
      }
    }

    if (!record.passwordHash || !record.salt) {
      return {
        user: null,
        lockedUntil: null,
        passwordResetAvailable: false,
        loginFailedCount: 0,
        passwordChangeRequired: false
      }
    }

    if (isLoginLocked(record)) {
      return {
        user: null,
        lockedUntil: record.lockedUntil || null,
        passwordResetAvailable:
          getLoginFailedCount(record) >= Math.max(1, policy?.forgotPasswordAfterFailures || 2),
        loginFailedCount: getLoginFailedCount(record),
        passwordChangeRequired: Boolean(record.passwordChangeRequired)
      }
    }

    if (!verifyAuthPassword(String(password ?? ''), record.salt, record.passwordHash)) {
      const updated = buildLoginFailureRecord(record, policy)
      this.setRecord(updated)
      const failedCount = getLoginFailedCount(updated)
      return {
        user: null,
        lockedUntil: isLoginLocked(updated) ? updated.lockedUntil || null : null,
        passwordResetAvailable: failedCount >= Math.max(1, policy?.forgotPasswordAfterFailures || 2),
        loginFailedCount: failedCount,
        passwordChangeRequired: Boolean(updated.passwordChangeRequired)
      }
    }

    const updated = buildLoginSuccessRecord(record)
    this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      lockedUntil: null,
      passwordResetAvailable: false,
      loginFailedCount: 0,
      passwordChangeRequired: Boolean(updated.passwordChangeRequired)
    }
  }

  async addUser(username: string, password: string): Promise<AuthUserListItem> {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    const key = normalizeUsernameKey(normalizedUsername)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    if (this.users.has(key)) {
      throw createAuthStoreError(409, 'User already exists')
    }

    const record = buildLegacyUserRecord(normalizedUsername, normalizedPassword)
    this.setRecord(record)
    return toAuthUserListItem(record)
  }

  async createInvite(
    username: string,
    recipientEmail: string,
    inviteTtlMinutes: number
  ): Promise<AuthInviteResult> {
    return this.createInviteInternal(username, recipientEmail, inviteTtlMinutes)
  }

  async resendInvite(username: string, inviteTtlMinutes: number): Promise<AuthInviteResult | null> {
    const existing = this.getRecord(username)
    if (!existing) {
      return null
    }
    if (getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'User already exists')
    }
    if (!existing.recipientEmail) {
      throw createAuthStoreError(400, 'Recipient email is required')
    }
    return this.createInviteInternal(existing.username, existing.recipientEmail, inviteTtlMinutes)
  }

  async revokeInvite(username: string): Promise<AuthUserListItem | null> {
    const existing = this.getRecord(username)
    if (!existing) {
      return null
    }
    if (getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'Invite already accepted')
    }

    const record = buildStoredAuthUserRecord({
      ...existing,
      updatedAt: new Date().toISOString(),
      inviteStatus: 'revoked',
      inviteRevokedAt: new Date().toISOString()
    })
    this.setRecord(record)
    return toAuthUserListItem(record)
  }

  async getInviteByToken(token: string): Promise<AuthUserListItem | null> {
    const record = this.findRecordByInviteToken(token)
    return record && getInviteStatus(record) === 'pending' ? toAuthUserListItem(record) : null
  }

  async acceptInvite(
    token: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthUserListItem> {
    const record = this.findRecordByInviteToken(token)
    if (!record) {
      throw createAuthStoreError(404, 'Invite not found')
    }

    const status = getInviteStatus(record)
    if (status === 'revoked') {
      throw createAuthStoreError(410, 'Invite has been revoked')
    }
    if (status === 'expired') {
      throw createAuthStoreError(410, 'Invite has expired')
    }
    if (status === 'active') {
      throw createAuthStoreError(409, 'Invite already accepted')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    const policyIssues = validatePasswordAgainstPolicy(normalizedPassword, policy || mergePasswordPolicy(buildPasswordPolicyView({
      minLength: 12,
      requireUppercase: true,
      requireLowercase: true,
      requireNumber: true,
      requireSpecial: true,
      forgotPasswordAfterFailures: 2,
      lockoutThreshold: 5,
      lockoutDurationSeconds: 30,
      resetTokenTtlMinutes: 60,
      enforceMfa: false
    }), {}))
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const now = new Date().toISOString()
    const salt = randomBytes(16).toString('hex')
    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: now,
      salt,
      passwordHash: hashAuthPassword(normalizedPassword, salt),
      inviteStatus: 'active',
      inviteAcceptedAt: now,
      inviteRevokedAt: '',
      loginFailedCount: 0,
      lockedUntil: '',
      passwordResetTokenHash: '',
      passwordResetIssuedAt: '',
      passwordResetExpiresAt: ''
    })
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async requestPasswordReset(
    usernameOrEmail: string,
    resetTtlMinutes: number,
    policy?: PasswordPolicyRecord,
    options?: {
      bypassGate?: boolean
      allowMissingRecipient?: boolean
    }
  ): Promise<AuthPasswordResetResult | null> {
    const record = this.getRecordByLoginIdentifier(usernameOrEmail)
    if (
      !record ||
      getInviteStatus(record) !== 'active' ||
      (!record.recipientEmail && !options?.allowMissingRecipient)
    ) {
      return null
    }

    const resetThreshold = Math.max(1, policy?.forgotPasswordAfterFailures || 2)
    if (!options?.bypassGate && getLoginFailedCount(record) < resetThreshold) {
      return null
    }

    const resetToken = randomBytes(24).toString('hex')
    const resetExpiresAt = new Date(
      Date.now() + Math.max(1, resetTtlMinutes || 60) * 60 * 1000
    ).toISOString()
    const updated = buildPasswordResetRecord(record, resetToken, resetExpiresAt, this.masterSecret)
    this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      resetToken,
      resetExpiresAt
    }
  }

  async getPasswordResetByToken(token: string): Promise<AuthUserListItem | null> {
    const tokenHash = hashPasswordResetToken(token, this.masterSecret)
    for (const record of this.users.values()) {
      if (record.passwordResetTokenHash !== tokenHash) {
        continue
      }
      const expiresAt = Date.parse(record.passwordResetExpiresAt || '')
      if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
        return null
      }
      return toAuthUserListItem(record)
    }
    return null
  }

  async resetPassword(
    token: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthUserListItem> {
    const tokenHash = hashPasswordResetToken(token, this.masterSecret)
    const record = [...this.users.values()].find((entry) => entry.passwordResetTokenHash === tokenHash) || null
    if (!record) {
      throw createAuthStoreError(404, 'Password reset token not found')
    }

    const expiresAt = Date.parse(record.passwordResetExpiresAt || '')
    if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
      throw createAuthStoreError(410, 'Password reset token has expired')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    const policyIssues = validatePasswordAgainstPolicy(
      normalizedPassword,
      policy || buildPasswordPolicyView({
        minLength: 12,
        requireUppercase: true,
        requireLowercase: true,
        requireNumber: true,
        requireSpecial: true,
        forgotPasswordAfterFailures: 2,
        lockoutThreshold: 5,
        lockoutDurationSeconds: 30,
        resetTokenTtlMinutes: 60,
        enforceMfa: false
      })
    )
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const updated = buildPasswordChangedRecord(record, normalizedPassword, false)
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async changePassword(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord,
    passwordChangeRequired = false
  ): Promise<AuthUserListItem> {
    const record = this.getRecord(username)
    if (!record || getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(404, 'User not found')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    const policyIssues = validatePasswordAgainstPolicy(
      normalizedPassword,
      policy || buildPasswordPolicyView({
        minLength: 12,
        requireUppercase: true,
        requireLowercase: true,
        requireNumber: true,
        requireSpecial: true,
        forgotPasswordAfterFailures: 2,
        lockoutThreshold: 5,
        lockoutDurationSeconds: 30,
        resetTokenTtlMinutes: 60,
        enforceMfa: false
      })
    )
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const updated = buildPasswordChangedRecord(record, normalizedPassword, passwordChangeRequired)
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async startMfaEnrollment(username: string, issuer = DEFAULT_MFA_ISSUER): Promise<AuthMfaSetupResult> {
    const record = this.getRecord(username)
    if (!record) {
      throw createAuthStoreError(404, 'User not found')
    }
    if (getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(409, 'User is not active')
    }
    if (record.mfaEnabled) {
      throw createAuthStoreError(409, 'MFA is already enabled')
    }

    let secret = ''
    if (record.mfaPendingSecretEncrypted && record.mfaPendingSecretNonce) {
      secret = decryptSecret(
        {
          nonce: record.mfaPendingSecretNonce,
          ciphertext: record.mfaPendingSecretEncrypted
        },
        this.masterSecret
      )
    }
    if (!secret) {
      secret = generateTotpSecret()
      this.setRecord(buildMfaSetupRecord(record, secret, this.masterSecret))
    }

    const nextRecord = this.getRecord(username) || record
    return {
      user: toAuthUserListItem(nextRecord),
      secret,
      otpauthUri: buildTotpOtpauthUri({
        issuer: normalizeText(issuer) || DEFAULT_MFA_ISSUER,
        accountName: nextRecord.username,
        secret
      })
    }
  }

  async completeMfaEnrollment(username: string, code: string): Promise<AuthMfaCompletionResult> {
    const record = this.getRecord(username)
    if (!record) {
      throw createAuthStoreError(404, 'User not found')
    }
    if (getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(409, 'User is not active')
    }
    if (record.mfaEnabled) {
      throw createAuthStoreError(409, 'MFA is already enabled')
    }
    if (!record.mfaPendingSecretEncrypted || !record.mfaPendingSecretNonce) {
      throw createAuthStoreError(409, 'MFA enrollment has not started')
    }

    const secret = decryptSecret(
      {
        nonce: record.mfaPendingSecretNonce,
        ciphertext: record.mfaPendingSecretEncrypted
      },
      this.masterSecret
    )
    if (!secret || !verifyTotpCode(secret, code)) {
      throw createAuthStoreError(400, 'Invalid verification code')
    }

    const recoveryCodes = generateRecoveryCodes(DEFAULT_MFA_RECOVERY_CODE_COUNT)
    const updated = buildCompletedMfaRecord(record, secret, recoveryCodes, this.masterSecret)
    this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      recoveryCodes
    }
  }

  async verifyMfaChallenge(username: string, code: string): Promise<AuthUserListItem | null> {
    const record = this.getRecord(username)
    if (!record || !record.mfaEnabled) {
      return null
    }

    const activeSecret = decryptSecret(
      {
        nonce: record.mfaSecretNonce,
        ciphertext: record.mfaSecretEncrypted
      },
      this.masterSecret
    )
    if (activeSecret && verifyTotpCode(activeSecret, code)) {
      return toAuthUserListItem(record)
    }

    const normalizedCode = String(code ?? '')
      .trim()
      .replace(/\s+/g, '')
      .replace(/-/g, '')
    const recoveryCodes = [...record.mfaRecoveryCodes]
    for (let index = 0; index < recoveryCodes.length; index += 1) {
      const recovery = recoveryCodes[index]
      if (!recovery.usedAt && verifyRecoveryCode(normalizedCode, recovery.salt, recovery.hash)) {
        recoveryCodes[index] = {
          ...recovery,
          usedAt: new Date().toISOString()
        }
        const updated = buildStoredAuthUserRecord({
          ...record,
          updatedAt: new Date().toISOString(),
          mfaRecoveryCodes: recoveryCodes
        })
        this.setRecord(updated)
        return toAuthUserListItem(updated)
      }
    }

    return null
  }

  async resetMfa(username: string): Promise<AuthUserListItem | null> {
    const record = this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildDisabledMfaRecord(record)
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async setMfaEnforced(username: string, enforced: boolean): Promise<AuthUserListItem | null> {
    const record = this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: new Date().toISOString(),
      mfaEnforced: Boolean(enforced)
    })
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async setAssignedCasePaths(username: string, assignedCasePaths: string[]): Promise<AuthUserListItem | null> {
    const record = this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: new Date().toISOString(),
      assignedCasePaths
    })
    this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async deleteUser(username: string): Promise<AuthUserListItem | null> {
    const record = this.removeRecord(username)
    return record ? toAuthUserListItem(record) : null
  }

  async close(): Promise<void> {
    this.users.clear()
  }
}

export function createMemoryAuthUserStore(
  seedUsers: Array<{ username: string; password: string; recipientEmail?: string }> = []
): AuthUserStore {
  return new MemoryAuthUserStore(seedUsers)
}

export class MongoAuthUserStore implements AuthUserStore {
  constructor(
    private readonly collection: Collection<StoredAuthUserRecord>,
    private readonly metaCollection: Collection<AuthMetaRecord>,
    private readonly client: MongoClient,
    private readonly masterSecret: string
  ) {}

  static async connect(
    uri: string,
    dbName = 'pst-extractor',
    seedUsers: Array<{ username: string; password: string; recipientEmail?: string }> = [],
    collectionName = DEFAULT_COLLECTION_NAME,
    metaCollectionName = DEFAULT_META_COLLECTION_NAME
  ): Promise<MongoAuthUserStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const collection = client.db(dbName).collection<StoredAuthUserRecord>(collectionName)
    const metaCollection = client.db(dbName).collection<AuthMetaRecord>(metaCollectionName)
    await collection.createIndex({ usernameKey: 1 }, { unique: true })
    await collection.createIndex({ recipientEmail: 1 })
    await collection.createIndex({ inviteTokenHash: 1 })
    await collection.createIndex({ passwordResetTokenHash: 1 })
    await collection.createIndex({ createdAt: 1 })
    await metaCollection.createIndex({ key: 1 }, { unique: true })
    const masterSecret = await MongoAuthUserStore.ensureMasterSecret(metaCollection)
    const store = new MongoAuthUserStore(collection, metaCollection, client, masterSecret)
    await store.seedUsers(seedUsers)
    return store
  }

  private static async ensureMasterSecret(
    metaCollection: Collection<AuthMetaRecord>
  ): Promise<string> {
    const existing = await metaCollection.findOne({ key: 'auth-master-secret' })
    if (existing && normalizeText(existing.value)) {
      return normalizeText(existing.value)
    }

    const secret = createAuthMasterSecret()
    const now = new Date().toISOString()
    await metaCollection.updateOne(
      { key: 'auth-master-secret' },
      {
        $set: {
          key: 'auth-master-secret',
          value: secret,
          createdAt: existing?.createdAt || now,
          updatedAt: now
        }
      },
      { upsert: true }
    )
    return secret
  }

  private async seedUsers(seedUsers: Array<{ username: string; password: string; recipientEmail?: string }>): Promise<void> {
    for (const seed of Array.isArray(seedUsers) ? seedUsers : []) {
      await this.seedUser(seed.username, seed.password, seed.recipientEmail)
    }
  }

  private async seedUser(username: string, password: string, recipientEmail = ''): Promise<void> {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    const normalizedEmail = normalizeEmail(recipientEmail)
    if (!normalizedUsername || !normalizedPassword.trim()) {
      return
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = await this.getRecordByKey(key)
    if (!existing) {
      await this.collection.insertOne(
        buildLegacyUserRecord(normalizedUsername, normalizedPassword, undefined, normalizedEmail)
      )
      return
    }

    if (
      existing.username !== normalizedUsername ||
      !existing.passwordHash ||
      !verifyAuthPassword(normalizedPassword, existing.salt, existing.passwordHash)
    ) {
      await this.collection.updateOne(
        { usernameKey: key },
        {
          $set: buildLegacyUserRecord(
            normalizedUsername,
            normalizedPassword,
            existing.createdAt,
            normalizedEmail || existing.recipientEmail
          )
        },
        { upsert: true }
      )
      return
    }

    if (normalizedEmail && existing.recipientEmail !== normalizedEmail) {
      await this.collection.updateOne(
        { usernameKey: key },
        {
          $set: {
            ...existing,
            recipientEmail: normalizedEmail,
            updatedAt: new Date().toISOString()
          }
        }
      )
    }
  }

  private async getRecordByKey(key: string): Promise<StoredAuthUserRecord | null> {
    return normalizeStoredAuthUserRecord(await this.collection.findOne({ usernameKey: key }))
  }

  private async getRecordByLoginIdentifier(identifier: string): Promise<StoredAuthUserRecord | null> {
    const normalizedIdentifier = normalizeLoginIdentifier(identifier)
    if (!normalizedIdentifier) {
      return null
    }

    const byUsername = await this.getRecordByKey(normalizedIdentifier)
    if (byUsername) {
      return byUsername
    }

    return normalizeStoredAuthUserRecord(
      await this.collection.findOne({ recipientEmail: normalizedIdentifier })
    )
  }

  private async getRecord(username: string): Promise<StoredAuthUserRecord | null> {
    return this.getRecordByKey(normalizeUsernameKey(username))
  }

  private async setRecord(record: StoredAuthUserRecord): Promise<void> {
    await this.collection.updateOne(
      { usernameKey: record.usernameKey },
      {
        $set: record
      },
      { upsert: true }
    )
  }

  private async removeRecord(username: string): Promise<StoredAuthUserRecord | null> {
    const key = normalizeUsernameKey(username)
    const record = await this.getRecordByKey(key)
    if (!record) {
      return null
    }
    await this.collection.deleteOne({ usernameKey: key })
    return record
  }

  private async findRecordByInviteToken(token: string): Promise<StoredAuthUserRecord | null> {
    const tokenHash = hashInviteToken(token, this.masterSecret)
    return normalizeStoredAuthUserRecord(await this.collection.findOne({ inviteTokenHash: tokenHash }))
  }

  private async createInviteInternal(
    username: string,
    recipientEmail: string,
    inviteTtlMinutes: number
  ): Promise<AuthInviteResult> {
    const normalizedUsername = normalizeUsername(username)
    const normalizedEmail = normalizeEmail(recipientEmail)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!normalizedEmail) {
      throw createAuthStoreError(400, 'Recipient email is required')
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = await this.getRecordByKey(key)
    if (existing && getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'User already exists')
    }

    const { record, inviteToken, inviteExpiresAt } = buildInviteRecord(
      existing,
      normalizedUsername,
      normalizedEmail,
      inviteTtlMinutes,
      this.masterSecret
    )
    await this.setRecord(record)
    return {
      user: toAuthUserListItem(record),
      inviteToken,
      inviteExpiresAt
    }
  }

  async listUsers(): Promise<AuthUserListItem[]> {
    const records = await this.collection.find({}).sort({ createdAt: 1, username: 1 }).toArray()
    return sortAuthUsers(
      records
        .map((record) => normalizeStoredAuthUserRecord(record))
        .filter((record): record is StoredAuthUserRecord => Boolean(record))
        .map(toAuthUserListItem)
    )
  }

  async getUser(username: string): Promise<AuthUserListItem | null> {
    const record = await this.getRecord(username)
    return record ? toAuthUserListItem(record) : null
  }

  async findUsersByLoginIdentifier(identifier: string): Promise<AuthUserListItem[]> {
    const normalizedIdentifier = normalizeLoginIdentifier(identifier)
    if (!normalizedIdentifier) {
      return []
    }

    const records = await this.collection
      .find({
        $or: [{ usernameKey: normalizedIdentifier }, { recipientEmail: normalizedIdentifier }]
      })
      .sort({ createdAt: 1, username: 1 })
      .toArray()

    const unique = new Map<string, AuthUserListItem>()
    for (const record of records) {
      const normalized = normalizeStoredAuthUserRecord(record)
      if (!normalized) {
        continue
      }
      unique.set(normalized.usernameKey, toAuthUserListItem(normalized))
    }

    return sortAuthUsers([...unique.values()])
  }

  async authenticate(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthLoginResult> {
    const record = await this.getRecordByLoginIdentifier(username)
    if (!record || getInviteStatus(record) !== 'active') {
      return {
        user: null,
        lockedUntil: null,
        passwordResetAvailable: false,
        loginFailedCount: 0,
        passwordChangeRequired: false
      }
    }

    if (!record.passwordHash || !record.salt) {
      return {
        user: null,
        lockedUntil: null,
        passwordResetAvailable: false,
        loginFailedCount: 0,
        passwordChangeRequired: false
      }
    }

    if (isLoginLocked(record)) {
      const failedCount = getLoginFailedCount(record)
      return {
        user: null,
        lockedUntil: record.lockedUntil || null,
        passwordResetAvailable: failedCount >= Math.max(1, policy?.forgotPasswordAfterFailures || 2),
        loginFailedCount: failedCount,
        passwordChangeRequired: Boolean(record.passwordChangeRequired)
      }
    }

    if (!verifyAuthPassword(String(password ?? ''), record.salt, record.passwordHash)) {
      const updated = buildLoginFailureRecord(record, policy)
      await this.setRecord(updated)
      const failedCount = getLoginFailedCount(updated)
      return {
        user: null,
        lockedUntil: isLoginLocked(updated) ? updated.lockedUntil || null : null,
        passwordResetAvailable: failedCount >= Math.max(1, policy?.forgotPasswordAfterFailures || 2),
        loginFailedCount: failedCount,
        passwordChangeRequired: Boolean(updated.passwordChangeRequired)
      }
    }

    const updated = buildLoginSuccessRecord(record)
    await this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      lockedUntil: null,
      passwordResetAvailable: false,
      loginFailedCount: 0,
      passwordChangeRequired: Boolean(updated.passwordChangeRequired)
    }
  }

  async addUser(username: string, password: string): Promise<AuthUserListItem> {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    const key = normalizeUsernameKey(normalizedUsername)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    if (await this.getRecordByKey(key)) {
      throw createAuthStoreError(409, 'User already exists')
    }

    const record = buildLegacyUserRecord(normalizedUsername, normalizedPassword)
    await this.collection.insertOne(record)
    return toAuthUserListItem(record)
  }

  async createInvite(
    username: string,
    recipientEmail: string,
    inviteTtlMinutes: number
  ): Promise<AuthInviteResult> {
    return this.createInviteInternal(username, recipientEmail, inviteTtlMinutes)
  }

  async resendInvite(username: string, inviteTtlMinutes: number): Promise<AuthInviteResult | null> {
    const existing = await this.getRecord(username)
    if (!existing) {
      return null
    }
    if (getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'User already exists')
    }
    if (!existing.recipientEmail) {
      throw createAuthStoreError(400, 'Recipient email is required')
    }
    return this.createInviteInternal(existing.username, existing.recipientEmail, inviteTtlMinutes)
  }

  async revokeInvite(username: string): Promise<AuthUserListItem | null> {
    const existing = await this.getRecord(username)
    if (!existing) {
      return null
    }
    if (getInviteStatus(existing) === 'active') {
      throw createAuthStoreError(409, 'Invite already accepted')
    }

    const record = buildStoredAuthUserRecord({
      ...existing,
      updatedAt: new Date().toISOString(),
      inviteStatus: 'revoked',
      inviteRevokedAt: new Date().toISOString()
    })
    await this.setRecord(record)
    return toAuthUserListItem(record)
  }

  async getInviteByToken(token: string): Promise<AuthUserListItem | null> {
    const record = await this.findRecordByInviteToken(token)
    return record && getInviteStatus(record) === 'pending' ? toAuthUserListItem(record) : null
  }

  async acceptInvite(
    token: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthUserListItem> {
    const record = await this.findRecordByInviteToken(token)
    if (!record) {
      throw createAuthStoreError(404, 'Invite not found')
    }

    const status = getInviteStatus(record)
    if (status === 'revoked') {
      throw createAuthStoreError(410, 'Invite has been revoked')
    }
    if (status === 'expired') {
      throw createAuthStoreError(410, 'Invite has expired')
    }
    if (status === 'active') {
      throw createAuthStoreError(409, 'Invite already accepted')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    const policyIssues = validatePasswordAgainstPolicy(
      normalizedPassword,
      policy || buildPasswordPolicyView({
        minLength: 12,
        requireUppercase: true,
        requireLowercase: true,
        requireNumber: true,
        requireSpecial: true,
        forgotPasswordAfterFailures: 2,
        lockoutThreshold: 5,
        lockoutDurationSeconds: 30,
        resetTokenTtlMinutes: 60,
        enforceMfa: false
      })
    )
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const now = new Date().toISOString()
    const salt = randomBytes(16).toString('hex')
    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: now,
      salt,
      passwordHash: hashAuthPassword(normalizedPassword, salt),
      inviteStatus: 'active',
      inviteAcceptedAt: now,
      inviteRevokedAt: '',
      loginFailedCount: 0,
      lockedUntil: '',
      passwordResetTokenHash: '',
      passwordResetIssuedAt: '',
      passwordResetExpiresAt: ''
    })
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async requestPasswordReset(
    usernameOrEmail: string,
    resetTtlMinutes: number,
    policy?: PasswordPolicyRecord,
    options?: {
      bypassGate?: boolean
      allowMissingRecipient?: boolean
    }
  ): Promise<AuthPasswordResetResult | null> {
    const record = await this.getRecordByLoginIdentifier(usernameOrEmail)
    if (
      !record ||
      getInviteStatus(record) !== 'active' ||
      (!record.recipientEmail && !options?.allowMissingRecipient)
    ) {
      return null
    }

    const resetThreshold = Math.max(1, policy?.forgotPasswordAfterFailures || 2)
    if (!options?.bypassGate && getLoginFailedCount(record) < resetThreshold) {
      return null
    }

    const resetToken = randomBytes(24).toString('hex')
    const resetExpiresAt = new Date(
      Date.now() + Math.max(1, resetTtlMinutes || 60) * 60 * 1000
    ).toISOString()
    const updated = buildPasswordResetRecord(record, resetToken, resetExpiresAt, this.masterSecret)
    await this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      resetToken,
      resetExpiresAt
    }
  }

  async getPasswordResetByToken(token: string): Promise<AuthUserListItem | null> {
    const tokenHash = hashPasswordResetToken(token, this.masterSecret)
    const record = await this.collection.findOne({ passwordResetTokenHash: tokenHash })
    const normalized = normalizeStoredAuthUserRecord(record)
    if (!normalized) {
      return null
    }
    const expiresAt = Date.parse(normalized.passwordResetExpiresAt || '')
    if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
      return null
    }
    return toAuthUserListItem(normalized)
  }

  async resetPassword(
    token: string,
    password: string,
    policy?: PasswordPolicyRecord
  ): Promise<AuthUserListItem> {
    const tokenHash = hashPasswordResetToken(token, this.masterSecret)
    const record = normalizeStoredAuthUserRecord(
      await this.collection.findOne({ passwordResetTokenHash: tokenHash })
    )
    if (!record) {
      throw createAuthStoreError(404, 'Password reset token not found')
    }

    const expiresAt = Date.parse(record.passwordResetExpiresAt || '')
    if (!Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
      throw createAuthStoreError(410, 'Password reset token has expired')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    const policyIssues = validatePasswordAgainstPolicy(
      normalizedPassword,
      policy || buildPasswordPolicyView({
        minLength: 12,
        requireUppercase: true,
        requireLowercase: true,
        requireNumber: true,
        requireSpecial: true,
        forgotPasswordAfterFailures: 2,
        lockoutThreshold: 5,
        lockoutDurationSeconds: 30,
        resetTokenTtlMinutes: 60,
        enforceMfa: false
      })
    )
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const updated = buildPasswordChangedRecord(record, normalizedPassword, false)
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async changePassword(
    username: string,
    password: string,
    policy?: PasswordPolicyRecord,
    passwordChangeRequired = false
  ): Promise<AuthUserListItem> {
    const record = await this.getRecord(username)
    if (!record || getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(404, 'User not found')
    }

    const normalizedPassword = String(password ?? '')
    if (!normalizedPassword.trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }

    const policyIssues = validatePasswordAgainstPolicy(
      normalizedPassword,
      policy || buildPasswordPolicyView({
        minLength: 12,
        requireUppercase: true,
        requireLowercase: true,
        requireNumber: true,
        requireSpecial: true,
        forgotPasswordAfterFailures: 2,
        lockoutThreshold: 5,
        lockoutDurationSeconds: 30,
        resetTokenTtlMinutes: 60,
        enforceMfa: false
      })
    )
    if (policyIssues.length) {
      throw createAuthStoreError(400, policyIssues.join(' '))
    }

    const updated = buildPasswordChangedRecord(record, normalizedPassword, passwordChangeRequired)
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async startMfaEnrollment(username: string, issuer = DEFAULT_MFA_ISSUER): Promise<AuthMfaSetupResult> {
    const record = await this.getRecord(username)
    if (!record) {
      throw createAuthStoreError(404, 'User not found')
    }
    if (getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(409, 'User is not active')
    }
    if (record.mfaEnabled) {
      throw createAuthStoreError(409, 'MFA is already enabled')
    }

    let secret = ''
    if (record.mfaPendingSecretEncrypted && record.mfaPendingSecretNonce) {
      secret = decryptSecret(
        {
          nonce: record.mfaPendingSecretNonce,
          ciphertext: record.mfaPendingSecretEncrypted
        },
        this.masterSecret
      )
    }
    if (!secret) {
      secret = generateTotpSecret()
      await this.setRecord(buildMfaSetupRecord(record, secret, this.masterSecret))
    }

    const nextRecord = (await this.getRecord(username)) || record
    return {
      user: toAuthUserListItem(nextRecord),
      secret,
      otpauthUri: buildTotpOtpauthUri({
        issuer: normalizeText(issuer) || DEFAULT_MFA_ISSUER,
        accountName: nextRecord.username,
        secret
      })
    }
  }

  async completeMfaEnrollment(username: string, code: string): Promise<AuthMfaCompletionResult> {
    const record = await this.getRecord(username)
    if (!record) {
      throw createAuthStoreError(404, 'User not found')
    }
    if (getInviteStatus(record) !== 'active') {
      throw createAuthStoreError(409, 'User is not active')
    }
    if (record.mfaEnabled) {
      throw createAuthStoreError(409, 'MFA is already enabled')
    }
    if (!record.mfaPendingSecretEncrypted || !record.mfaPendingSecretNonce) {
      throw createAuthStoreError(409, 'MFA enrollment has not started')
    }

    const secret = decryptSecret(
      {
        nonce: record.mfaPendingSecretNonce,
        ciphertext: record.mfaPendingSecretEncrypted
      },
      this.masterSecret
    )
    if (!secret || !verifyTotpCode(secret, code)) {
      throw createAuthStoreError(400, 'Invalid verification code')
    }

    const recoveryCodes = generateRecoveryCodes(DEFAULT_MFA_RECOVERY_CODE_COUNT)
    const updated = buildCompletedMfaRecord(record, secret, recoveryCodes, this.masterSecret)
    await this.setRecord(updated)
    return {
      user: toAuthUserListItem(updated),
      recoveryCodes
    }
  }

  async verifyMfaChallenge(username: string, code: string): Promise<AuthUserListItem | null> {
    const record = await this.getRecord(username)
    if (!record || !record.mfaEnabled) {
      return null
    }

    const activeSecret = decryptSecret(
      {
        nonce: record.mfaSecretNonce,
        ciphertext: record.mfaSecretEncrypted
      },
      this.masterSecret
    )
    if (activeSecret && verifyTotpCode(activeSecret, code)) {
      return toAuthUserListItem(record)
    }

    const normalizedCode = String(code ?? '')
      .trim()
      .replace(/\s+/g, '')
      .replace(/-/g, '')
    const recoveryCodes = [...record.mfaRecoveryCodes]
    for (let index = 0; index < recoveryCodes.length; index += 1) {
      const recovery = recoveryCodes[index]
      if (!recovery.usedAt && verifyRecoveryCode(normalizedCode, recovery.salt, recovery.hash)) {
        recoveryCodes[index] = {
          ...recovery,
          usedAt: new Date().toISOString()
        }
        const updated = buildStoredAuthUserRecord({
          ...record,
          updatedAt: new Date().toISOString(),
          mfaRecoveryCodes: recoveryCodes
        })
        await this.setRecord(updated)
        return toAuthUserListItem(updated)
      }
    }

    return null
  }

  async resetMfa(username: string): Promise<AuthUserListItem | null> {
    const record = await this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildDisabledMfaRecord(record)
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async setMfaEnforced(username: string, enforced: boolean): Promise<AuthUserListItem | null> {
    const record = await this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: new Date().toISOString(),
      mfaEnforced: Boolean(enforced)
    })
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async setAssignedCasePaths(username: string, assignedCasePaths: string[]): Promise<AuthUserListItem | null> {
    const record = await this.getRecord(username)
    if (!record) {
      return null
    }

    const updated = buildStoredAuthUserRecord({
      ...record,
      updatedAt: new Date().toISOString(),
      assignedCasePaths
    })
    await this.setRecord(updated)
    return toAuthUserListItem(updated)
  }

  async deleteUser(username: string): Promise<AuthUserListItem | null> {
    return this.removeRecord(username).then((record) => (record ? toAuthUserListItem(record) : null))
  }

  async close(): Promise<void> {
    await this.client.close()
  }
}
