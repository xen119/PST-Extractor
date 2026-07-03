import { MongoClient, type Collection } from 'mongodb'
import {
  buildPasswordPolicyDefaultsFromEnv,
  buildPasswordPolicyView,
  mergePasswordPolicy,
  normalizePasswordPolicyInput,
  type PasswordPolicyInput,
  type PasswordPolicyRecord,
  type PasswordPolicyView
} from './passwordPolicy'

export interface SmtpSettingsRecord {
  enabled: boolean
  host: string
  port: number
  secure: boolean
  username: string
  password: string
  fromName: string
  fromAddress: string
  replyTo: string
}

export interface SmtpSettingsView extends Omit<SmtpSettingsRecord, 'password'> {
  hasPassword: boolean
}

export interface SmtpSettingsInput {
  enabled?: boolean
  host?: string
  port?: number | string
  secure?: boolean
  username?: string
  password?: string
  fromName?: string
  fromAddress?: string
  replyTo?: string
}

export interface AppSettingsStore {
  getSmtpSettings(): Promise<SmtpSettingsRecord>
  updateSmtpSettings(input: SmtpSettingsInput): Promise<SmtpSettingsRecord>
  getPasswordPolicy(): Promise<PasswordPolicyRecord>
  updatePasswordPolicy(input: PasswordPolicyInput): Promise<PasswordPolicyRecord>
  close(): Promise<void>
}

interface StoredSmtpSettingsRecord extends SmtpSettingsRecord {
  settingsKey: 'smtp'
  createdAt: string
  updatedAt: string
}

interface StoredPasswordPolicyRecord extends PasswordPolicyRecord {
  settingsKey: 'password-policy'
  createdAt: string
  updatedAt: string
}

interface SmtpDefaults extends Partial<SmtpSettingsRecord> {}

interface PasswordPolicyDefaults extends Partial<PasswordPolicyRecord> {}

const DEFAULT_COLLECTION_NAME = 'pst_app_settings'
const DEFAULT_SMTP_PORT = 587

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function parseBoolean(value: unknown, fallback = false): boolean {
  if (value === undefined || value === null || value === '') {
    return fallback
  }

  if (typeof value === 'boolean') {
    return value
  }

  return !['0', 'false', 'no', 'off'].includes(normalizeText(value).toLowerCase())
}

function parsePort(value: unknown, fallback = DEFAULT_SMTP_PORT): number {
  if (value === undefined || value === null || value === '') {
    return fallback
  }

  const parsed = Number.parseInt(String(value), 10)
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > 65535) {
    return fallback
  }

  return parsed
}

function cloneSmtpSettings(settings: SmtpSettingsRecord): SmtpSettingsRecord {
  return {
    enabled: Boolean(settings.enabled),
    host: normalizeText(settings.host),
    port: parsePort(settings.port, DEFAULT_SMTP_PORT),
    secure: Boolean(settings.secure),
    username: normalizeText(settings.username),
    password: normalizeText(settings.password),
    fromName: normalizeText(settings.fromName),
    fromAddress: normalizeText(settings.fromAddress),
    replyTo: normalizeText(settings.replyTo)
  }
}

function buildDefaultSmtpSettings(defaults: SmtpDefaults = {}): SmtpSettingsRecord {
  return cloneSmtpSettings({
    enabled: parseBoolean(defaults.enabled, false),
    host: normalizeText(defaults.host),
    port: parsePort(defaults.port, DEFAULT_SMTP_PORT),
    secure: parseBoolean(defaults.secure, false),
    username: normalizeText(defaults.username),
    password: normalizeText(defaults.password),
    fromName: normalizeText(defaults.fromName),
    fromAddress: normalizeText(defaults.fromAddress),
    replyTo: normalizeText(defaults.replyTo)
  })
}

export function mergeSmtpSettings(
  base: SmtpSettingsRecord,
  input: SmtpSettingsInput = {}
): SmtpSettingsRecord {
  const nextPassword =
    input.password === undefined
      ? base.password
      : String(input.password ?? '').trim() || base.password

  return cloneSmtpSettings({
    enabled: input.enabled === undefined ? base.enabled : Boolean(input.enabled),
    host: input.host === undefined ? base.host : normalizeText(input.host),
    port: input.port === undefined ? base.port : parsePort(input.port, base.port),
    secure: input.secure === undefined ? base.secure : Boolean(input.secure),
    username: input.username === undefined ? base.username : normalizeText(input.username),
    password: nextPassword,
    fromName: input.fromName === undefined ? base.fromName : normalizeText(input.fromName),
    fromAddress:
      input.fromAddress === undefined ? base.fromAddress : normalizeText(input.fromAddress),
    replyTo: input.replyTo === undefined ? base.replyTo : normalizeText(input.replyTo)
  })
}

function normalizeStoredSmtpSettingsRecord(value: unknown): StoredSmtpSettingsRecord | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const record = value as Partial<StoredSmtpSettingsRecord>
  if (normalizeText(record.settingsKey) !== 'smtp') {
    return null
  }

  const createdAt = normalizeText(record.createdAt)
  const updatedAt = normalizeText(record.updatedAt)
  const settings = buildDefaultSmtpSettings(record)
  return {
    settingsKey: 'smtp',
    createdAt: createdAt || new Date().toISOString(),
    updatedAt: updatedAt || createdAt || new Date().toISOString(),
    ...settings
  }
}

function toStoredRecord(
  settings: SmtpSettingsRecord,
  createdAt = new Date().toISOString(),
  updatedAt = createdAt
): StoredSmtpSettingsRecord {
  return {
    settingsKey: 'smtp',
    createdAt,
    updatedAt,
    ...cloneSmtpSettings(settings)
  }
}

function normalizeStoredPasswordPolicyRecord(value: unknown): StoredPasswordPolicyRecord | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const record = value as Partial<StoredPasswordPolicyRecord>
  if (normalizeText(record.settingsKey) !== 'password-policy') {
    return null
  }

  const createdAt = normalizeText(record.createdAt)
  const updatedAt = normalizeText(record.updatedAt)
  const policy = mergePasswordPolicy(
    buildPasswordPolicyDefaultsFromEnv(),
    normalizePasswordPolicyInput(record as PasswordPolicyInput)
  )
  return {
    settingsKey: 'password-policy',
    createdAt: createdAt || new Date().toISOString(),
    updatedAt: updatedAt || createdAt || new Date().toISOString(),
    ...policy
  }
}

function toStoredPasswordPolicyRecord(
  policy: PasswordPolicyRecord,
  createdAt = new Date().toISOString(),
  updatedAt = createdAt
): StoredPasswordPolicyRecord {
  return {
    settingsKey: 'password-policy',
    createdAt,
    updatedAt,
    ...buildPasswordPolicyView(policy)
  }
}

function createSettingsStoreError(
  statusCode: number,
  message: string
): Error & { statusCode: number } {
  const error = new Error(message) as Error & { statusCode: number }
  error.statusCode = statusCode
  return error
}

class MemoryAppSettingsStore implements AppSettingsStore {
  private smtp: StoredSmtpSettingsRecord | null
  private passwordPolicy: StoredPasswordPolicyRecord | null

  constructor(defaults: SmtpDefaults = {}, passwordPolicyDefaults: PasswordPolicyDefaults = {}) {
    this.smtp = null
    this.passwordPolicy = null
    this.defaults = buildDefaultSmtpSettings(defaults)
    this.passwordPolicyDefaults = mergePasswordPolicy(
      buildPasswordPolicyDefaultsFromEnv(),
      normalizePasswordPolicyInput(passwordPolicyDefaults)
    )
  }

  private readonly defaults: SmtpSettingsRecord
  private readonly passwordPolicyDefaults: PasswordPolicyRecord

  async getSmtpSettings(): Promise<SmtpSettingsRecord> {
    return cloneSmtpSettings(this.smtp || this.defaults)
  }

  async updateSmtpSettings(input: SmtpSettingsInput): Promise<SmtpSettingsRecord> {
    const base = cloneSmtpSettings(this.smtp || this.defaults)
    const next = mergeSmtpSettings(base, input)
    const now = new Date().toISOString()
    this.smtp = toStoredRecord(next, this.smtp?.createdAt || now, now)
    return cloneSmtpSettings(next)
  }

  async getPasswordPolicy(): Promise<PasswordPolicyRecord> {
    return { ...(this.passwordPolicy ? this.passwordPolicy : this.passwordPolicyDefaults) }
  }

  async updatePasswordPolicy(input: PasswordPolicyInput): Promise<PasswordPolicyRecord> {
    const base = { ...(this.passwordPolicy || this.passwordPolicyDefaults) }
    const next = mergePasswordPolicy(base, normalizePasswordPolicyInput(input))
    const now = new Date().toISOString()
    this.passwordPolicy = toStoredPasswordPolicyRecord(next, this.passwordPolicy?.createdAt || now, now)
    return { ...next }
  }

  async close(): Promise<void> {
    this.smtp = null
    this.passwordPolicy = null
  }
}

export function createMemoryAppSettingsStore(
  defaults: SmtpDefaults = {},
  passwordPolicyDefaults: PasswordPolicyDefaults = {}
): AppSettingsStore {
  return new MemoryAppSettingsStore(defaults, passwordPolicyDefaults)
}

export class MongoAppSettingsStore implements AppSettingsStore {
  constructor(
    private readonly collection: Collection<StoredSmtpSettingsRecord | StoredPasswordPolicyRecord>,
    private readonly client: MongoClient,
    private readonly defaults: SmtpSettingsRecord,
    private readonly passwordPolicyDefaults: PasswordPolicyRecord
  ) {}

  static async connect(
    uri: string,
    dbName = 'pst-extractor',
    defaults: SmtpDefaults = {},
    passwordPolicyDefaults: PasswordPolicyDefaults = {},
    collectionName = DEFAULT_COLLECTION_NAME
  ): Promise<MongoAppSettingsStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const collection = client.db(dbName).collection<StoredSmtpSettingsRecord | StoredPasswordPolicyRecord>(collectionName)
    await collection.createIndex({ settingsKey: 1 }, { unique: true })
    return new MongoAppSettingsStore(
      collection,
      client,
      buildDefaultSmtpSettings(defaults),
      mergePasswordPolicy(buildPasswordPolicyDefaultsFromEnv(), normalizePasswordPolicyInput(passwordPolicyDefaults))
    )
  }

  async getSmtpSettings(): Promise<SmtpSettingsRecord> {
    const record = normalizeStoredSmtpSettingsRecord(
      await this.collection.findOne({ settingsKey: 'smtp' })
    )
    return cloneSmtpSettings(record || this.defaults)
  }

  async updateSmtpSettings(input: SmtpSettingsInput): Promise<SmtpSettingsRecord> {
    const existing = normalizeStoredSmtpSettingsRecord(
      await this.collection.findOne({ settingsKey: 'smtp' })
    )
    const base = existing ? cloneSmtpSettings(existing) : cloneSmtpSettings(this.defaults)
    const next = mergeSmtpSettings(base, input)
    const now = new Date().toISOString()
    await this.collection.updateOne(
      { settingsKey: 'smtp' },
      {
        $set: toStoredRecord(next, existing?.createdAt || now, now)
      },
      { upsert: true }
    )
    return cloneSmtpSettings(next)
  }

  async getPasswordPolicy(): Promise<PasswordPolicyRecord> {
    const record = normalizeStoredPasswordPolicyRecord(
      await this.collection.findOne({ settingsKey: 'password-policy' })
    )
    return { ...(record || this.passwordPolicyDefaults) }
  }

  async updatePasswordPolicy(input: PasswordPolicyInput): Promise<PasswordPolicyRecord> {
    const existing = normalizeStoredPasswordPolicyRecord(
      await this.collection.findOne({ settingsKey: 'password-policy' })
    )
    const base = existing ? { ...existing } : { ...this.passwordPolicyDefaults }
    const next = mergePasswordPolicy(base, normalizePasswordPolicyInput(input))
    const now = new Date().toISOString()
    await this.collection.updateOne(
      { settingsKey: 'password-policy' },
      {
        $set: toStoredPasswordPolicyRecord(next, existing?.createdAt || now, now)
      },
      { upsert: true }
    )
    return { ...next }
  }

  async close(): Promise<void> {
    await this.client.close()
  }
}

export function buildSmtpSettingsView(settings: SmtpSettingsRecord): SmtpSettingsView {
  return {
    enabled: Boolean(settings.enabled),
    host: normalizeText(settings.host),
    port: parsePort(settings.port, DEFAULT_SMTP_PORT),
    secure: Boolean(settings.secure),
    username: normalizeText(settings.username),
    fromName: normalizeText(settings.fromName),
    fromAddress: normalizeText(settings.fromAddress),
    replyTo: normalizeText(settings.replyTo),
    hasPassword: Boolean(normalizeText(settings.password))
  }
}

export function normalizeSmtpSettingsInput(input: Partial<SmtpSettingsInput>): SmtpSettingsInput {
  return {
    enabled: input.enabled === undefined ? undefined : Boolean(input.enabled),
    host: input.host === undefined ? undefined : normalizeText(input.host),
    port: input.port === undefined ? undefined : parsePort(input.port, DEFAULT_SMTP_PORT),
    secure: input.secure === undefined ? undefined : Boolean(input.secure),
    username: input.username === undefined ? undefined : normalizeText(input.username),
    password: input.password === undefined ? undefined : String(input.password),
    fromName: input.fromName === undefined ? undefined : normalizeText(input.fromName),
    fromAddress: input.fromAddress === undefined ? undefined : normalizeText(input.fromAddress),
    replyTo: input.replyTo === undefined ? undefined : normalizeText(input.replyTo)
  }
}

export function buildSmtpDefaultsFromEnv(
  env: NodeJS.ProcessEnv = process.env
): SmtpDefaults {
  return {
    enabled: parseBoolean(env.SMTP_ENABLED, false),
    host: normalizeText(env.SMTP_HOST),
    port: parsePort(env.SMTP_PORT, DEFAULT_SMTP_PORT),
    secure: parseBoolean(env.SMTP_SECURE, false),
    username: normalizeText(env.SMTP_USERNAME),
    password: String(env.SMTP_PASSWORD || ''),
    fromName: normalizeText(env.SMTP_FROM_NAME),
    fromAddress: normalizeText(env.SMTP_FROM_ADDRESS),
    replyTo: normalizeText(env.SMTP_REPLY_TO)
  }
}

export async function createAppSettingsStoreFromEnv(
  env: NodeJS.ProcessEnv = process.env
): Promise<AppSettingsStore> {
  const defaults = buildSmtpDefaultsFromEnv(env)
  const passwordPolicyDefaults = buildPasswordPolicyDefaultsFromEnv(env)
  const uri = normalizeText(env.MONGODB_URI)
  if (!uri) {
    return createMemoryAppSettingsStore(defaults, passwordPolicyDefaults)
  }

  const dbName = normalizeText(env.MONGODB_DB) || 'pst-extractor'
  return MongoAppSettingsStore.connect(uri, dbName, defaults, passwordPolicyDefaults)
}

export function createAppSettingsStoreError(
  statusCode: number,
  message: string
): Error & { statusCode: number } {
  return createSettingsStoreError(statusCode, message)
}
