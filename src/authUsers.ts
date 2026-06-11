import { randomBytes, scryptSync, timingSafeEqual } from 'crypto'
import { MongoClient, type Collection } from 'mongodb'

export interface AuthUserListItem {
  username: string
  createdAt: string
}

export interface AuthUserStore {
  listUsers(): Promise<AuthUserListItem[]>
  authenticate(username: string, password: string): Promise<AuthUserListItem | null>
  addUser(username: string, password: string): Promise<AuthUserListItem>
  deleteUser(username: string): Promise<AuthUserListItem | null>
  close(): Promise<void>
}

interface StoredAuthUserRecord extends AuthUserListItem {
  usernameKey: string
  salt: string
  passwordHash: string
  updatedAt: string
}

const DEFAULT_COLLECTION_NAME = 'pst_auth_users'

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

function hashAuthPassword(password: string, salt: string): string {
  return scryptSync(password, salt, 64).toString('hex')
}

function verifyAuthPassword(password: string, record: StoredAuthUserRecord): boolean {
  try {
    const expected = Buffer.from(record.passwordHash, 'hex')
    const actual = scryptSync(password, record.salt, expected.length)
    return expected.length === actual.length && timingSafeEqual(expected, actual)
  } catch {
    return false
  }
}

function buildAuthUserRecord(
  username: string,
  password: string,
  createdAt = new Date().toISOString(),
  updatedAt = createdAt
): StoredAuthUserRecord {
  const normalizedUsername = normalizeUsername(username)
  const salt = randomBytes(16).toString('hex')
  return {
    username: normalizedUsername,
    usernameKey: normalizeUsernameKey(normalizedUsername),
    salt,
    passwordHash: hashAuthPassword(password, salt),
    createdAt,
    updatedAt
  }
}

function createAuthStoreError(statusCode: number, message: string): Error & { statusCode: number } {
  const error = new Error(message) as Error & { statusCode: number }
  error.statusCode = statusCode
  return error
}

function normalizeStoredAuthUserRecord(value: unknown): StoredAuthUserRecord | null {
  if (!value || typeof value !== 'object') {
    return null
  }

  const username = normalizeUsername((value as { username?: unknown }).username)
  const usernameKey = normalizeUsernameKey(
    (value as { usernameKey?: unknown }).usernameKey || username
  )
  const salt = normalizeText((value as { salt?: unknown }).salt)
  const passwordHash = normalizeText((value as { passwordHash?: unknown }).passwordHash)
  const createdAt = normalizeText((value as { createdAt?: unknown }).createdAt)
  const updatedAt = normalizeText((value as { updatedAt?: unknown }).updatedAt)

  if (!username || !usernameKey || !salt || !passwordHash) {
    return null
  }

  return {
    username,
    usernameKey,
    salt,
    passwordHash,
    createdAt: createdAt || new Date().toISOString(),
    updatedAt: updatedAt || createdAt || new Date().toISOString()
  }
}

function toAuthUserListItem(record: StoredAuthUserRecord): AuthUserListItem {
  return {
    username: record.username,
    createdAt: record.createdAt
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

class MemoryAuthUserStore implements AuthUserStore {
  private readonly users = new Map<string, StoredAuthUserRecord>()

  constructor(seedUsers: Array<{ username: string; password: string }> = []) {
    for (const seed of Array.isArray(seedUsers) ? seedUsers : []) {
      this.seedUser(seed.username, seed.password)
    }
  }

  private seedUser(username: string, password: string): void {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    if (!normalizedUsername || !normalizedPassword.trim()) {
      return
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = this.users.get(key) || null
    if (!existing) {
      this.users.set(key, buildAuthUserRecord(normalizedUsername, normalizedPassword))
      return
    }

    if (existing.username !== normalizedUsername || !verifyAuthPassword(normalizedPassword, existing)) {
      this.users.set(
        key,
        buildAuthUserRecord(normalizedUsername, normalizedPassword, existing.createdAt, new Date().toISOString())
      )
    }
  }

  async listUsers(): Promise<AuthUserListItem[]> {
    return sortAuthUsers([...this.users.values()].map(toAuthUserListItem))
  }

  async authenticate(username: string, password: string): Promise<AuthUserListItem | null> {
    const key = normalizeUsernameKey(username)
    const record = this.users.get(key)
    if (!record) {
      return null
    }

    if (!verifyAuthPassword(String(password ?? ''), record)) {
      return null
    }

    return toAuthUserListItem(record)
  }

  async addUser(username: string, password: string): Promise<AuthUserListItem> {
    const normalizedUsername = normalizeUsername(username)
    const key = normalizeUsernameKey(normalizedUsername)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!String(password ?? '').trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }
    if (this.users.has(key)) {
      throw createAuthStoreError(409, 'User already exists')
    }

    const record = buildAuthUserRecord(normalizedUsername, String(password))
    this.users.set(key, record)
    return toAuthUserListItem(record)
  }

  async deleteUser(username: string): Promise<AuthUserListItem | null> {
    const key = normalizeUsernameKey(username)
    const record = this.users.get(key) || null
    if (!record) {
      return null
    }

    this.users.delete(key)
    return toAuthUserListItem(record)
  }

  async close(): Promise<void> {
    this.users.clear()
  }
}

export function createMemoryAuthUserStore(
  seedUsers: Array<{ username: string; password: string }> = []
): AuthUserStore {
  return new MemoryAuthUserStore(seedUsers)
}

export class MongoAuthUserStore implements AuthUserStore {
  constructor(
    private readonly collection: Collection<StoredAuthUserRecord>,
    private readonly client: MongoClient
  ) {}

  static async connect(
    uri: string,
    dbName = 'pst-extractor',
    seedUsers: Array<{ username: string; password: string }> = [],
    collectionName = DEFAULT_COLLECTION_NAME
  ): Promise<MongoAuthUserStore> {
    const client = new MongoClient(uri)
    await client.connect()
    const collection = client.db(dbName).collection<StoredAuthUserRecord>(collectionName)
    await collection.createIndex({ usernameKey: 1 }, { unique: true })
    await collection.createIndex({ createdAt: 1 })
    const store = new MongoAuthUserStore(collection, client)
    await store.seedUsers(seedUsers)
    return store
  }

  private async seedUsers(seedUsers: Array<{ username: string; password: string }>): Promise<void> {
    for (const seed of Array.isArray(seedUsers) ? seedUsers : []) {
      await this.seedUser(seed.username, seed.password)
    }
  }

  private async seedUser(username: string, password: string): Promise<void> {
    const normalizedUsername = normalizeUsername(username)
    const normalizedPassword = String(password ?? '')
    if (!normalizedUsername || !normalizedPassword.trim()) {
      return
    }

    const key = normalizeUsernameKey(normalizedUsername)
    const existing = normalizeStoredAuthUserRecord(
      await this.collection.findOne({ usernameKey: key })
    )

    if (!existing) {
      await this.collection.insertOne(buildAuthUserRecord(normalizedUsername, normalizedPassword))
      return
    }

    if (existing.username !== normalizedUsername || !verifyAuthPassword(normalizedPassword, existing)) {
      const now = new Date().toISOString()
      await this.collection.updateOne(
        { usernameKey: key },
        {
          $set: buildAuthUserRecord(normalizedUsername, normalizedPassword, existing.createdAt, now)
        },
        { upsert: true }
      )
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

  async authenticate(username: string, password: string): Promise<AuthUserListItem | null> {
    const key = normalizeUsernameKey(username)
    const record = normalizeStoredAuthUserRecord(await this.collection.findOne({ usernameKey: key }))
    if (!record) {
      return null
    }

    if (!verifyAuthPassword(String(password ?? ''), record)) {
      return null
    }

    return toAuthUserListItem(record)
  }

  async addUser(username: string, password: string): Promise<AuthUserListItem> {
    const normalizedUsername = normalizeUsername(username)
    const key = normalizeUsernameKey(normalizedUsername)
    if (!normalizedUsername) {
      throw createAuthStoreError(400, 'Username is required')
    }
    if (!String(password ?? '').trim()) {
      throw createAuthStoreError(400, 'Password is required')
    }

    const existing = normalizeStoredAuthUserRecord(await this.collection.findOne({ usernameKey: key }))
    if (existing) {
      throw createAuthStoreError(409, 'User already exists')
    }

    const record = buildAuthUserRecord(normalizedUsername, String(password))
    await this.collection.insertOne(record)
    return toAuthUserListItem(record)
  }

  async deleteUser(username: string): Promise<AuthUserListItem | null> {
    const key = normalizeUsernameKey(username)
    const result = normalizeStoredAuthUserRecord(
      await this.collection.findOneAndDelete({ usernameKey: key })
    )
    if (!result) {
      return null
    }

    return toAuthUserListItem(result)
  }

  async close(): Promise<void> {
    await this.client.close()
  }
}
