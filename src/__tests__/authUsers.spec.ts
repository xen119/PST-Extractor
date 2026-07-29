import { MongoAuthUserStore, createMemoryAuthUserStore, type AuthUserStore } from '../authUsers'
import { DEFAULT_PASSWORD_POLICY } from '../passwordPolicy'

function createMongoAuthUserStore(): { store: AuthUserStore; close: () => Promise<void> } {
  const records = new Map<string, Record<string, unknown>>()

  function matchesFilter(record: Record<string, unknown>, filter: Record<string, unknown>): boolean {
    for (const [key, expected] of Object.entries(filter || {})) {
      if (expected === undefined) {
        continue
      }
      if (key === '$or' && Array.isArray(expected)) {
        return expected.some((entry) => entry && typeof entry === 'object' && matchesFilter(record, entry as Record<string, unknown>))
      }
      if (record[key] !== expected) {
        return false
      }
    }
    return true
  }

  const collection = {
    async findOne(filter: Record<string, unknown>): Promise<Record<string, unknown> | null> {
      for (const record of records.values()) {
        if (matchesFilter(record, filter)) {
          return { ...record }
        }
      }
      return null
    },
    async updateOne(filter: Record<string, unknown>, update: { $set?: Record<string, unknown> }): Promise<void> {
      const key = String(filter.usernameKey || '')
      const current = key ? records.get(key) || {} : {}
      const next = { ...current, ...(update.$set || {}) }
      if (next.usernameKey) {
        records.set(String(next.usernameKey), next)
      }
    },
    async deleteOne(filter: Record<string, unknown>): Promise<void> {
      const key = String(filter.usernameKey || '')
      if (key) {
        records.delete(key)
      }
    },
    async insertOne(document: Record<string, unknown>): Promise<void> {
      if (document.usernameKey) {
        records.set(String(document.usernameKey), { ...document })
      }
    },
    find(filter: Record<string, unknown> = {}): { sort: () => { toArray: () => Promise<Record<string, unknown>[]> } } {
      return {
        sort: () => ({
          toArray: async () =>
            [...records.values()]
              .filter((record) => matchesFilter(record, filter))
              .map((record) => ({ ...record }))
              .sort((left, right) => {
                const leftTime = Date.parse(String(left.createdAt || ''))
                const rightTime = Date.parse(String(right.createdAt || ''))
                if (rightTime !== leftTime) {
                  return leftTime - rightTime
                }
                return String(left.username || '').localeCompare(String(right.username || ''), undefined, {
                  sensitivity: 'base'
                })
              })
        })
      }
    },
    async createIndex(): Promise<void> {
      return undefined
    }
  }

  const metaCollection = {
    async findOne(): Promise<null> {
      return null
    },
    async updateOne(): Promise<void> {
      return undefined
    },
    async createIndex(): Promise<void> {
      return undefined
    }
  }

  const client = {
    async close(): Promise<void> {
      return undefined
    }
  }

  return {
    store: new MongoAuthUserStore(collection as never, metaCollection as never, client as never, 'test-master-secret'),
    close: async () => {
      await client.close()
    }
  }
}

async function createStoreHarness(kind: 'memory' | 'mongo'): Promise<{ store: AuthUserStore; close: () => Promise<void> }> {
  if (kind === 'memory') {
    const store = createMemoryAuthUserStore()
    return {
      store,
      close: () => store.close()
    }
  }

  return createMongoAuthUserStore()
}

describe.each(['memory', 'mongo'] as const)('auth user store (%s)', (kind) => {
  it('gates self-service resets and supports admin reset flows', async () => {
    const harness = await createStoreHarness(kind)
    try {
      await harness.store.addUser('admin', 'Admin12345!!')

      const invite = await harness.store.createInvite('alice', 'alice@example.com', 60)
      await harness.store.acceptInvite(invite.inviteToken, 'Alice12345!!')

      await harness.store.addUser('bob', 'BobInitial123!!')

      const beforeFailures = await harness.store.requestPasswordReset(
        'alice',
        DEFAULT_PASSWORD_POLICY.resetTokenTtlMinutes,
        DEFAULT_PASSWORD_POLICY
      )
      expect(beforeFailures).toBeNull()

      const firstFailure = await harness.store.authenticate('alice', 'wrong-password', DEFAULT_PASSWORD_POLICY)
      expect(firstFailure.user).toBeNull()
      expect(firstFailure.loginFailedCount).toBe(1)
      expect(firstFailure.passwordResetAvailable).toBe(false)

      const secondFailure = await harness.store.authenticate('alice', 'wrong-password', DEFAULT_PASSWORD_POLICY)
      expect(secondFailure.user).toBeNull()
      expect(secondFailure.loginFailedCount).toBe(2)
      expect(secondFailure.passwordResetAvailable).toBe(true)

      const thresholdReset = await harness.store.requestPasswordReset(
        'alice',
        DEFAULT_PASSWORD_POLICY.resetTokenTtlMinutes,
        DEFAULT_PASSWORD_POLICY
      )
      expect(thresholdReset).not.toBeNull()
      expect(thresholdReset?.user.username).toBe('alice')
      expect(thresholdReset?.resetToken).toBeTruthy()

      const bypassReset = await harness.store.requestPasswordReset(
        'bob',
        DEFAULT_PASSWORD_POLICY.resetTokenTtlMinutes,
        DEFAULT_PASSWORD_POLICY,
        {
          bypassGate: true,
          allowMissingRecipient: true
        }
      )
      expect(bypassReset).not.toBeNull()
      expect(bypassReset?.user.username).toBe('bob')
      expect(bypassReset?.resetToken).toBeTruthy()

      const forcedChange = await harness.store.changePassword(
        'bob',
        'TempPass123!!',
        DEFAULT_PASSWORD_POLICY,
        true
      )
      expect(forcedChange.passwordChangeRequired).toBe(true)
      expect(forcedChange.passwordChangeRequired).toBe(true)

      const forcedLogin = await harness.store.authenticate('bob', 'TempPass123!!', DEFAULT_PASSWORD_POLICY)
      expect(forcedLogin.user?.username).toBe('bob')
      expect(forcedLogin.user?.passwordChangeRequired).toBe(true)
      expect(forcedLogin.passwordChangeRequired).toBe(true)

      const clearedChange = await harness.store.changePassword(
        'bob',
        'BetterPass123!!',
        DEFAULT_PASSWORD_POLICY,
        false
      )
      expect(clearedChange.passwordChangeRequired).toBe(false)

      const clearedLogin = await harness.store.authenticate('bob', 'BetterPass123!!', DEFAULT_PASSWORD_POLICY)
      expect(clearedLogin.user?.username).toBe('bob')
      expect(clearedLogin.user?.passwordChangeRequired).toBe(false)
      expect(clearedLogin.passwordChangeRequired).toBe(false)
    } finally {
      await harness.close()
    }
  })

  it('looks up local login identifiers by username or recipient email', async () => {
      const harness = await createStoreHarness(kind)
    try {
      await harness.store.addUser('alice', 'Alice12345!!')
      const invite = await harness.store.createInvite('bob', 'bob@example.com', 60)
      await harness.store.acceptInvite(invite.inviteToken, 'BobInitial123!!')

      const usernameMatches = await harness.store.findUsersByLoginIdentifier('ALICE')
      expect(usernameMatches.map((entry) => entry.username)).toContain('alice')

      const emailMatches = await harness.store.findUsersByLoginIdentifier('bob@example.com')
      expect(emailMatches.map((entry) => entry.username)).toContain('bob')
      expect(await harness.store.findUsersByLoginIdentifier('missing@example.com')).toEqual([])
    } finally {
      await harness.close()
    }
  })
})
