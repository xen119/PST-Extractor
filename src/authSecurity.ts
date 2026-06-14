import {
  createCipheriv,
  createDecipheriv,
  createHash,
  createHmac,
  randomBytes,
  scryptSync,
  timingSafeEqual
} from 'crypto'

const BASE32_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567'
const TOTP_PERIOD_SECONDS = 30
const TOTP_DIGITS = 6

export interface RecoveryCodeRecord {
  salt: string
  hash: string
  usedAt: string
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeCode(value: unknown): string {
  return normalizeText(value).replace(/[\s-]+/g, '')
}

function base32Encode(buffer: Buffer): string {
  if (!buffer.length) {
    return ''
  }

  let bits = 0
  let value = 0
  let output = ''

  for (const byte of buffer.values()) {
    value = (value << 8) | byte
    bits += 8

    while (bits >= 5) {
      output += BASE32_ALPHABET[(value >>> (bits - 5)) & 31]
      bits -= 5
    }
  }

  if (bits > 0) {
    output += BASE32_ALPHABET[(value << (5 - bits)) & 31]
  }

  return output
}

function base32Decode(value: string): Buffer {
  const normalized = normalizeText(value).replace(/=+$/g, '').toUpperCase()
  if (!normalized) {
    return Buffer.alloc(0)
  }

  let bits = 0
  let accumulator = 0
  const bytes: number[] = []

  for (const character of normalized) {
    const index = BASE32_ALPHABET.indexOf(character)
    if (index < 0) {
      continue
    }

    accumulator = (accumulator << 5) | index
    bits += 5

    if (bits >= 8) {
      bytes.push((accumulator >>> (bits - 8)) & 0xff)
      bits -= 8
    }
  }

  return Buffer.from(bytes)
}

function deriveSecretKey(secret: string): Buffer {
  return createHash('sha256').update(normalizeText(secret), 'utf8').digest()
}

export function createAuthMasterSecret(): string {
  return randomBytes(32).toString('hex')
}

export function normalizeAuthSecret(value: unknown): string {
  return normalizeText(value)
}

export function hashAuthPassword(password: string, salt: string): string {
  return scryptSync(String(password ?? ''), String(salt ?? ''), 64).toString('hex')
}

export function verifyAuthPassword(
  password: string,
  salt: string,
  passwordHash: string
): boolean {
  try {
    const expected = Buffer.from(String(passwordHash ?? ''), 'hex')
    if (!expected.length) {
      return false
    }
    const actual = scryptSync(String(password ?? ''), String(salt ?? ''), expected.length)
    return expected.length === actual.length && timingSafeEqual(expected, actual)
  } catch {
    return false
  }
}

export function encryptSecret(value: string, secretKey: string): { nonce: string; ciphertext: string } {
  const iv = randomBytes(12)
  const key = deriveSecretKey(secretKey)
  const cipher = createCipheriv('aes-256-gcm', key, iv)
  const ciphertext = Buffer.concat([
    cipher.update(String(value ?? ''), 'utf8'),
    cipher.final()
  ])
  const tag = cipher.getAuthTag()
  return {
    nonce: iv.toString('hex'),
    ciphertext: Buffer.concat([ciphertext, tag]).toString('hex')
  }
}

export function decryptSecret(
  payload: { nonce?: string; ciphertext?: string } | null | undefined,
  secretKey: string
): string {
  const nonce = normalizeText(payload?.nonce)
  const ciphertext = normalizeText(payload?.ciphertext)
  if (!nonce || !ciphertext) {
    return ''
  }

  const raw = Buffer.from(ciphertext, 'hex')
  if (raw.length <= 16) {
    return ''
  }

  const key = deriveSecretKey(secretKey)
  const decipher = createDecipheriv('aes-256-gcm', key, Buffer.from(nonce, 'hex'))
  decipher.setAuthTag(raw.subarray(raw.length - 16))
  const payloadBuffer = raw.subarray(0, raw.length - 16)
  return Buffer.concat([decipher.update(payloadBuffer), decipher.final()]).toString('utf8')
}

export function generateTotpSecret(): string {
  return base32Encode(randomBytes(20))
}

export function buildTotpOtpauthUri(input: {
  issuer: string
  accountName: string
  secret: string
}): string {
  const issuer = normalizeText(input.issuer) || 'PST Mail Explorer'
  const accountName = normalizeText(input.accountName)
  const secret = normalizeText(input.secret)
  const label = encodeURIComponent(`${issuer}:${accountName}`)
  const query = new URLSearchParams({
    secret,
    issuer,
    algorithm: 'SHA1',
    digits: String(TOTP_DIGITS),
    period: String(TOTP_PERIOD_SECONDS)
  })
  return `otpauth://totp/${label}?${query.toString()}`
}

function hotp(secret: string, counter: number): string {
  const key = base32Decode(secret)
  const counterBuffer = Buffer.alloc(8)
  const high = Math.floor(counter / 0x100000000)
  const low = counter >>> 0
  counterBuffer.writeUInt32BE(high >>> 0, 0)
  counterBuffer.writeUInt32BE(low, 4)
  const hmac = createHmac('sha1', key).update(counterBuffer).digest()
  const offset = hmac[hmac.length - 1] & 0x0f
  const code = ((hmac.readUInt32BE(offset) & 0x7fffffff) % 10 ** TOTP_DIGITS)
    .toString()
    .padStart(TOTP_DIGITS, '0')
  return code
}

export function generateTotpCode(secret: string, timestamp = Date.now()): string {
  const counter = Math.floor(timestamp / 1000 / TOTP_PERIOD_SECONDS)
  return hotp(secret, counter)
}

export function verifyTotpCode(
  secret: string,
  code: string,
  timestamp = Date.now(),
  window = 1
): boolean {
  const normalizedCode = normalizeCode(code)
  if (!/^\d{6}$/.test(normalizedCode)) {
    return false
  }

  const counter = Math.floor(timestamp / 1000 / TOTP_PERIOD_SECONDS)
  for (let offset = -window; offset <= window; offset += 1) {
    if (hotp(secret, counter + offset) === normalizedCode) {
      return true
    }
  }

  return false
}

export function generateRecoveryCodes(count = 8): string[] {
  const codes = new Set<string>()
  while (codes.size < Math.max(1, count)) {
    const raw = randomBytes(5).toString('hex').toUpperCase()
    codes.add(`${raw.slice(0, 5)}-${raw.slice(5, 10)}`)
  }
  return [...codes]
}

export function hashRecoveryCode(code: string, salt: string): string {
  return scryptSync(normalizeCode(code), String(salt ?? ''), 64).toString('hex')
}

export function verifyRecoveryCode(code: string, salt: string, hash: string): boolean {
  try {
    const expected = Buffer.from(String(hash ?? ''), 'hex')
    if (!expected.length) {
      return false
    }
    const actual = scryptSync(normalizeCode(code), String(salt ?? ''), expected.length)
    return expected.length === actual.length && timingSafeEqual(expected, actual)
  } catch {
    return false
  }
}

export function buildRecoveryCodeRecords(codes: string[]): RecoveryCodeRecord[] {
  return codes.map((code) => {
    const salt = randomBytes(16).toString('hex')
    return {
      salt,
      hash: hashRecoveryCode(code, salt),
      usedAt: ''
    }
  })
}
