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

function parsePositiveInt(value: unknown, fallback: number): number {
  if (value === undefined || value === null || value === '') {
    return fallback
  }

  const parsed = Number.parseInt(String(value), 10)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback
}

export interface PasswordPolicyRecord {
  minLength: number
  requireUppercase: boolean
  requireLowercase: boolean
  requireNumber: boolean
  requireSpecial: boolean
  forgotPasswordAfterFailures: number
  lockoutThreshold: number
  lockoutDurationSeconds: number
  resetTokenTtlMinutes: number
  enforceMfa: boolean
}

export interface PasswordPolicyInput {
  minLength?: number | string
  requireUppercase?: boolean
  requireLowercase?: boolean
  requireNumber?: boolean
  requireSpecial?: boolean
  forgotPasswordAfterFailures?: number | string
  lockoutThreshold?: number | string
  lockoutDurationSeconds?: number | string
  resetTokenTtlMinutes?: number | string
  enforceMfa?: boolean
}

export interface PasswordPolicyView extends PasswordPolicyRecord {}

export const DEFAULT_PASSWORD_POLICY: PasswordPolicyRecord = {
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
}

export function buildPasswordPolicyDefaultsFromEnv(
  env: NodeJS.ProcessEnv = process.env
): PasswordPolicyRecord {
  return {
    minLength: parsePositiveInt(env.PASSWORD_POLICY_MIN_LENGTH, DEFAULT_PASSWORD_POLICY.minLength),
    requireUppercase: parseBoolean(env.PASSWORD_POLICY_REQUIRE_UPPERCASE, DEFAULT_PASSWORD_POLICY.requireUppercase),
    requireLowercase: parseBoolean(env.PASSWORD_POLICY_REQUIRE_LOWERCASE, DEFAULT_PASSWORD_POLICY.requireLowercase),
    requireNumber: parseBoolean(env.PASSWORD_POLICY_REQUIRE_NUMBER, DEFAULT_PASSWORD_POLICY.requireNumber),
    requireSpecial: parseBoolean(env.PASSWORD_POLICY_REQUIRE_SPECIAL, DEFAULT_PASSWORD_POLICY.requireSpecial),
    forgotPasswordAfterFailures: parsePositiveInt(
      env.PASSWORD_POLICY_FORGOT_PASSWORD_AFTER_FAILURES,
      DEFAULT_PASSWORD_POLICY.forgotPasswordAfterFailures
    ),
    lockoutThreshold: parsePositiveInt(
      env.PASSWORD_POLICY_LOCKOUT_THRESHOLD,
      DEFAULT_PASSWORD_POLICY.lockoutThreshold
    ),
    lockoutDurationSeconds: parsePositiveInt(
      env.PASSWORD_POLICY_LOCKOUT_DURATION_SECONDS,
      DEFAULT_PASSWORD_POLICY.lockoutDurationSeconds
    ),
    resetTokenTtlMinutes: parsePositiveInt(
      env.PASSWORD_POLICY_RESET_TOKEN_TTL_MINUTES,
      DEFAULT_PASSWORD_POLICY.resetTokenTtlMinutes
    ),
    enforceMfa: parseBoolean(env.PASSWORD_POLICY_ENFORCE_MFA, DEFAULT_PASSWORD_POLICY.enforceMfa)
  }
}

export function normalizePasswordPolicyInput(input: Partial<PasswordPolicyInput>): PasswordPolicyInput {
  return {
    minLength: input.minLength === undefined ? undefined : parsePositiveInt(input.minLength, DEFAULT_PASSWORD_POLICY.minLength),
    requireUppercase: input.requireUppercase === undefined ? undefined : Boolean(input.requireUppercase),
    requireLowercase: input.requireLowercase === undefined ? undefined : Boolean(input.requireLowercase),
    requireNumber: input.requireNumber === undefined ? undefined : Boolean(input.requireNumber),
    requireSpecial: input.requireSpecial === undefined ? undefined : Boolean(input.requireSpecial),
    forgotPasswordAfterFailures:
      input.forgotPasswordAfterFailures === undefined
        ? undefined
        : parsePositiveInt(input.forgotPasswordAfterFailures, DEFAULT_PASSWORD_POLICY.forgotPasswordAfterFailures),
    lockoutThreshold:
      input.lockoutThreshold === undefined
        ? undefined
        : parsePositiveInt(input.lockoutThreshold, DEFAULT_PASSWORD_POLICY.lockoutThreshold),
    lockoutDurationSeconds:
      input.lockoutDurationSeconds === undefined
        ? undefined
        : parsePositiveInt(input.lockoutDurationSeconds, DEFAULT_PASSWORD_POLICY.lockoutDurationSeconds),
    resetTokenTtlMinutes:
      input.resetTokenTtlMinutes === undefined
        ? undefined
        : parsePositiveInt(input.resetTokenTtlMinutes, DEFAULT_PASSWORD_POLICY.resetTokenTtlMinutes),
    enforceMfa: input.enforceMfa === undefined ? undefined : Boolean(input.enforceMfa)
  }
}

export function mergePasswordPolicy(
  base: PasswordPolicyRecord,
  input: PasswordPolicyInput = {}
): PasswordPolicyRecord {
  return {
    minLength: input.minLength === undefined ? base.minLength : parsePositiveInt(input.minLength, base.minLength),
    requireUppercase: input.requireUppercase === undefined ? base.requireUppercase : Boolean(input.requireUppercase),
    requireLowercase: input.requireLowercase === undefined ? base.requireLowercase : Boolean(input.requireLowercase),
    requireNumber: input.requireNumber === undefined ? base.requireNumber : Boolean(input.requireNumber),
    requireSpecial: input.requireSpecial === undefined ? base.requireSpecial : Boolean(input.requireSpecial),
    forgotPasswordAfterFailures:
      input.forgotPasswordAfterFailures === undefined
        ? base.forgotPasswordAfterFailures
        : parsePositiveInt(input.forgotPasswordAfterFailures, base.forgotPasswordAfterFailures),
    lockoutThreshold:
      input.lockoutThreshold === undefined
        ? base.lockoutThreshold
        : parsePositiveInt(input.lockoutThreshold, base.lockoutThreshold),
    lockoutDurationSeconds:
      input.lockoutDurationSeconds === undefined
        ? base.lockoutDurationSeconds
        : parsePositiveInt(input.lockoutDurationSeconds, base.lockoutDurationSeconds),
    resetTokenTtlMinutes:
      input.resetTokenTtlMinutes === undefined
        ? base.resetTokenTtlMinutes
        : parsePositiveInt(input.resetTokenTtlMinutes, base.resetTokenTtlMinutes),
    enforceMfa: input.enforceMfa === undefined ? base.enforceMfa : Boolean(input.enforceMfa)
  }
}

export function buildPasswordPolicyView(policy: PasswordPolicyRecord): PasswordPolicyView {
  return { ...policy }
}

export function validatePasswordAgainstPolicy(
  password: string,
  policy: PasswordPolicyRecord
): string[] {
  const normalized = String(password ?? '')
  const issues: string[] = []
  if (normalized.length < policy.minLength) {
    issues.push(`Password must be at least ${policy.minLength} characters long.`)
  }
  if (policy.requireUppercase && !/[A-Z]/.test(normalized)) {
    issues.push('Password must include an uppercase letter.')
  }
  if (policy.requireLowercase && !/[a-z]/.test(normalized)) {
    issues.push('Password must include a lowercase letter.')
  }
  if (policy.requireNumber && !/[0-9]/.test(normalized)) {
    issues.push('Password must include a number.')
  }
  if (policy.requireSpecial && !/[^A-Za-z0-9]/.test(normalized)) {
    issues.push('Password must include a special character.')
  }
  return issues
}
