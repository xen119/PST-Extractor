import * as dotenv from 'dotenv'
import * as fs from 'fs'
import * as http from 'http'
import * as https from 'https'
import * as path from 'path'
import { buildOpenApiDocument } from '../src/openApi'
import { MongoAuthUserStore } from '../src/authUsers'
import { createAppSettingsStoreFromEnv, type AppSettingsStore } from '../src/appSettings'
import { createPstReviewApp, type ApiSecurityConfig, type AppAuthConfig } from '../src/pstReviewApp'
import { createFlaggedBundleStoreFromEnv } from '../src/flaggedBundleStore'
import { createReviewStoreFromEnv } from '../src/reviewStore'
import { createSearchIndexStoreFromEnv } from '../src/searchIndex'
import { resolvePstRootDirectory } from '../src/pstCatalog'

interface PackageJson {
  version: string
}

const packageJson = require('../package.json') as PackageJson
const publicDir = path.join(__dirname, 'public')

function loadEnvFile(filePath: string): void {
  if (fs.existsSync(filePath)) {
    dotenv.config({ path: filePath })
  }
}

function parseList(value: string | undefined): string[] {
  if (!value) {
    return []
  }

  return value
    .split(/[,\n;]/g)
    .map((item) => item.trim())
    .filter(Boolean)
}

function parsePositiveInt(value: string | undefined, fallback: number): number {
  if (!value) {
    return fallback
  }
  const parsed = Number.parseInt(value, 10)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback
}

function parseBoolean(value: string | undefined, fallback = false): boolean {
  if (value === undefined) {
    return fallback
  }

  const normalized = value.trim().toLowerCase()
  if (!normalized) {
    return fallback
  }

  return !['0', 'false', 'no', 'off'].includes(normalized)
}

function resolveExamplePath(value: string | undefined): string {
  const normalized = (value || '').trim()
  if (!normalized) {
    return ''
  }

  return path.isAbsolute(normalized) ? normalized : path.resolve(__dirname, normalized)
}

function loadHttpsServerOptions(): https.ServerOptions | null {
  if (!parseBoolean(process.env.HTTPS_ENABLED)) {
    return null
  }

  const passphrase = (process.env.HTTPS_PASSPHRASE || '').trim()
  const pfxPath = resolveExamplePath(process.env.HTTPS_PFX_FILE)
  if (pfxPath) {
    if (!fs.existsSync(pfxPath)) {
      throw new Error(`HTTPS_PFX_FILE does not exist: ${pfxPath}`)
    }

    const options: https.ServerOptions = {
      pfx: fs.readFileSync(pfxPath)
    }

    if (passphrase) {
      options.passphrase = passphrase
    }

    return options
  }

  const certPath = resolveExamplePath(process.env.HTTPS_CERT_FILE)
  const keyPath = resolveExamplePath(process.env.HTTPS_KEY_FILE)
  if (!certPath || !keyPath) {
    throw new Error(
      'HTTPS_ENABLED is true, but no TLS certificate was configured. Set HTTPS_PFX_FILE or both HTTPS_CERT_FILE and HTTPS_KEY_FILE.'
    )
  }

  if (!fs.existsSync(certPath)) {
    throw new Error(`HTTPS_CERT_FILE does not exist: ${certPath}`)
  }
  if (!fs.existsSync(keyPath)) {
    throw new Error(`HTTPS_KEY_FILE does not exist: ${keyPath}`)
  }

  const options: https.ServerOptions = {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath)
  }

  const caPath = resolveExamplePath(process.env.HTTPS_CA_FILE)
  if (caPath) {
    if (!fs.existsSync(caPath)) {
      throw new Error(`HTTPS_CA_FILE does not exist: ${caPath}`)
    }
    options.ca = fs.readFileSync(caPath)
  }

  if (passphrase) {
    options.passphrase = passphrase
  }

  return options
}

function getFallbackRequestInfo(req: any) {
  const headers = req.headers || {}
  const forwardedFor = typeof headers['x-forwarded-for'] === 'string' ? headers['x-forwarded-for'] : ''
  return {
    origin: typeof headers.origin === 'string' ? headers.origin : '',
    referer: typeof headers.referer === 'string' ? headers.referer : '',
    ip: forwardedFor.split(',')[0].trim() || req.ip || req.socket?.remoteAddress || '',
    method: req.method || '',
    url: req.originalUrl || req.url || '',
    contentType: typeof headers['content-type'] === 'string' ? headers['content-type'] : '',
    tenantId: typeof headers['x-tenantid'] === 'string' ? headers['x-tenantid'] : ''
  }
}

loadEnvFile(path.join(__dirname, '.env'))
loadEnvFile(path.join(__dirname, '..', '.env'))

const host = process.env.HOST || '127.0.0.1'
const port = Number(process.env.PORT || 3030)
const httpsPort = parsePositiveInt(process.env.HTTPS_PORT, port)
const auditLogDir = path.join(__dirname, 'logs')
const pstRootDir = resolvePstRootDirectory(process.env.PST_ROOT_DIR, __dirname)
const webChecks = {
  getRequestInfo: getFallbackRequestInfo
}
const m365Auth = {
  CheckTokens: async (_req: any, res: any) => {
    return res.status(500).json({
      success: false,
      message: 'M365 auth middleware is not available on this host.'
    })
  }
}

const apiSecurity: ApiSecurityConfig = {
  webChecks,
  m365Auth,
  bypassIps: parseList(process.env.M365_AUTH_BYPASS_IPS),
  allowedOrigins: parseList(process.env.CORS_ALLOWED_ORIGINS)
}

const auth: AppAuthConfig = {
  username: (process.env.AUTH_USERNAME || 'admin').trim() || 'admin',
  password: (process.env.AUTH_PASSWORD || 'pst-extractor').trim() || 'pst-extractor',
  sessionTtlMinutes: parsePositiveInt(process.env.AUTH_SESSION_TTL_MINUTES, 180),
  inviteTtlMinutes: parsePositiveInt(process.env.AUTH_INVITE_TTL_MINUTES, 24 * 60),
  mfaIssuer: (process.env.AUTH_MFA_ISSUER || 'PST Mail Explorer').trim() || 'PST Mail Explorer',
  publicBaseUrl: (process.env.PUBLIC_BASE_URL || '').trim()
}

let server: http.Server | https.Server | null = null
let reviewStore = null as Awaited<ReturnType<typeof createReviewStoreFromEnv>> | null
let searchIndexStore = null as Awaited<ReturnType<typeof createSearchIndexStoreFromEnv>> | null
let flaggedBundleStore = null as Awaited<ReturnType<typeof createFlaggedBundleStoreFromEnv>> | null
let authUserStore = null as Awaited<ReturnType<typeof MongoAuthUserStore.connect>> | null
let appSettingsStore: AppSettingsStore | null = null
let shuttingDown = false

async function shutdown(exitCode = 0): Promise<void> {
  if (shuttingDown) {
    return
  }
  shuttingDown = true

  try {
    if (server) {
      await new Promise<void>((resolve, reject) => {
        server!.close((error) => {
          if (error) {
            reject(error)
            return
          }
          resolve()
        })
      })
    }
  } catch (error) {
    console.error(error)
  }

  try {
    if (reviewStore) {
      await reviewStore.close()
    }
  } catch (error) {
    console.error(error)
  }

  try {
    if (searchIndexStore) {
      await searchIndexStore.close()
    }
  } catch (error) {
    console.error(error)
  }

  try {
    if (flaggedBundleStore) {
      await flaggedBundleStore.close()
    }
  } catch (error) {
    console.error(error)
  }

  try {
    if (authUserStore) {
      await authUserStore.close()
    }
  } catch (error) {
    console.error(error)
  }

  try {
    if (appSettingsStore) {
      await appSettingsStore.close()
    }
  } catch (error) {
    console.error(error)
  }

  if (exitCode !== 0) {
    process.exitCode = exitCode
  }
}

async function main(): Promise<void> {
  reviewStore = await createReviewStoreFromEnv(process.env)
  searchIndexStore = await createSearchIndexStoreFromEnv(process.env)
  flaggedBundleStore = await createFlaggedBundleStoreFromEnv(process.env)
  appSettingsStore = await createAppSettingsStoreFromEnv(process.env)
  const mongoUri = String(process.env.MONGODB_URI || '').trim()
  if (mongoUri) {
    authUserStore = await MongoAuthUserStore.connect(
      mongoUri,
      process.env.MONGODB_DB || 'pst-extractor',
      [{ username: auth.username, password: auth.password }]
    )
  }
  const openApiSpec = buildOpenApiDocument({
    version: packageJson.version,
    reviewStorageMode: reviewStore.kind
  })

  const app = createPstReviewApp({
    publicDir,
    pstRootDir,
    reviewStore,
    searchIndexStore,
    flaggedBundleStore,
    openApiSpec,
    auth,
    auditLogDir,
    appSettingsStore,
    authUserStore: authUserStore || undefined,
    apiSecurity
  })

  if (parseBoolean(process.env.SEARCH_INDEX_REFRESH_ON_STARTUP, true)) {
    const refreshCoordinator = app.get('searchIndexRefreshCoordinator') as {
      start: (source: 'mailboxes' | 'items', trigger: 'startup' | 'manual') => Promise<unknown>
    }
    void refreshCoordinator
      .start('mailboxes', 'startup')
      .then(() => {
        console.log('Mailbox search index startup refresh started.')
      })
      .catch((error) => {
        console.error('Unable to start mailbox search index startup refresh:', error)
      })
  }

  const httpsOptions = loadHttpsServerOptions()
  const listenerProtocol = httpsOptions ? 'https' : 'http'
  const listenerPort = httpsOptions ? httpsPort : port
  server = httpsOptions ? https.createServer(httpsOptions, app) : http.createServer(app)
  server.listen(listenerPort, host, () => {
    console.log(`PST Mail Explorer running at ${listenerProtocol}://${host}:${listenerPort}`)
  })

  process.once('SIGINT', () => {
    void shutdown(0)
  })
  process.once('SIGTERM', () => {
    void shutdown(0)
  })
}

main().catch((error) => {
  console.error(error)
  void shutdown(1)
})
