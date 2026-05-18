import * as dotenv from 'dotenv'
import * as fs from 'fs'
import * as http from 'http'
import * as path from 'path'
import { buildOpenApiDocument } from '../src/openApi'
import { createPstReviewApp, type ApiSecurityConfig } from '../src/pstReviewApp'
import { createReviewStoreFromEnv } from '../src/reviewStore'
import {
  createSearchIndexStoreFromEnv,
  refreshSearchIndexFromCatalog
} from '../src/searchIndex'
import { getDefaultPstRootDirectory } from '../src/pstCatalog'

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

function loadOptionalModule<T>(modulePath: string, fallback: T, label: string): T {
  try {
    return require(modulePath) as T
  } catch (error) {
    console.warn(`Unable to load ${label} from ${modulePath}:`, error)
    return fallback
  }
}

loadEnvFile(path.join(__dirname, '.env'))
loadEnvFile(path.join(__dirname, '..', '.env'))

const host = process.env.HOST || '127.0.0.1'
const port = Number(process.env.PORT || 3030)
const webChecks = loadOptionalModule(
  'C:\\Coding\\NodeFunctions\\httpSecurity.js',
  {
    getRequestInfo: getFallbackRequestInfo
  },
  'webChecks'
)
const m365Auth = loadOptionalModule(
  'C:\\Coding\\NodeFunctions\\m365-auth.js',
  {
    CheckTokens: async (_req: any, res: any) => {
      return res.status(500).json({
        success: false,
        message: 'M365 auth middleware is not available on this host.'
      })
    }
  },
  'm365Auth'
)

const apiSecurity: ApiSecurityConfig = {
  webChecks,
  m365Auth,
  bypassIps: parseList(process.env.M365_AUTH_BYPASS_IPS),
  allowedOrigins: parseList(process.env.CORS_ALLOWED_ORIGINS)
}

let server: http.Server | null = null
let reviewStore = null as Awaited<ReturnType<typeof createReviewStoreFromEnv>> | null
let searchIndexStore = null as Awaited<ReturnType<typeof createSearchIndexStoreFromEnv>> | null
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

  if (exitCode !== 0) {
    process.exitCode = exitCode
  }
}

async function main(): Promise<void> {
  reviewStore = await createReviewStoreFromEnv(process.env)
  searchIndexStore = await createSearchIndexStoreFromEnv(process.env)
  await refreshSearchIndexFromCatalog(
    getDefaultPstRootDirectory(),
    reviewStore,
    searchIndexStore
  )
  const openApiSpec = buildOpenApiDocument({
    version: packageJson.version,
    reviewStorageMode: reviewStore.kind
  })

  const app = createPstReviewApp({
    publicDir,
    pstRootDir: getDefaultPstRootDirectory(),
    reviewStore,
    searchIndexStore,
    openApiSpec,
    apiSecurity
  })

  server = app.listen(port, host, () => {
    console.log(`PST Mail Explorer running at http://${host}:${port}`)
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
