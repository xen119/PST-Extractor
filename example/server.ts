import * as dotenv from 'dotenv'
import * as fs from 'fs'
import * as http from 'http'
import * as path from 'path'
import { buildOpenApiDocument } from '../src/openApi'
import { createPstReviewApp } from '../src/pstReviewApp'
import { createReviewStoreFromEnv } from '../src/reviewStore'
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

loadEnvFile(path.join(__dirname, '.env'))
loadEnvFile(path.join(__dirname, '..', '.env'))

const host = process.env.HOST || '127.0.0.1'
const port = Number(process.env.PORT || 3030)

let server: http.Server | null = null
let reviewStore = null as Awaited<ReturnType<typeof createReviewStoreFromEnv>> | null
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

  if (exitCode !== 0) {
    process.exitCode = exitCode
  }
}

async function main(): Promise<void> {
  reviewStore = await createReviewStoreFromEnv(process.env)
  const openApiSpec = buildOpenApiDocument({
    version: packageJson.version,
    reviewStorageMode: reviewStore.kind
  })

  const app = createPstReviewApp({
    publicDir,
    pstRootDir: getDefaultPstRootDirectory(),
    reviewStore,
    openApiSpec
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
