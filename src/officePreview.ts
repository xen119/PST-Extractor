import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { spawn } from 'child_process'
import { pathToFileURL } from 'url'

export interface OfficePreviewInput {
  cacheKey: string
  fileName: string
  contentType?: string
  buffer: Buffer
  previewText?: string
}

export interface OfficePreviewOutput {
  contentType: string
  buffer: Buffer
  fileName: string
}

const OFFICE_CONTENT_TYPES = new Set([
  'application/msword',
  'application/vnd.ms-excel',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
])

const OFFICE_EXTENSIONS = new Set(['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx'])

const previewCache = new Map<string, Promise<OfficePreviewOutput>>()

function isOfficeDocument(contentType?: string, fileName = ''): boolean {
  const normalizedType = (contentType || '').toLowerCase().split(';')[0].trim()
  if (normalizedType && OFFICE_CONTENT_TYPES.has(normalizedType)) {
    return true
  }
  const extension = path.extname(fileName).toLowerCase()
  return OFFICE_EXTENSIONS.has(extension)
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function buildFallbackHtml(fileName: string, previewText = ''): string {
  const safeTitle = escapeHtml(fileName || 'Office document')
  const safeText = escapeHtml(previewText.trim() || 'No text preview is available for this file.')
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${safeTitle}</title>
    <style>
      :root {
        color-scheme: light;
      }
      html, body {
        margin: 0;
        min-height: 100%;
        background: #fff;
        color: #1f2937;
        font-family: Arial, Helvetica, sans-serif;
      }
      body {
        box-sizing: border-box;
        padding: 24px;
      }
      .card {
        max-width: 1200px;
        margin: 0 auto;
        border: 1px solid #dbe3f0;
        border-radius: 18px;
        background: #f8fbff;
        padding: 20px;
      }
      .label {
        margin: 0 0 8px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: #64748b;
      }
      h1 {
        margin: 0 0 12px;
        font-size: 24px;
        line-height: 1.2;
      }
      .note {
        margin: 0 0 16px;
        color: #475569;
        font-size: 14px;
      }
      pre {
        margin: 0;
        white-space: pre-wrap;
        word-break: break-word;
        font-size: 14px;
        line-height: 1.6;
        color: #111827;
      }
    </style>
  </head>
  <body>
    <main class="card">
      <p class="label">Document preview</p>
      <h1>${safeTitle}</h1>
      <p class="note">LibreOffice preview was not available, so the extracted text is shown instead.</p>
      <pre>${safeText}</pre>
    </main>
  </body>
</html>`
}

async function convertOfficeBufferToPdf(input: OfficePreviewInput): Promise<OfficePreviewOutput | null> {
  if (!isOfficeDocument(input.contentType, input.fileName)) {
    return null
  }

  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-office-preview-'))
  try {
    const profileDir = path.join(tempRoot, 'profile')
    fs.mkdirSync(profileDir, { recursive: true })

    const inputExtension = path.extname(input.fileName) || '.bin'
    const safeBaseName = path
      .basename(input.fileName, inputExtension)
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, '_')
      .trim() || 'document'
    const sourcePath = path.join(tempRoot, `${safeBaseName}${inputExtension}`)
    fs.writeFileSync(sourcePath, input.buffer)

    const commands: string[] = Array.from(
      new Set(
        [process.env.LIBREOFFICE_BIN?.trim(), process.env.SOFFICE_BIN?.trim(), 'soffice', 'libreoffice'].filter(
          (value): value is string => Boolean(value)
        )
      )
    )

    let converted = false
    for (const command of commands) {
      converted = await new Promise<boolean>((resolve) => {
        const child = spawn(
          command,
          [
            `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
            '--headless',
            '--invisible',
            '--nologo',
            '--nofirststartwizard',
            '--convert-to',
            'pdf',
            '--outdir',
            tempRoot,
            sourcePath
          ],
          {
            windowsHide: true,
            stdio: ['ignore', 'ignore', 'ignore']
          }
        )
        child.on('error', () => resolve(false))
        child.on('close', (code: number | null) => resolve(code === 0))
      })

      if (converted) {
        break
      }
    }

    if (!converted) {
      return null
    }

    const pdfPath = path.join(tempRoot, `${safeBaseName}.pdf`)
    if (!fs.existsSync(pdfPath)) {
      return null
    }

    return {
      contentType: 'application/pdf',
      buffer: fs.readFileSync(pdfPath),
      fileName: `${safeBaseName}.pdf`
    }
  } catch {
    return null
  } finally {
    try {
      fs.rmSync(tempRoot, { recursive: true, force: true })
    } catch {
      // ignore cleanup failures
    }
  }
}

export async function buildOfficePreview(input: OfficePreviewInput): Promise<OfficePreviewOutput> {
  const cacheKey = input.cacheKey || `${input.fileName}:${input.contentType || ''}`
  const cached = previewCache.get(cacheKey)
  if (cached) {
    return cached
  }

  const promise = (async (): Promise<OfficePreviewOutput> => {
    const converted = await convertOfficeBufferToPdf(input)
    if (converted) {
      return converted
    }
    return {
      contentType: 'text/html; charset=utf-8',
      buffer: Buffer.from(buildFallbackHtml(input.fileName, input.previewText), 'utf8'),
      fileName: input.fileName
    }
  })()

  previewCache.set(cacheKey, promise)
  return promise
}

export function isOfficePreviewable(contentType?: string, fileName = ''): boolean {
  return isOfficeDocument(contentType, fileName)
}
