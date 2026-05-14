import express from 'express'
import * as fs from 'fs'
import * as path from 'path'
import { randomBytes } from 'crypto'
import { API_ROUTES } from './apiRoutes'
import {
  buildFolderExtractionPage,
  buildMessageExtractionRecord,
  buildSummaryExtractionRecord,
  isSummaryOnlyExtraction,
  normalizeExtractionFields,
  type ExtractionFieldGroup
} from './extraction'
import {
  buildReviewContext,
  type ReviewStore
} from './reviewStore'
import type { ReviewState } from './reviewTypes'
import {
  buildEmptyMessageDetail,
  buildFolderTree,
  buildMessageDetail,
  buildMessageDetailFromSession,
  buildSessionSummary,
  exportMessageAsEml,
  exportMessageAsEmlFromSession,
  exportMessageAsJson,
  getAttachmentDownloadBuffer,
  getFolderSummary,
  getMessageSummary,
  isMailLikeSummary,
  listFolderMessages,
  messageMatchesQuery,
  sanitizeFileNameForDownload,
  sortMessageSummaries,
  type FolderMessagePage,
  type MessageDetail,
  type MessageSummary,
  type ViewerSessionIndex
} from './viewer'
import {
  getDefaultPstRootDirectory,
  listPstMailboxFiles,
  openPstMailbox
} from './pstCatalog'

export interface CreatePstReviewAppOptions {
  publicDir: string
  reviewStore: ReviewStore
  openApiSpec: Record<string, unknown>
  pstRootDir?: string
}

interface SessionRecord {
  id: string
  index: ViewerSessionIndex
  filePath: string
  fileName: string
}

interface OpenMailboxRequestBody {
  fileName?: string
}

interface ReviewPatchBody {
  flagged?: boolean
  tags?: unknown
}

interface ListFolderOptions {
  query: string
  mailOnly: boolean
  sort: string
  page: number
  pageSize: number
  reviewFlaggedOnly: boolean
  reviewTaggedOnly: boolean
  reviewTag: string
}

interface SessionResponse {
  sessionId: string
  fileName: string
  summary: ReturnType<typeof buildSessionSummary>
  tree: ReturnType<typeof buildFolderTree>
}

interface ReviewedMessageSummary extends MessageSummary {
  review: ReviewState
}

interface ReviewedMessageDetail extends MessageDetail {
  review: ReviewState
}

interface ReviewedFolderPage extends FolderMessagePage {
  items: ReviewedMessageSummary[]
  reviewFilters: {
    flaggedOnly: boolean
    taggedOnly: boolean
    tag: string
  }
}

const DEFAULT_PAGE_SIZE = 50
const DEFAULT_DOC_TITLE = 'PST API Documentation'
const DEFAULT_SWAGGER_ASSET_PATH = path.dirname(
  require.resolve('swagger-ui-dist/swagger-ui.css')
)

function createSessionId(): string {
  return randomBytes(12).toString('hex')
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function parsePositiveInt(
  value: string | string[] | undefined,
  fallback: number
): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return fallback
  }
  const parsed = Number.parseInt(raw, 10)
  if (Number.isNaN(parsed) || parsed < 1) {
    return fallback
  }
  return parsed
}

function parseBoolean(value: string | string[] | undefined, fallback = false): boolean {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined) {
    return fallback
  }
  return !['0', 'false', 'no', 'off'].includes(raw.trim().toLowerCase())
}

function parseSort(value: string | string[] | undefined): string {
  const raw = Array.isArray(value) ? value[0] : value
  if (!raw) {
    return 'date-desc'
  }
  const normalized = raw.trim().toLowerCase()
  if (normalized === 'order' || normalized === 'folder-order') {
    return 'order'
  }
  return 'date-desc'
}

function parseZeroBasedInt(value: string | string[] | undefined): number {
  const raw = Array.isArray(value) ? value[0] : value
  if (raw === undefined) {
    return -1
  }
  const parsed = Number.parseInt(raw, 10)
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : -1
}

function parseReviewFilters(value: Record<string, string | string[] | undefined>): ListFolderOptions {
  return {
    query: normalizeText(value.q),
    mailOnly: parseBoolean(value.mailOnly, true),
    sort: parseSort(value.sort),
    page: parsePositiveInt(value.page, 1),
    pageSize: parsePositiveInt(value.pageSize, DEFAULT_PAGE_SIZE),
    reviewFlaggedOnly: parseBoolean(value.reviewFlagged, false),
    reviewTaggedOnly: parseBoolean(value.reviewTagged, false),
    reviewTag: normalizeText(value.reviewTag)
  }
}

function createAttachmentBaseUrl(sessionId: string, messageId: string): string {
  return `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(
    messageId
  )}/attachments/`
}

function normalizeReviewState(review: ReviewState | null): ReviewState {
  return (
    review || {
      flagged: false,
      tags: [],
      createdAt: '',
      updatedAt: ''
    }
  )
}

function buildReviewedSummary(
  summary: MessageSummary,
  review: ReviewState | null
): ReviewedMessageSummary {
  return {
    ...summary,
    review: normalizeReviewState(review)
  }
}

function buildReviewedDetail(
  detail: MessageDetail,
  review: ReviewState | null
): ReviewedMessageDetail {
  return {
    ...detail,
    review: normalizeReviewState(review)
  }
}

function buildDocsHtml(): string {
  return `<!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>${DEFAULT_DOC_TITLE}</title>
        <link rel="stylesheet" href="/api/docs/swagger-ui.css" />
        <style>
          html { box-sizing: border-box; }
          *, *::before, *::after { box-sizing: inherit; }
          body { margin: 0; background: #f5f7fb; }
          .swagger-ui .topbar { display: none; }
        </style>
      </head>
      <body>
        <div id="swagger-ui"></div>
        <script src="/api/docs/swagger-ui-bundle.js"></script>
        <script src="/api/docs/swagger-ui-standalone-preset.js"></script>
        <script>
          window.onload = function () {
            window.ui = SwaggerUIBundle({
              url: '/api/openapi.json',
              dom_id: '#swagger-ui',
              deepLinking: true,
              displayRequestDuration: true,
              presets: [SwaggerUIBundle.presets.apis, SwaggerUIBundle.SwaggerUIStandalonePreset],
              layout: 'BaseLayout'
            })
          }
        </script>
      </body>
    </html>`
}

function createAppError(statusCode: number, message: string): Error {
  const error = new Error(message)
  ;(error as Error & { statusCode?: number }).statusCode = statusCode
  return error
}

function toReviewableContext(
  session: SessionRecord,
  summary: MessageSummary
) {
  return buildReviewContext(session.filePath, session.fileName, summary)
}

function getSessionOrThrow(
  sessions: Map<string, SessionRecord>,
  sessionId: string
): SessionRecord {
  const session = sessions.get(sessionId)
  if (!session) {
    throw createAppError(404, 'Session not found')
  }
  return session
}

async function buildMessageDetailResponse(
  session: SessionRecord,
  messageId: string,
  reviewStore: ReviewStore
): Promise<ReviewedMessageDetail> {
  const summary = session.index.messages.get(messageId)
  if (!summary) {
    throw createAppError(404, `Unknown message: ${messageId}`)
  }

  const review = await reviewStore.getReview(session.filePath, messageId)
  if (summary.parseError) {
    return buildReviewedDetail(buildEmptyMessageDetail(summary), review)
  }

  try {
    const detail = await Promise.resolve(
      buildMessageDetailFromSession(session.index, messageId, 1)
    )
    return buildReviewedDetail(
      {
        ...detail,
        attachments: detail.attachments.map((attachment) => ({
          ...attachment,
          downloadUrl: attachment.downloadUrl || createAttachmentBaseUrl(session.id, messageId)
        }))
      },
      review
    )
  } catch (error) {
    const parseError = error instanceof Error ? error.message : String(error)
    return buildReviewedDetail(
      buildEmptyMessageDetail({
        ...summary,
        parseError
      }),
      review
    )
  }
}

async function buildReviewedFolderPage(
  session: SessionRecord,
  folderId: string,
  options: ListFolderOptions,
  reviewStore: ReviewStore
): Promise<ReviewedFolderPage> {
  const folder = getFolderSummary(session.index, folderId)
  const messageIds = [...folder.messageIds]
  const reviewMap = await reviewStore.getMany(session.filePath, messageIds)

  let items = messageIds
    .map((messageId) => session.index.messages.get(messageId))
    .filter((message): message is MessageSummary => Boolean(message))
    .filter((message) => (options.mailOnly ? isMailLikeSummary(message) : true))
    .filter((message) => messageMatchesQuery(message, options.query))

  if (options.reviewFlaggedOnly || options.reviewTaggedOnly || options.reviewTag) {
    items = items.filter((message) => {
      const review = reviewMap.get(message.id) || null
      if (options.reviewFlaggedOnly && !review?.flagged) {
        return false
      }
      if (options.reviewTaggedOnly && (!review || review.tags.length === 0)) {
        return false
      }
      if (options.reviewTag) {
        const tag = options.reviewTag.toLowerCase()
        if (!review || !review.tags.some((value) => value.toLowerCase() === tag)) {
          return false
        }
      }
      return true
    })
  }

  const sorted = sortMessageSummaries(items, options.sort)
  const total = sorted.length
  const totalPages = Math.max(1, Math.ceil(total / options.pageSize))
  const page = Math.min(Math.max(options.page, 1), totalPages)
  const start = (page - 1) * options.pageSize
  const pageItems = sorted.slice(start, start + options.pageSize)
  const pageReviewMap = await reviewStore.getMany(
    session.filePath,
    pageItems.map((item) => item.id)
  )

  return {
    folder,
    items: pageItems.map((item) => buildReviewedSummary(item, pageReviewMap.get(item.id) || null)),
    total,
    page,
    pageSize: options.pageSize,
    totalPages,
    query: options.query,
    mailOnly: options.mailOnly,
    sort: options.sort,
    reviewFilters: {
      flaggedOnly: options.reviewFlaggedOnly,
      taggedOnly: options.reviewTaggedOnly,
      tag: options.reviewTag
    }
  }
}

async function buildFolderPageWithReviews(
  session: SessionRecord,
  folderId: string,
  options: ListFolderOptions,
  reviewStore: ReviewStore
): Promise<ReviewedFolderPage> {
  if (options.reviewFlaggedOnly || options.reviewTaggedOnly || options.reviewTag) {
    return buildReviewedFolderPage(session, folderId, options, reviewStore)
  }

  const basePage = listFolderMessages(session.index, folderId, {
    query: options.query,
    mailOnly: options.mailOnly,
    page: options.page,
    pageSize: options.pageSize,
    sort: options.sort
  })
  const reviewMap = await reviewStore.getMany(
    session.filePath,
    basePage.items.map((item) => item.id)
  )

  return {
    ...basePage,
    items: basePage.items.map((item) => buildReviewedSummary(item, reviewMap.get(item.id) || null)),
    reviewFilters: {
      flaggedOnly: options.reviewFlaggedOnly,
      taggedOnly: options.reviewTaggedOnly,
      tag: options.reviewTag
    }
  }
}

function assertReviewableMessage(summary: MessageSummary): void {
  if (!isMailLikeSummary(summary)) {
    throw createAppError(400, 'Review state is available for mail items only')
  }
}

function responseJson(
  res: express.Response,
  statusCode: number,
  payload: unknown
): void {
  res.status(statusCode).json(payload)
}

function responseBinary(
  res: express.Response,
  statusCode: number,
  contentType: string,
  fileName: string,
  data: Buffer
): void {
  res.status(statusCode)
    .type(contentType)
    .set('Content-Disposition', `attachment; filename="${fileName}"`)
    .set('Content-Length', String(data.length))
    .send(data)
}

function toOpenApiPath(pathTemplate: string): string {
  return pathTemplate.replace(/:([A-Za-z0-9_]+)/g, '{$1}')
}

function createRouteErrorHandler(
  res: express.Response,
  error: unknown
): void {
  const statusCode = typeof error === 'object' && error && 'statusCode' in error
    ? Number((error as { statusCode?: number }).statusCode || 500)
    : 500
  const message = error instanceof Error ? error.message : String(error)
  responseJson(res, statusCode, { error: message })
}

function safeDownloadName(name: string, fallback: string): string {
  return sanitizeFileNameForDownload(name || fallback, fallback)
}

export function createPstReviewApp(options: CreatePstReviewAppOptions): express.Express {
  const app = express()
  const sessions = new Map<string, SessionRecord>()
  const publicDir = options.publicDir
  const pstRootDir = options.pstRootDir || getDefaultPstRootDirectory()
  const reviewStore = options.reviewStore

  app.disable('x-powered-by')
  app.use(express.json({ limit: '2mb' }))
  app.use(express.static(publicDir))

  app.get(API_ROUTES.openApiJson, (_req, res) => {
    responseJson(res, 200, options.openApiSpec)
  })

  app.use(
    `${API_ROUTES.docs}/`,
    express.static(DEFAULT_SWAGGER_ASSET_PATH, {
      index: false
    })
  )
  app.get(API_ROUTES.docs, (_req, res) => {
    res.status(200).type('html').send(buildDocsHtml())
  })

  app.get(API_ROUTES.pstCatalog, (_req, res) => {
    responseJson(res, 200, listPstMailboxFiles(pstRootDir))
  })

  app.post(API_ROUTES.pstOpen, async (req, res) => {
    try {
      const body = (req.body || {}) as OpenMailboxRequestBody
      const fileName = normalizeText(body.fileName)
      if (!fileName) {
        throw createAppError(400, 'Mailbox file name is required')
      }

      const index = openPstMailbox(pstRootDir, fileName)
      const sessionId = createSessionId()
      const record: SessionRecord = {
        id: sessionId,
        index,
        filePath: index.filePath,
        fileName: index.fileName
      }
      sessions.set(sessionId, record)

      const summary = buildSessionSummary(index)
      responseJson(res, 200, {
        sessionId,
        fileName,
        summary,
        tree: buildFolderTree(index)
      } satisfies SessionResponse)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.sessionSummary, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      responseJson(res, 200, {
        sessionId: session.id,
        summary: buildSessionSummary(session.index)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.sessionTree, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      responseJson(res, 200, {
        sessionId: session.id,
        tree: buildFolderTree(session.index)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.folderMessages, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const folderId = normalizeText(req.params.folderId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const page = await buildFolderPageWithReviews(session, folderId, filters, reviewStore)
      responseJson(res, 200, {
        sessionId: session.id,
        page
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageDetail, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      responseJson(res, 200, {
        sessionId: session.id,
        detail
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExtract, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      const fields = normalizeExtractionFields(req.query.fields)
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        fields,
        record: buildMessageExtractionRecord(detail, detail.review, fields)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.folderExtract, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const folderId = normalizeText(req.params.folderId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const fields = normalizeExtractionFields(req.query.fields)
      const page = await buildFolderPageWithReviews(session, folderId, filters, reviewStore)
      const folder = getFolderSummary(session.index, folderId)
      const extraction = buildFolderExtractionPage(
        page,
        new Map(page.items.map((item) => [item.id, item.review])),
        fields,
        (summary, review, fieldList) => {
          if (isSummaryOnlyExtraction(fieldList)) {
            return buildSummaryExtractionRecord(summary, review, fieldList)
          }
          const detail = buildMessageDetailFromSession(session.index, summary.id, 1)
          return buildMessageExtractionRecord(detail, review, fieldList)
        }
      )

      responseJson(res, 200, {
        sessionId: session.id,
        folder: {
          id: folder.id,
          descriptorId: folder.descriptorId,
          displayName: folder.displayName,
          path: folder.path
        },
        fields: extraction.fields,
        paging: extraction.paging,
        items: extraction.items
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExportJson, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const detail = await buildMessageDetailResponse(session, messageId, reviewStore)
      const fileName = `${safeDownloadName(detail.subject || 'message', 'message')}.json`
      responseBinary(
        res,
        200,
        'application/json; charset=utf-8',
        fileName,
        Buffer.from(exportMessageAsJson(detail), 'utf8')
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageExportEml, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const summary = getMessageSummary(session.index, messageId)
      const fileName = `${safeDownloadName(summary.subject || 'message', 'message')}.eml`
      const eml = summary.parseError
        ? exportMessageAsEml(buildReviewedDetail(buildEmptyMessageDetail(summary), null))
        : exportMessageAsEmlFromSession(session.index, messageId)
      responseBinary(
        res,
        200,
        'message/rfc822; charset=utf-8',
        fileName,
        Buffer.from(eml, 'utf8')
      )
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageAttachment, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const attachmentIndex = parseZeroBasedInt(req.params.attachmentIndex)
      if (attachmentIndex < 0) {
        throw createAppError(404, 'Attachment not found')
      }
      const payload = getAttachmentDownloadBuffer(session.index, messageId, attachmentIndex)
      const fileName = safeDownloadName(payload.filename, 'attachment')
      responseBinary(res, 200, payload.contentType, fileName, payload.buffer)
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const review = normalizeReviewState(await reviewStore.getReview(session.filePath, messageId))
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.patch(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      const body = (req.body || {}) as ReviewPatchBody
      if (body.flagged === undefined && body.tags === undefined) {
        throw createAppError(400, 'Provide flagged or tags to update review state')
      }

      const summary = getMessageSummary(session.index, messageId)
      assertReviewableMessage(summary)
      const review = await reviewStore.upsertReview({
        ...toReviewableContext(session, summary),
        flagged: body.flagged,
        tags: body.tags
      })

      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review: normalizeReviewState(review)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.delete(API_ROUTES.messageReview, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const messageId = normalizeText(req.params.messageId)
      await reviewStore.deleteReview(session.filePath, messageId)
      responseJson(res, 200, {
        sessionId: session.id,
        messageId,
        review: normalizeReviewState(null)
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.get(API_ROUTES.mailboxReviewQueue, async (req, res) => {
    try {
      const session = getSessionOrThrow(sessions, req.params.sessionId)
      const filters = parseReviewFilters(req.query as Record<string, string | string[] | undefined>)
      const pageSize = filters.pageSize
      const page = filters.page
      const records = await reviewStore.listReviews(session.filePath, {
        query: filters.query,
        flaggedOnly: filters.reviewFlaggedOnly,
        taggedOnly: filters.reviewTaggedOnly,
        tag: filters.reviewTag
      })
      const total = records.length
      const totalPages = Math.max(1, Math.ceil(total / pageSize))
      const currentPage = Math.min(Math.max(page, 1), totalPages)
      const start = (currentPage - 1) * pageSize
      const items = records.slice(start, start + pageSize)

      responseJson(res, 200, {
        sessionId: session.id,
        mailboxKey: session.filePath,
        total,
        page: currentPage,
        pageSize,
        totalPages,
        filters: {
          query: filters.query,
          flaggedOnly: filters.reviewFlaggedOnly,
          taggedOnly: filters.reviewTaggedOnly,
          tag: filters.reviewTag
        },
        items
      })
    } catch (error) {
      createRouteErrorHandler(res, error)
    }
  })

  app.use(
    (err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
    createRouteErrorHandler(res, err)
    }
  )

  return app
}
