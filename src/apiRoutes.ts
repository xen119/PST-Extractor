export const API_ROUTES = {
  pstCatalog: '/api/psts',
  pstOpen: '/api/psts/open',
  searchIndexRefresh: '/api/search/index/refresh',
  searchFilters: '/api/search/filters',
  searchFilter: '/api/search/filters/:filterId',
  search: '/api/search',
  sessionSummary: '/api/sessions/:sessionId',
  sessionTree: '/api/sessions/:sessionId/tree',
  folderMessages: '/api/sessions/:sessionId/folders/:folderId/messages',
  folderExtract: '/api/sessions/:sessionId/folders/:folderId/messages/extract',
  messageDetail: '/api/sessions/:sessionId/messages/:messageId',
  messageExtract: '/api/sessions/:sessionId/messages/:messageId/extract',
  messageReview: '/api/sessions/:sessionId/messages/:messageId/review',
  messageExportJson: '/api/sessions/:sessionId/messages/:messageId/export.json',
  messageExportEml: '/api/sessions/:sessionId/messages/:messageId/export.eml',
  messageAttachment: '/api/sessions/:sessionId/messages/:messageId/attachments/:attachmentIndex',
  mailboxReviewQueue: '/api/sessions/:sessionId/review',
  flaggedBundleExport: '/api/exports/flagged.zip',
  openApiJson: '/api/openapi.json',
  docs: '/api/docs'
} as const

export function toOpenApiPath(pathTemplate: string): string {
  return pathTemplate.replace(/:([A-Za-z0-9_]+)/g, '{$1}')
}
