import { API_ROUTES, toOpenApiPath } from './apiRoutes'
import { EXTRACTION_FIELD_GROUPS } from './extraction'

export interface BuildOpenApiOptions {
  version: string
  reviewStorageMode: 'memory' | 'mongo'
}

function jsonResponse(schema: Record<string, unknown>, description = 'OK'): Record<string, unknown> {
  return {
    description,
    content: {
      'application/json': {
        schema
      }
    }
  }
}

function pstCatalogEntrySchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['fileName', 'size', 'modifiedAt'],
    properties: {
      fileName: { type: 'string' },
      size: { type: 'integer' },
      modifiedAt: { type: ['string', 'null'] }
    }
  }
}

function pstCatalogScopeSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['scopePath', 'scopeLabel', 'fileCount', 'files'],
    properties: {
      scopePath: { type: 'string' },
      scopeLabel: { type: 'string' },
      fileCount: { type: 'integer' },
      files: {
        type: 'array',
        items: pstCatalogEntrySchema()
      }
    }
  }
}

function pstCatalogResponseSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['rootPath', 'rootExists', 'scopes', 'scopePath', 'scopeLabel', 'files', 'message'],
    properties: {
      rootPath: { type: 'string' },
      rootExists: { type: 'boolean' },
      message: { type: 'string' },
      scopePath: { type: 'string' },
      scopeLabel: { type: 'string' },
      scopes: {
        type: 'array',
        items: pstCatalogScopeSchema()
      },
      files: {
        type: 'array',
        items: pstCatalogEntrySchema()
      }
    }
  }
}

function errorResponse(statusCode: number, description: string): Record<string, unknown> {
  return {
    [String(statusCode)]: jsonResponse(
      {
        type: 'object',
        required: ['error'],
        properties: {
          error: { type: 'string' }
        },
        additionalProperties: true
      },
      description
    )
  }
}

function messageSummarySchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    properties: {
      id: { type: 'string' },
      descriptorId: { type: 'string' },
      folderId: { type: 'string' },
      folderPath: { type: 'string' },
      order: { type: 'integer' },
      messageClass: { type: 'string' },
      kind: { type: 'string' },
      subject: { type: 'string' },
      senderName: { type: 'string' },
      senderEmailAddress: { type: 'string' },
      recipientText: { type: 'string' },
      displayTo: { type: 'string' },
      displayCC: { type: 'string' },
      displayBCC: { type: 'string' },
      resolvedDisplayTo: { type: 'string' },
      resolvedDisplayCC: { type: 'string' },
      resolvedDisplayBCC: { type: 'string' },
      originalSubject: { type: 'string' },
      clientSubmitTime: { type: ['string', 'null'] },
      creationTime: { type: ['string', 'null'] },
      modificationTime: { type: ['string', 'null'] },
      messageDeliveryTime: { type: ['string', 'null'] },
      sortDate: { type: ['string', 'null'] },
      sortDateMs: { type: ['integer', 'null'] },
      importance: { type: 'integer' },
      hasAttachments: { type: 'boolean' },
      isRead: { type: 'boolean' },
      isMailLike: { type: 'boolean' },
      review: {
        type: 'object',
        additionalProperties: true,
        properties: {
          flagged: { type: 'boolean' },
          tags: {
            type: 'array',
            items: { type: 'string' }
          },
          createdAt: { type: 'string' },
          updatedAt: { type: 'string' }
        }
      }
    }
  }
}

function reviewStateSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['flagged', 'tags', 'createdAt', 'updatedAt'],
    properties: {
      flagged: { type: 'boolean' },
      tags: {
        type: 'array',
        items: { type: 'string' }
      },
      createdAt: { type: 'string' },
      updatedAt: { type: 'string' }
    }
  }
}

function reviewRecordSchema(): Record<string, unknown> {
  return {
    allOf: [
      reviewStateSchema(),
      {
        type: 'object',
        additionalProperties: true,
        properties: {
          mailboxKey: { type: 'string' },
          fileName: { type: 'string' },
          messageId: { type: 'string' },
          descriptorId: { type: 'string' },
          folderId: { type: 'string' },
          folderPath: { type: 'string' },
          messageClass: { type: 'string' },
          kind: { type: 'string' },
          isMailLike: { type: 'boolean' },
          subject: { type: 'string' },
          senderName: { type: 'string' },
          senderEmailAddress: { type: 'string' },
          displayTo: { type: 'string' },
          displayCC: { type: 'string' },
          displayBCC: { type: 'string' },
          resolvedDisplayTo: { type: 'string' },
          resolvedDisplayCC: { type: 'string' },
          resolvedDisplayBCC: { type: 'string' }
        }
      }
    ]
  }
}

function hiddenRuleSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['filterId', 'kind', 'value', 'label', 'createdAt', 'updatedAt'],
    properties: {
      filterId: { type: 'string' },
      kind: { type: 'string', enum: ['address', 'subject'] },
      value: { type: 'string' },
      label: { type: 'string' },
      createdAt: { type: 'string' },
      updatedAt: { type: 'string' }
    }
  }
}

function authUserSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['username'],
    properties: {
      username: { type: 'string' }
    }
  }
}

function authManagedUserSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['username', 'createdAt'],
    properties: {
      username: { type: 'string' },
      createdAt: { type: 'string' }
    }
  }
}

function authUsersResponseSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['users'],
    properties: {
      users: {
        type: 'array',
        items: authManagedUserSchema()
      }
    }
  }
}

function authUserCreateRequestSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['username', 'password'],
    properties: {
      username: { type: 'string' },
      password: { type: 'string' }
    }
  }
}

function authStatusSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['authenticated', 'enabled', 'canManageUsers', 'user', 'expiresAt'],
    properties: {
      authenticated: { type: 'boolean' },
      enabled: { type: 'boolean' },
      canManageUsers: { type: 'boolean' },
      user: {
        nullable: true,
        ...authUserSchema()
      },
      expiresAt: {
        type: 'string',
        nullable: true
      }
    }
  }
}

function authLoginRequestSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['username', 'password'],
    properties: {
      username: { type: 'string' },
      password: { type: 'string' }
    }
  }
}

function searchResultItemSchema(): Record<string, unknown> {
  return {
    allOf: [
      messageSummarySchema(),
      {
        type: 'object',
        additionalProperties: true,
        required: ['scopePath', 'scopeLabel', 'fileName', 'mailboxName'],
        properties: {
          scopePath: { type: 'string' },
          scopeLabel: { type: 'string' },
          fileName: { type: 'string' },
          mailboxName: { type: 'string' }
        }
      }
    ]
  }
}

function searchPageSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    required: ['items', 'total', 'page', 'pageSize', 'totalPages', 'query', 'mode', 'mailOnly', 'sort', 'scope', 'scopePath', 'scopeLabel', 'hiddenRules'],
    properties: {
      items: {
        type: 'array',
        items: searchResultItemSchema()
      },
      total: { type: 'integer' },
      page: { type: 'integer' },
      pageSize: { type: 'integer' },
      totalPages: { type: 'integer' },
      query: { type: 'string' },
      mode: { type: 'string', enum: ['and', 'or'] },
      mailOnly: { type: 'boolean' },
      sort: { type: 'string' },
      scope: { type: 'string' },
      scopePath: { type: 'string' },
      scopeLabel: { type: 'string' },
      hiddenRules: {
        type: 'array',
        items: hiddenRuleSchema()
      },
      reviewFilters: { type: 'object', additionalProperties: true }
    }
  }
}

function extractionRecordSchema(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: true,
    properties: {
      summary: messageSummarySchema(),
      participants: { type: 'object', additionalProperties: true },
      routing: { type: 'object', additionalProperties: true },
      dates: { type: 'object', additionalProperties: true },
      content: { type: 'object', additionalProperties: true },
      attachments: {
        type: 'object',
        additionalProperties: true,
        properties: {
          hasAttachments: { type: 'boolean' },
          attachments: {
            type: 'array',
            items: { type: 'object', additionalProperties: true }
          }
        }
      },
      headers: { type: 'object', additionalProperties: true },
      review: reviewStateSchema()
    }
  }
}

function openApiPath(pathTemplate: string): string {
  return toOpenApiPath(pathTemplate)
}

export function buildOpenApiDocument(options: BuildOpenApiOptions): Record<string, unknown> {
  const reviewStorageDescription =
    options.reviewStorageMode === 'mongo'
      ? 'Review state is persisted in MongoDB.'
      : 'Review state is stored in memory until the process restarts.'

  return {
    openapi: '3.0.3',
    info: {
      title: 'PST Review API',
      version: options.version,
      description:
        'Upload-free PST browser, extraction API, and review workflow for flagged/tagged mail items.'
    },
    servers: [{ url: '/' }],
    tags: [
      { name: 'Auth' },
      { name: 'PST catalog' },
      { name: 'PST archive' },
      { name: 'Sessions' },
      { name: 'Extraction' },
      { name: 'Review' }
    ],
    paths: {
      [openApiPath(API_ROUTES.authLogin)]: {
        post: {
          tags: ['Auth'],
          summary: 'Sign in to the viewer and receive a session cookie',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: authLoginRequestSchema()
              }
            }
          },
          responses: {
            200: jsonResponse(authStatusSchema()),
            ...errorResponse(401, 'Invalid username or password')
          }
        }
      },
      [openApiPath(API_ROUTES.authMe)]: {
        get: {
          tags: ['Auth'],
          summary: 'Check the current auth session',
          responses: {
            200: jsonResponse(authStatusSchema()),
            ...errorResponse(401, 'Authentication required')
          }
        }
      },
      [openApiPath(API_ROUTES.authLogout)]: {
        post: {
          tags: ['Auth'],
          summary: 'Clear the current auth session cookie',
          responses: {
            200: jsonResponse(authStatusSchema())
          }
        }
      },
      [openApiPath(API_ROUTES.authUsers)]: {
        get: {
          tags: ['Auth'],
          summary: 'List the configured local viewer users',
          responses: {
            200: jsonResponse(authUsersResponseSchema()),
            ...errorResponse(401, 'Authentication required'),
            ...errorResponse(403, 'Admin access required')
          }
        },
        post: {
          tags: ['Auth'],
          summary: 'Add a new local viewer user',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: authUserCreateRequestSchema()
              }
            }
          },
          responses: {
            200: jsonResponse({ type: 'object', additionalProperties: true, required: ['user'], properties: { user: authManagedUserSchema() } }),
            ...errorResponse(400, 'Username is required'),
            ...errorResponse(401, 'Authentication required'),
            ...errorResponse(403, 'Admin access required'),
            ...errorResponse(409, 'User already exists')
          }
        }
      },
      [openApiPath(API_ROUTES.pstCatalog)]: {
        get: {
          tags: ['PST catalog'],
          summary: 'List case/search scopes and PST/OST files from the project PST folder',
          parameters: [
            {
              name: 'scopePath',
              in: 'query',
              required: false,
              schema: { type: 'string' },
              description: 'Relative PST scope path such as Case1/Search1. Empty selects PST root.'
            }
          ],
          responses: {
            ...errorResponse(400, 'Invalid scope path'),
            ...errorResponse(404, 'PST folder missing'),
            200: jsonResponse(pstCatalogResponseSchema())
          }
        }
      },
      [openApiPath(API_ROUTES.pstRemovedCatalog)]: {
        get: {
          tags: ['PST archive'],
          summary: 'List removed case/search scopes and PST/OST files from PST/_removed',
          parameters: [
            {
              name: 'scopePath',
              in: 'query',
              required: false,
              schema: { type: 'string' },
              description: 'Relative removed scope path such as Case1/Search1.'
            }
          ],
          responses: {
            ...errorResponse(400, 'Invalid scope path'),
            200: jsonResponse(pstCatalogResponseSchema())
          }
        }
      },
      [openApiPath(API_ROUTES.pstOpen)]: {
        post: {
          tags: ['PST catalog'],
          summary: 'Open a PST/OST file from a PST case/search folder',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  required: ['fileName'],
                  properties: {
                    scopePath: {
                      type: 'string',
                      description: 'Relative PST scope path. Use an empty string for PST root.'
                    },
                    fileName: { type: 'string' }
                  }
                }
              }
            }
          },
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'scopePath', 'scopeLabel', 'fileName', 'summary', 'tree'],
              properties: {
                sessionId: { type: 'string' },
                scopePath: { type: 'string' },
                scopeLabel: { type: 'string' },
                fileName: { type: 'string' },
                summary: { type: 'object', additionalProperties: true },
                tree: { type: 'object', additionalProperties: true }
              }
            }),
            ...errorResponse(400, 'Invalid mailbox or scope path')
          }
        }
      },
      [openApiPath(API_ROUTES.pstRemove)]: {
        post: {
          tags: ['PST archive'],
          summary: 'Move a PST/OST out of the active catalog into PST/_removed',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  required: ['fileName'],
                  properties: {
                    scopePath: {
                      type: 'string',
                      description: 'Relative PST scope path such as Case1/Search1.'
                    },
                    fileName: { type: 'string' }
                  }
                }
              }
            }
          },
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['removed', 'closedSessionIds'],
              properties: {
                removed: {
                  type: 'object',
                  additionalProperties: true,
                  properties: {
                    sourcePath: { type: 'string' },
                    destinationPath: { type: 'string' },
                    scopePath: { type: 'string' },
                    scopeLabel: { type: 'string' },
                    fileName: { type: 'string' }
                  }
                },
                closedSessionIds: {
                  type: 'array',
                  items: { type: 'string' }
                }
              }
            }),
            ...errorResponse(400, 'Invalid mailbox or scope path'),
            ...errorResponse(404, 'Mailbox not found'),
            ...errorResponse(409, 'Mailbox conflict')
          }
        }
      },
      [openApiPath(API_ROUTES.pstRestore)]: {
        post: {
          tags: ['PST archive'],
          summary: 'Restore a removed PST/OST back into the active catalog',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  required: ['fileName'],
                  properties: {
                    scopePath: {
                      type: 'string',
                      description: 'Relative removed scope path such as Case1/Search1.'
                    },
                    fileName: { type: 'string' }
                  }
                }
              }
            }
          },
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['restored'],
              properties: {
                restored: {
                  type: 'object',
                  additionalProperties: true,
                  properties: {
                    sourcePath: { type: 'string' },
                    destinationPath: { type: 'string' },
                    scopePath: { type: 'string' },
                    scopeLabel: { type: 'string' },
                    fileName: { type: 'string' }
                  }
                }
              }
            }),
            ...errorResponse(400, 'Invalid mailbox or scope path'),
            ...errorResponse(404, 'Mailbox not found'),
            ...errorResponse(409, 'Mailbox conflict')
          }
        }
      },
      [openApiPath(API_ROUTES.search)]: {
        get: {
          tags: ['Extraction'],
          summary: 'Search across the selected PST, selected search folder, or all case/search folders',
          parameters: [
            {
              name: 'scope',
              in: 'query',
              required: false,
              schema: { type: 'string', enum: ['pst', 'search', 'all'] },
              description: 'Search within the selected PST, selected search folder, or all cases/searches.'
            },
            { name: 'scopePath', in: 'query', required: false, schema: { type: 'string' } },
            { name: 'sessionId', in: 'query', required: false, schema: { type: 'string' } },
            {
              name: 'query',
              in: 'query',
              required: false,
              schema: { type: 'string' },
              description:
                'Search terms. Use + in the query for AND or | in the query for OR. Quoted phrases stay exact.'
            },
            {
              name: 'mode',
              in: 'query',
              required: false,
              schema: { type: 'string', enum: ['and', 'or'] },
              description: 'Legacy fallback if the query does not include + or |.'
            },
            { name: 'mailOnly', in: 'query', required: false, schema: { type: 'boolean' } },
            { name: 'sort', in: 'query', required: false, schema: { type: 'string', enum: ['date-desc', 'order'] } },
            { name: 'page', in: 'query', required: false, schema: { type: 'integer', minimum: 1 } },
            { name: 'pageSize', in: 'query', required: false, schema: { type: 'integer', minimum: 1, maximum: 200 } },
            { name: 'reviewFlagged', in: 'query', required: false, schema: { type: 'boolean' } },
            { name: 'reviewTagged', in: 'query', required: false, schema: { type: 'boolean' } },
            { name: 'reviewTag', in: 'query', required: false, schema: { type: 'string' } }
          ],
          responses: {
            ...errorResponse(404, 'Search scope not found'),
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['scope', 'scopePath', 'scopeLabel', 'page'],
              properties: {
                scope: { type: 'string' },
                scopePath: { type: 'string' },
                scopeLabel: { type: 'string' },
                page: searchPageSchema()
              }
            })
          }
        }
      },
      [openApiPath(API_ROUTES.searchIndexRefresh)]: {
        post: {
          tags: ['Extraction'],
          summary: 'Rebuild the cached search index from the PST catalog',
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['summary', 'hiddenRules'],
              properties: {
                summary: {
                  type: 'object',
                  additionalProperties: true
                },
                hiddenRules: {
                  type: 'array',
                  items: hiddenRuleSchema()
                }
              }
            })
          }
        }
      },
      [openApiPath(API_ROUTES.searchFilters)]: {
        get: {
          tags: ['Extraction'],
          summary: 'List hidden search filters',
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['items'],
              properties: {
                items: {
                  type: 'array',
                  items: hiddenRuleSchema()
                }
              }
            })
          }
        },
        post: {
          tags: ['Extraction'],
          summary: 'Create a hidden search filter',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  required: ['kind', 'value'],
                  properties: {
                    kind: { type: 'string', enum: ['address', 'subject'] },
                    value: { type: 'string' },
                    label: { type: 'string' }
                  }
                }
              }
            }
          },
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['rule'],
              properties: {
                rule: hiddenRuleSchema()
              }
            }),
            ...errorResponse(400, 'Invalid filter payload')
          }
        }
      },
      [openApiPath(API_ROUTES.searchFilter)]: {
        delete: {
          tags: ['Extraction'],
          summary: 'Delete a hidden search filter',
          parameters: [
            { name: 'filterId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['deleted'],
              properties: {
                deleted: { type: 'boolean' }
              }
            }),
            ...errorResponse(400, 'Invalid filter id')
          }
        }
      },
      [openApiPath(API_ROUTES.sessionSummary)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Load the summary for a mailbox session',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'summary'],
              properties: {
                sessionId: { type: 'string' },
                summary: { type: 'object', additionalProperties: true }
              }
            }),
            ...errorResponse(404, 'Session not found')
          }
        }
      },
      [openApiPath(API_ROUTES.sessionTree)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Load the folder tree for a mailbox session',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'tree'],
              properties: {
                sessionId: { type: 'string' },
                tree: { type: 'object', additionalProperties: true }
              }
            }),
            ...errorResponse(404, 'Session not found')
          }
        }
      },
      [openApiPath(API_ROUTES.folderMessages)]: {
        get: {
          tags: ['Sessions'],
          summary: 'List messages in a folder',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'folderId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'q', in: 'query', schema: { type: 'string' } },
            { name: 'mailOnly', in: 'query', schema: { type: 'boolean' } },
            { name: 'sort', in: 'query', schema: { type: 'string', enum: ['date-desc', 'order'] } },
            { name: 'page', in: 'query', schema: { type: 'integer', minimum: 1 } },
            { name: 'pageSize', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200 } },
            { name: 'reviewFlagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTag', in: 'query', schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'page'],
              properties: {
                sessionId: { type: 'string' },
                page: {
                  type: 'object',
                  additionalProperties: true,
                  properties: {
                    folder: { type: 'object', additionalProperties: true },
                    items: {
                      type: 'array',
                      items: messageSummarySchema()
                    },
                    total: { type: 'integer' },
                    page: { type: 'integer' },
                    pageSize: { type: 'integer' },
                    totalPages: { type: 'integer' },
                    query: { type: 'string' },
                    mailOnly: { type: 'boolean' },
                    sort: { type: 'string' },
                    reviewFilters: {
                      type: 'object',
                      additionalProperties: true
                    }
                  }
                }
              }
            }),
            ...errorResponse(404, 'Folder not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageDetail)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Load a message with body, attachments, and review state',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'detail'],
              properties: {
                sessionId: { type: 'string' },
                detail: {
                  type: 'object',
                  additionalProperties: true,
                  properties: {
                    review: reviewStateSchema()
                  }
                }
              }
            }),
            ...errorResponse(404, 'Message not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageExtract)]: {
        get: {
          tags: ['Extraction'],
          summary: 'Extract specific data points for a single message',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } },
            {
              name: 'fields',
              in: 'query',
              schema: {
                type: 'string',
                enum: [...EXTRACTION_FIELD_GROUPS, 'all']
              },
              description: 'Comma-separated extraction groups'
            }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'messageId', 'fields', 'record'],
              properties: {
                sessionId: { type: 'string' },
                messageId: { type: 'string' },
                fields: {
                  type: 'array',
                  items: { type: 'string', enum: [...EXTRACTION_FIELD_GROUPS] }
                },
                record: extractionRecordSchema()
              }
            }),
            ...errorResponse(404, 'Message not found')
          }
        }
      },
      [openApiPath(API_ROUTES.folderExtract)]: {
        get: {
          tags: ['Extraction'],
          summary: 'Extract specific data points for a folder page',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'folderId', in: 'path', required: true, schema: { type: 'string' } },
            {
              name: 'fields',
              in: 'query',
              schema: {
                type: 'string',
                enum: [...EXTRACTION_FIELD_GROUPS, 'all']
              },
              description: 'Comma-separated extraction groups'
            },
            { name: 'q', in: 'query', schema: { type: 'string' } },
            { name: 'mailOnly', in: 'query', schema: { type: 'boolean' } },
            { name: 'sort', in: 'query', schema: { type: 'string', enum: ['date-desc', 'order'] } },
            { name: 'page', in: 'query', schema: { type: 'integer', minimum: 1 } },
            { name: 'pageSize', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200 } },
            { name: 'reviewFlagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTag', in: 'query', schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'folder', 'fields', 'paging', 'items'],
              properties: {
                sessionId: { type: 'string' },
                folder: { type: 'object', additionalProperties: true },
                fields: {
                  type: 'array',
                  items: { type: 'string', enum: [...EXTRACTION_FIELD_GROUPS] }
                },
                paging: {
                  type: 'object',
                  additionalProperties: true
                },
                items: {
                  type: 'array',
                  items: {
                    type: 'object',
                    additionalProperties: true,
                    properties: {
                      messageId: { type: 'string' },
                      record: extractionRecordSchema()
                    }
                  }
                }
              }
            }),
            ...errorResponse(404, 'Folder not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageReview)]: {
        get: {
          tags: ['Review'],
          summary: 'Read the current flag/tag state for a message',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'messageId', 'review'],
              properties: {
                sessionId: { type: 'string' },
                messageId: { type: 'string' },
                review: reviewStateSchema()
              }
            }),
            ...errorResponse(404, 'Message not found')
          }
        },
        patch: {
          tags: ['Review'],
          summary: 'Update the flag/tag state for a message',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  properties: {
                    flagged: { type: 'boolean' },
                    tags: {
                      oneOf: [
                        {
                          type: 'array',
                          items: { type: 'string' }
                        },
                        { type: 'string' }
                      ]
                    }
                  }
                }
              }
            }
          },
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'messageId', 'review'],
              properties: {
                sessionId: { type: 'string' },
                messageId: { type: 'string' },
                review: reviewStateSchema()
              }
            }),
            ...errorResponse(400, 'Invalid review payload'),
            ...errorResponse(404, 'Message not found')
          }
        },
        delete: {
          tags: ['Review'],
          summary: 'Clear the flag/tag state for a message',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'messageId', 'review'],
              properties: {
                sessionId: { type: 'string' },
                messageId: { type: 'string' },
                review: reviewStateSchema()
              }
            }),
            ...errorResponse(404, 'Message not found')
          }
        }
      },
      [openApiPath(API_ROUTES.mailboxReviewQueue)]: {
        get: {
          tags: ['Review'],
          summary: 'List reviewed messages for the current mailbox',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'q', in: 'query', schema: { type: 'string' } },
            { name: 'page', in: 'query', schema: { type: 'integer', minimum: 1 } },
            { name: 'pageSize', in: 'query', schema: { type: 'integer', minimum: 1, maximum: 200 } },
            { name: 'reviewFlagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTagged', in: 'query', schema: { type: 'boolean' } },
            { name: 'reviewTag', in: 'query', schema: { type: 'string' } }
          ],
          responses: {
            200: jsonResponse({
              type: 'object',
              additionalProperties: true,
              required: ['sessionId', 'mailboxKey', 'items'],
              properties: {
                sessionId: { type: 'string' },
                mailboxKey: { type: 'string' },
                total: { type: 'integer' },
                page: { type: 'integer' },
                pageSize: { type: 'integer' },
                totalPages: { type: 'integer' },
                filters: { type: 'object', additionalProperties: true },
                items: {
                  type: 'array',
                  items: reviewRecordSchema()
                }
              }
            }),
            ...errorResponse(404, 'Session not found')
          }
        }
      },
      [openApiPath(API_ROUTES.flaggedBundleExport)]: {
        get: {
          tags: ['Review'],
          summary: 'Download a ZIP bundle of flagged mail and appointment items',
          parameters: [
            {
              name: 'scope',
              in: 'query',
              schema: { type: 'string', enum: ['all', 'search', 'pst'] }
            },
            { name: 'scopePath', in: 'query', schema: { type: 'string' } },
            { name: 'sessionId', in: 'query', schema: { type: 'string' } }
          ],
          responses: {
            200: {
              description: 'ZIP bundle export',
              content: {
                'application/zip': {
                  schema: { type: 'string', format: 'binary' }
                }
              }
            },
            ...errorResponse(404, 'Session or search scope not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageExportJson)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Download the current message as JSON',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: {
              description: 'JSON export',
              content: {
                'application/json': {
                  schema: { type: 'string' }
                }
              }
            },
            ...errorResponse(404, 'Message not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageExportEml)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Download the current message as EML',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } }
          ],
          responses: {
            200: {
              description: 'EML export',
              content: {
                'message/rfc822': {
                  schema: { type: 'string', format: 'binary' }
                }
              }
            },
            ...errorResponse(404, 'Message not found')
          }
        }
      },
      [openApiPath(API_ROUTES.messageAttachment)]: {
        get: {
          tags: ['Sessions'],
          summary: 'Download a raw attachment',
          parameters: [
            { name: 'sessionId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'messageId', in: 'path', required: true, schema: { type: 'string' } },
            { name: 'attachmentIndex', in: 'path', required: true, schema: { type: 'integer', minimum: 0 } }
          ],
          responses: {
            200: {
              description: 'Attachment bytes',
              content: {
                'application/octet-stream': {
                  schema: { type: 'string', format: 'binary' }
                }
              }
            },
            ...errorResponse(404, 'Attachment not found')
          }
        }
      }
    },
    components: {
      schemas: {
        AuthStatus: authStatusSchema(),
        AuthUser: authUserSchema(),
        AuthManagedUser: authManagedUserSchema(),
        AuthUsersResponse: authUsersResponseSchema(),
        MessageSummary: messageSummarySchema(),
        ReviewState: reviewStateSchema(),
        ReviewRecord: reviewRecordSchema(),
        ExtractionRecord: extractionRecordSchema()
      }
    },
    'x-review-storage': reviewStorageDescription
  }
}
