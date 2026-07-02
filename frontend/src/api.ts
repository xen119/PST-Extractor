import type {
  ActivityLogResponse,
  AuthStatus,
  MessageDetail,
  HiddenRulesResponse,
  InviteAcceptResponse,
  InviteLookupResponse,
  MfaEnrollmentCompleteResponse,
  MfaEnrollmentStartResponse,
  PstCatalogResponse,
  SearchResponse,
  SessionOpenResponse,
  FolderMessagesResponse,
  MessageDetailResponse,
  ReviewUpdateResponse,
  SearchIndexRefreshResponse,
  SearchIndexRefreshSource,
  SearchIndexRefreshStatus,
  SmtpSettingsResponse,
  SmtpTestResponse,
  UserInviteResponse,
  UsersResponse,
  UserInvite,
  ReviewQueueRecord
} from './types'

export class ApiError extends Error {
  statusCode: number
  payload: unknown

  constructor(message: string, statusCode: number, payload: unknown) {
    super(message)
    this.statusCode = statusCode
    this.payload = payload
  }
}

type JsonHeaders = Record<string, string>

async function requestJson<T>(url: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(url, {
    credentials: 'include',
    ...init,
    headers: {
      Accept: 'application/json',
      ...(init.headers as JsonHeaders | undefined)
    }
  })

  const contentType = response.headers.get('content-type') || ''
  const payload = contentType.includes('application/json') ? await response.json() : await response.text()

  if (!response.ok) {
    const message =
      (payload && typeof payload === 'object' && 'error' in payload && String((payload as Record<string, unknown>).error)) ||
      (typeof payload === 'string' ? payload : '') ||
      response.statusText ||
      'Request failed'
    throw new ApiError(message, response.status, payload)
  }

  return payload as T
}

function buildQuery(params: Record<string, string | number | boolean | undefined | null>): string {
  const query = new URLSearchParams()
  Object.entries(params).forEach(([key, value]) => {
    if (value == null || value === '') {
      return
    }
    query.set(key, String(value))
  })
  return query.toString()
}

export const api = {
  auth: {
    me: () => requestJson<AuthStatus>('/api/auth/me', { cache: 'no-store' }),
    login: (username: string, password: string) =>
      requestJson<AuthStatus>('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      }),
    logout: () => requestJson<AuthStatus>('/api/auth/logout', { method: 'POST' }),
    inviteLookup: (token: string) =>
      requestJson<InviteLookupResponse>(`/api/auth/invites/${encodeURIComponent(token)}`, {
        cache: 'no-store'
      }),
    inviteAccept: (token: string, password: string) =>
      requestJson<InviteAcceptResponse>(`/api/auth/invites/${encodeURIComponent(token)}/accept`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password })
      }),
    mfaChallenge: (code: string) =>
      requestJson<AuthStatus>('/api/auth/mfa/challenge', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code })
      }),
    mfaEnrollmentStart: () =>
      requestJson<MfaEnrollmentStartResponse>('/api/auth/mfa/enrollment/start', { method: 'POST' }),
    mfaEnrollmentComplete: (code: string) =>
      requestJson<MfaEnrollmentCompleteResponse>('/api/auth/mfa/enrollment/complete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code })
      }),
    users: () => requestJson<UsersResponse>('/api/auth/users', { cache: 'no-store' }),
    inviteUser: (username: string, recipientEmail: string) =>
      requestJson<UserInviteResponse>('/api/auth/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, recipientEmail })
      }),
    deleteUser: (username: string) =>
      requestJson<{ user: UserInvite }>(`/api/auth/users/${encodeURIComponent(username)}`, {
        method: 'DELETE'
      }),
    resendInvite: (username: string) =>
      requestJson<UserInviteResponse>(`/api/auth/users/${encodeURIComponent(username)}/invite/resend`, {
        method: 'POST'
      }),
    revokeInvite: (username: string) =>
      requestJson<{ user: UserInvite }>(`/api/auth/users/${encodeURIComponent(username)}/invite`, {
        method: 'DELETE'
      }),
    resetMfa: (username: string) =>
      requestJson<{ user: UserInvite }>(`/api/auth/users/${encodeURIComponent(username)}/mfa/reset`, {
        method: 'POST'
      }),
    setMfaEnforced: (username: string, enforced: boolean) =>
      requestJson<{ user: UserInvite }>(`/api/auth/users/${encodeURIComponent(username)}/mfa/enforce`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enforced })
      }),
    setUserAccess: (username: string, assignedCasePaths: string[]) =>
      requestJson<{ user: UserInvite }>(`/api/auth/users/${encodeURIComponent(username)}/access`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ assignedCasePaths })
      })
  },
  settings: {
    smtpGet: () => requestJson<SmtpSettingsResponse>('/api/settings/smtp', { cache: 'no-store' }),
    smtpPut: (payload: Record<string, unknown>) =>
      requestJson<SmtpSettingsResponse>('/api/settings/smtp', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }),
    smtpTest: (payload: Record<string, unknown>) =>
      requestJson<SmtpTestResponse>('/api/settings/smtp/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      })
  },
  activityLog: {
    list: (limit = 100, username = '') =>
      requestJson<ActivityLogResponse>(
        `/api/activity-log${buildQuery({ limit, username }) ? `?${buildQuery({ limit, username })}` : ''}`,
        { cache: 'no-store' }
      ),
    csvUrl: (username = '') => {
      const query = buildQuery({ username })
      return query ? `/api/activity-log.csv?${query}` : '/api/activity-log.csv'
    }
  },
  workspace: {
    itemsCsvUrl: (params: Record<string, string | number | boolean | undefined>) => {
      const query = buildQuery(params)
      return query ? `/api/exports/items.csv?${query}` : '/api/exports/items.csv'
    },
    clearAllFlags: (params: Record<string, string | number | boolean | undefined>) =>
      requestJson<{
        clearedCount: number
        itemCount: number
        scopePath: string
        scopeLabel: string
      }>(
        `/api/reviews/clear-flags${buildQuery(params) ? `?${buildQuery(params)}` : ''}`,
        { method: 'POST' }
      )
  },
  pst: {
    catalog: (scopePath = '', removed = false) => {
      const query = buildQuery({ scopePath })
      const route = removed ? '/api/psts/removed' : '/api/psts'
      return requestJson<PstCatalogResponse>(query ? `${route}?${query}` : route, {
        cache: 'no-store'
      })
    },
    open: (scopePath: string, fileName: string) =>
      requestJson<SessionOpenResponse>('/api/psts/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ scopePath, fileName })
      }),
    remove: (scopePath: string, fileName: string) =>
      requestJson('/api/psts/remove', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ scopePath, fileName })
      }),
    restore: (scopePath: string, fileName: string) =>
      requestJson('/api/psts/restore', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ scopePath, fileName })
      }),
    refreshSearchIndex: (source: SearchIndexRefreshSource) =>
      requestJson<SearchIndexRefreshResponse>(`/api/search/index/refresh?${buildQuery({ source })}`, {
        method: 'POST'
      }),
    refreshSearchIndexStatus: (source: SearchIndexRefreshSource) =>
      requestJson<SearchIndexRefreshResponse>(`/api/search/index/refresh/status?${buildQuery({ source })}`, {
        cache: 'no-store'
      })
  },
  hiddenFilters: {
    list: () => requestJson<HiddenRulesResponse>('/api/search/filters', { cache: 'no-store' }),
    create: (kind: 'address' | 'subject', value: string, label?: string) =>
      requestJson('/api/search/filters', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ kind, value, label })
      }),
    delete: (filterId: string) =>
      requestJson(`/api/search/filters/${encodeURIComponent(filterId)}`, {
        method: 'DELETE'
      })
  },
  search: (params: {
    scope: 'all' | 'search' | 'pst'
    sourceType?: 'mailbox' | 'teams' | 'sharepoint'
    query: string
    mode: 'and' | 'or'
    page: number
    pageSize: number
    mailOnly: boolean
    sort: string
    reviewFlagged?: boolean
    reviewTagged?: boolean
    reviewTag?: string
    scopePath?: string
    sessionId?: string
  }) => {
    const query = buildQuery({
      scope: params.scope,
      sourceType: params.sourceType,
      query: params.query,
      mode: params.mode,
      page: params.page,
      pageSize: params.pageSize,
      mailOnly: params.mailOnly,
      sort: params.sort,
      reviewFlagged: params.reviewFlagged ? '1' : undefined,
      reviewTagged: params.reviewTagged ? '1' : undefined,
      reviewTag: params.reviewTag,
      scopePath: params.scopePath,
      sessionId: params.sessionId
    })
    return requestJson<SearchResponse>(query ? `/api/search?${query}` : '/api/search', {
      cache: 'no-store'
    })
  },
  item: {
    detail: (itemId: string) =>
      requestJson<{ detail: MessageDetail }>(`/api/items/${encodeURIComponent(itemId)}`, {
        cache: 'no-store'
      }),
    review: (itemId: string) =>
      requestJson<ReviewUpdateResponse>(`/api/items/${encodeURIComponent(itemId)}/review`, {
        cache: 'no-store'
      }),
    updateReview: (itemId: string, payload: { flagged?: boolean; tags?: string[] | string }) =>
      requestJson<ReviewUpdateResponse>(`/api/items/${encodeURIComponent(itemId)}/review`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }),
    clearReview: (itemId: string) =>
      requestJson<ReviewUpdateResponse>(`/api/items/${encodeURIComponent(itemId)}/review`, {
        method: 'DELETE'
      }),
    contentUrl: (itemId: string) => `/api/items/${encodeURIComponent(itemId)}/content`
  },
  session: {
    summary: (sessionId: string) =>
      requestJson<{ sessionId: string; summary: unknown }>(`/api/sessions/${encodeURIComponent(sessionId)}`, {
        cache: 'no-store'
      }),
    tree: (sessionId: string) =>
      requestJson<{ sessionId: string; tree: unknown }>(`/api/sessions/${encodeURIComponent(sessionId)}/tree`, {
        cache: 'no-store'
      }),
    folderMessages: (sessionId: string, folderId: string, params: Record<string, string | number | boolean | undefined>) => {
      const query = buildQuery(params)
      const url = `/api/sessions/${encodeURIComponent(sessionId)}/folders/${encodeURIComponent(folderId)}/messages`
      return requestJson<FolderMessagesResponse>(query ? `${url}?${query}` : url, { cache: 'no-store' })
    },
    messageDetail: (sessionId: string, messageId: string) =>
      requestJson<MessageDetailResponse>(`/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}`, {
        cache: 'no-store'
      }),
    messageReview: (sessionId: string, messageId: string) =>
      requestJson<ReviewUpdateResponse>(`/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/review`, {
        cache: 'no-store'
      }),
    updateReview: (sessionId: string, messageId: string, payload: { flagged?: boolean; tags?: string[] | string }) =>
      requestJson<ReviewUpdateResponse>(`/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/review`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      }),
    clearReview: (sessionId: string, messageId: string) =>
      requestJson<ReviewUpdateResponse>(`/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/review`, {
        method: 'DELETE'
      }),
    messageJsonUrl: (sessionId: string, messageId: string) =>
      `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/export.json`,
    messageEmlUrl: (sessionId: string, messageId: string) =>
      `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/export.eml`,
    attachmentUrl: (sessionId: string, messageId: string, attachmentIndex: number) =>
      `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/attachments/${attachmentIndex}`,
    folderExtractUrl: (sessionId: string, folderId: string, query: Record<string, string | number | boolean | undefined>) =>
      `/api/sessions/${encodeURIComponent(sessionId)}/folders/${encodeURIComponent(folderId)}/messages/extract?${buildQuery(query)}`,
    messageExtractUrl: (sessionId: string, messageId: string, fields: string) =>
      `/api/sessions/${encodeURIComponent(sessionId)}/messages/${encodeURIComponent(messageId)}/extract?${buildQuery({ fields })}`,
    flaggedBundleUrl: (params: Record<string, string | number | boolean | undefined>) =>
      `/api/exports/flagged.zip?${buildQuery(params)}`
  }
}
