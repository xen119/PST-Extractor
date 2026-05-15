import type { MessageSummary } from './viewer'

export interface ReviewState {
  flagged: boolean
  tags: string[]
  createdAt: string
  updatedAt: string
}

export interface ReviewRecord extends ReviewState {
  mailboxKey: string
  fileName: string
  messageId: string
  descriptorId: string
  folderId: string
  folderPath: string
  messageClass: string
  kind: MessageSummary['kind']
  isMailLike: boolean
  subject: string
  senderName: string
  senderEmailAddress: string
  displayTo: string
  displayCC: string
  displayBCC: string
  resolvedDisplayTo: string
  resolvedDisplayCC: string
  resolvedDisplayBCC: string
}

export interface ReviewContext {
  mailboxKey: string
  fileName: string
  messageId: string
  descriptorId: string
  folderId: string
  folderPath: string
  messageClass: string
  kind: MessageSummary['kind']
  isMailLike: boolean
  subject: string
  senderName: string
  senderEmailAddress: string
  displayTo: string
  displayCC: string
  displayBCC: string
  resolvedDisplayTo: string
  resolvedDisplayCC: string
  resolvedDisplayBCC: string
}

export interface ReviewPatchInput extends ReviewContext {
  flagged?: boolean
  tags?: unknown
}

export interface ReviewSearchOptions {
  query?: string
  flaggedOnly?: boolean
  taggedOnly?: boolean
  tag?: string
  messageIds?: string[]
}
