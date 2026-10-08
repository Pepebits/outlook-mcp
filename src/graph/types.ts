export interface EmailAddress {
  name?: string;
  address?: string;
}
export interface Recipient {
  emailAddress: EmailAddress;
}
export interface Message {
  id: string;
  subject?: string | null;
  bodyPreview?: string;
  body?: { contentType: 'text' | 'html'; content: string };
  from?: Recipient;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  isDraft?: boolean;
  importance?: string;
  hasAttachments?: boolean;
  internetMessageId?: string;
  webLink?: string;
  flag?: { flagStatus: 'flagged' | 'complete' | 'notFlagged' };
  parentFolderId?: string;
}
export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
}
export interface Attachment {
  id: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  '@odata.type'?: string;
}
export interface Page<T> {
  value: T[];
  '@odata.nextLink'?: string;
}
