import { existsSync } from 'node:fs';
import { dirname } from 'node:path';
import { z } from 'zod';
import type { IRepository } from '../database/repository.js';
import type { Folder, EmailSummary, Email, AttachmentInfo, PaginatedResult } from '../types/index.js';
import { paginate } from '../types/index.js';
import type { IAttachmentReader } from '../applescript/content-readers.js';
import { MAX_ATTACHMENT_DOWNLOAD_SIZE } from '../types/mail.js';
import { appleTimestampToIso } from '../utils/dates.js';
import { extractPlainText } from '../parsers/html-stripper.js';
import { NotFoundError, ValidationError, AttachmentTooLargeError, AttachmentSaveError } from '../utils/errors.js';

// ---------------------------------------------------------------------------
// Zod input schemas for mail MCP tools
// ---------------------------------------------------------------------------

export const ListAccountsInput = z.strictObject({});

export const ListFoldersInput = z.strictObject({});

export const ListEmailsInput = z.strictObject({
    folder_id: z.number().int().positive().describe('The folder ID to list emails from (e.g., from list_folders)'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of emails to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of emails to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
    unread_only: z.boolean().default(false).describe('If true, only return unread emails. Defaults to false if omitted.'),
    after: z.string().optional().describe('Only include emails received on or after this ISO 8601 date (e.g., "2025-01-01T00:00:00Z"). If omitted, no start date filter.'),
    before: z.string().optional().describe('Only include emails received on or before this ISO 8601 date (e.g., "2025-12-31T23:59:59Z"). If omitted, no end date filter.'),
});

export const SearchEmailsInput = z.strictObject({
    query: z.string().min(1).describe('Search query text matched against subject and sender address (e.g., "invoice")'),
    folder_id: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('Folder ID to limit search to (e.g., from list_folders). If omitted, searches all folders.'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(200)
        .default(25)
        .describe('Maximum number of emails to return, 1-200 (e.g., 50). Defaults to 25 if omitted. Search results are metadata-only (~400 bytes each), so larger pages are efficient.'),
    offset: z.number().int().min(0).default(0).describe('Number of emails to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
    after: z.string().optional().describe('Only include emails received on or after this ISO 8601 date (e.g., "2025-01-01T00:00:00Z"). If omitted, no start date filter.'),
    before: z.string().optional().describe('Only include emails received on or before this ISO 8601 date (e.g., "2025-12-31T23:59:59Z"). If omitted, no end date filter.'),
});

export const GetEmailInput = z.strictObject({
    email_id: z.number().int().positive().describe('The numeric email ID to retrieve (e.g., from list_emails or search_emails)'),
    include_body: z.boolean().default(true).describe('Include the email body content in the response. Defaults to true if omitted.'),
    strip_html: z.boolean().default(true).describe('Strip HTML tags from the body, returning plain text. Defaults to true if omitted.'),
});

export const GetUnreadCountInput = z.strictObject({
    folder_id: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('Folder ID to get unread count for (e.g., from list_folders). If omitted, returns total unread count across all folders.'),
});

export const ListAttachmentsInput = z.strictObject({
    email_id: z.number().int().positive().describe('The email ID to list attachments for (e.g., from list_emails or search_emails)'),
});

export const DownloadAttachmentInput = z.strictObject({
    email_id: z.number().int().positive().describe('The email ID containing the attachment (e.g., from list_emails)'),
    attachment_index: z.number().int().positive().describe('The 1-based index of the attachment (e.g., 1 for the first attachment from list_attachments)'),
    save_path: z.string().min(1).describe('Absolute file path where the attachment should be saved (e.g., "/tmp/report.pdf")'),
});

export const ListFoldersWithAccountInput = z.strictObject({
    account_id: z.union([
        z.number().int().positive(),
        z.array(z.number().int().positive()),
        z.literal('all'),
    ]).optional().describe('Account filter: a single account ID, an array of IDs, "all" for every account, or omit to use the default account (e.g., from list_accounts).'),
});

export const SendEmailInput = z.strictObject({
    to: z.array(z.string().min(1)).min(1).describe('Recipient email addresses (e.g., ["user@example.com"])'),
    subject: z.string().min(1).describe('Email subject line (e.g., "Meeting follow-up")'),
    body: z.string().describe('Email body content — plain text or HTML depending on body_type'),
    body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type: "plain" for text or "html" for rich formatting. Defaults to "plain" if omitted.'),
    cc: z.array(z.string().min(1)).optional().describe('CC recipient email addresses. If omitted, no CC recipients are added.'),
    bcc: z.array(z.string().min(1)).optional().describe('BCC recipient email addresses. If omitted, no BCC recipients are added.'),
    reply_to: z.string().optional().describe('Reply-to email address. If omitted, replies go to the sender address.'),
    attachments: z.array(z.strictObject({
        path: z.string().min(1).describe('Absolute file path to attachment (e.g., "/Users/me/report.pdf")'),
        name: z.string().optional().describe('Display name for attachment. If omitted, uses the filename.'),
    })).optional().describe('File attachments to include. If omitted, no files are attached.'),
    inline_images: z.array(z.strictObject({
        path: z.string().min(1).describe('Absolute file path to the image (e.g., "/Users/me/logo.png")'),
        content_id: z.string().min(1).describe('Content ID for referencing in HTML body via cid: (e.g., "logo1")'),
    })).optional().describe('Inline images for HTML body. If omitted, no inline images are included.'),
    account_id: z.number().int().positive().optional().describe('Account ID to send from (e.g., from list_accounts). If omitted, uses the default account.'),
});

export const CreateDraftInput = z.strictObject({
    to: z.array(z.string().min(1)).min(1).describe('Recipient email addresses (e.g., ["user@example.com"])'),
    subject: z.string().min(1).describe('Email subject line (e.g., "Meeting follow-up")'),
    body: z.string().describe('Email body content — plain text or HTML depending on body_type'),
    body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type: "plain" for text or "html" for rich formatting. Defaults to "plain" if omitted.'),
    cc: z.array(z.string().min(1)).optional().describe('CC recipient email addresses. If omitted, no CC recipients are added.'),
    bcc: z.array(z.string().min(1)).optional().describe('BCC recipient email addresses. If omitted, no BCC recipients are added.'),
    reply_to: z.string().optional().describe('Reply-to email address. If omitted, replies go to the sender address.'),
    attachments: z.array(z.strictObject({
        path: z.string().min(1).describe('Absolute file path to attachment (e.g., "/Users/me/report.pdf")'),
        name: z.string().optional().describe('Display name for attachment. If omitted, uses the filename.'),
    })).optional().describe('File attachments to include. If omitted, no files are attached.'),
    inline_images: z.array(z.strictObject({
        path: z.string().min(1).describe('Absolute file path to the image (e.g., "/Users/me/logo.png")'),
        content_id: z.string().min(1).describe('Content ID for referencing in HTML body via cid: (e.g., "logo1")'),
    })).optional().describe('Inline images for HTML body. If omitted, no inline images are included.'),
    account_id: z.number().int().positive().optional().describe('Account ID to use (e.g., from list_accounts). If omitted, uses the default account.'),
});

export const ReplyEmailInput = z.strictObject({
    email_id: z.number().int().positive().describe('The numeric ID of the email to reply to (e.g., from list_emails or search_emails)'),
    body: z.string().min(1).describe('Reply body content — plain text or HTML depending on body_type'),
    body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type: "plain" for text or "html" for rich formatting. Defaults to "plain" if omitted.'),
});

export const ReplyAllEmailInput = z.strictObject({
    email_id: z.number().int().positive().describe('The numeric ID of the email to reply-all to (e.g., from list_emails or search_emails)'),
    body: z.string().min(1).describe('Reply body content — will be prepended to the quoted original message as plain text'),
    body_type: z.enum(['plain', 'html']).default('plain').describe('Body content type: "plain" for text. HTML bodies will be converted to plain text for this draft. Defaults to "plain" if omitted.'),
});

/** Validated parameters for listing mail folders. */
export type ListFoldersParams = z.infer<typeof ListFoldersInput>;
/** Validated parameters for listing mail folders with account filtering. */
export type ListFoldersWithAccountParams = z.infer<typeof ListFoldersWithAccountInput>;
/** Validated parameters for listing emails in a folder. */
export type ListEmailsParams = z.infer<typeof ListEmailsInput>;
/** Validated parameters for searching emails. */
export type SearchEmailsParams = z.infer<typeof SearchEmailsInput>;
/** Validated parameters for retrieving a single email. */
export type GetEmailParams = z.infer<typeof GetEmailInput>;
/** Validated parameters for getting an unread email count. */
export type GetUnreadCountParams = z.infer<typeof GetUnreadCountInput>;
/** Validated parameters for listing attachments on an email. */
export type ListAttachmentsParams = z.infer<typeof ListAttachmentsInput>;
/** Validated parameters for downloading an email attachment. */
export type DownloadAttachmentParams = z.infer<typeof DownloadAttachmentInput>;
/** Validated parameters for sending an email. */
export type SendEmailParams = z.infer<typeof SendEmailInput>;
/** Validated parameters for creating an outgoing message draft. */
export type CreateDraftParams = z.infer<typeof CreateDraftInput>;
/** Validated parameters for replying to an email. */
export type ReplyEmailParams = z.infer<typeof ReplyEmailInput>;
/** Validated parameters for creating a reply-all draft. */
export type ReplyAllEmailParams = z.infer<typeof ReplyAllEmailInput>;

/** Reads the HTML body content of an email from its data file. */
export interface IContentReader {
    readEmailBody(dataFilePath: string | null): string | null;
}

// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------

/** Converts a raw folder repository row into a Folder domain object. */
function transformFolder(row: ReturnType<IRepository['listFolders']>[number]): Folder {
    return {
        id: row.id,
        name: row.name ?? 'Unnamed',
        parentId: row.parentId,
        specialType: row.specialType,
        folderType: row.folderType,
        accountId: row.accountId,
        messageCount: row.messageCount,
        unreadCount: row.unreadCount,
    };
}

/** Converts a raw email repository row into an EmailSummary domain object. */
function transformEmailSummary(row: ReturnType<IRepository['getEmail']> & {}): EmailSummary {
    return {
        id: row.id,
        folderId: row.folderId,
        subject: row.subject,
        sender: row.sender,
        senderAddress: row.senderAddress,
        preview: row.preview,
        isRead: row.isRead === 1,
        timeReceived: appleTimestampToIso(row.timeReceived),
        timeSent: appleTimestampToIso(row.timeSent),
        hasAttachment: row.hasAttachment === 1,
        priority: row.priority as EmailSummary['priority'],
        flagStatus: row.flagStatus as EmailSummary['flagStatus'],
        categories: parseCategories(row.categories),
    };
}

/** Parses a raw categories buffer into an array of category strings. Handles both null-delimited and comma-delimited formats. */
function parseCategories(buffer: Buffer | null): string[] {
    if (buffer == null || buffer.length === 0) {
        return [];
    }
    try {
        const text = buffer.toString('utf-8');
        const categories = text.includes('\0')
            ? text.split('\0').filter(s => s.length > 0)
            : text.split(',').map(s => s.trim()).filter(s => s.length > 0);
        return categories;
    }
    catch {
        return [];
    }
}

/** Converts a raw email repository row, optional body, and strip preference into a full Email domain object. */
function transformEmail(row: ReturnType<IRepository['getEmail']> & {}, body: string | null, stripHtml: boolean): Email {
    const summary = transformEmailSummary(row);
    let processedBody: string | null = null;
    let htmlBody: string | null = null;
    if (body != null) {
        htmlBody = body;
        processedBody = stripHtml ? extractPlainText(body) : body;
    }
    return {
        ...summary,
        recipients: row.recipients,
        displayTo: row.displayTo,
        toAddresses: row.toAddresses,
        ccAddresses: row.ccAddresses,
        size: row.size,
        messageId: row.messageId ?? null,
        conversationId: row.conversationId ?? null,
        body: processedBody,
        htmlBody: stripHtml ? null : htmlBody,
    };
}

/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullContentReader: IContentReader = {
    readEmailBody: () => null,
};

// ---------------------------------------------------------------------------
// MailTools -- provides read and download operations for Outlook mail
// ---------------------------------------------------------------------------

/** Exposes mail read, search, and attachment operations backed by a repository and optional readers. */
export class MailTools {
    private readonly repository: IRepository;
    private readonly contentReader: IContentReader;
    private readonly attachmentReader?: IAttachmentReader;

    constructor(repository: IRepository, contentReader: IContentReader = nullContentReader, attachmentReader?: IAttachmentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
        this.attachmentReader = attachmentReader;
    }

    /** Returns all mail folders across configured accounts. */
    listFolders(_params: ListFoldersParams): Folder[] {
        const rows = this.repository.listFolders();
        return rows.map(transformFolder);
    }

    /** Returns a paginated list of email summaries from a specific folder, optionally filtered to unread only and/or by date range. */
    listEmails(params: ListEmailsParams): PaginatedResult<EmailSummary> {
        const { folder_id, limit, offset, unread_only, after, before } = params;
        const rows = unread_only
            ? this.repository.listUnreadEmails(folder_id, limit + 1, offset, after, before)
            : this.repository.listEmails(folder_id, limit + 1, offset, after, before);
        return paginate(rows.map(transformEmailSummary), limit);
    }

    /** Searches emails by subject, sender, or preview text, optionally scoped to a single folder and/or by date range. */
    searchEmails(params: SearchEmailsParams): PaginatedResult<EmailSummary> {
        const { query, folder_id, limit, offset, after, before } = params;
        const rows = folder_id != null
            ? this.repository.searchEmailsInFolder(folder_id, query, limit + 1, offset, after, before)
            : this.repository.searchEmails(query, limit + 1, offset, after, before);
        return paginate(rows.map(transformEmailSummary), limit);
    }

    /** Retrieves a single email by ID with body content and attachment metadata, or null if not found. */
    getEmail(params: GetEmailParams): (Email & { attachments: AttachmentInfo[] }) | null {
        const { email_id, include_body, strip_html } = params;
        const row = this.repository.getEmail(email_id);
        if (row == null) {
            return null;
        }
        let body: string | null = null;
        if (include_body && row.dataFilePath != null) {
            body = this.contentReader.readEmailBody(row.dataFilePath);
        }
        let attachments: AttachmentInfo[] = [];
        if (this.attachmentReader != null && row.hasAttachment === 1) {
            attachments = this.attachmentReader.listAttachments(email_id);
        }
        return {
            ...transformEmail(row, body, strip_html),
            attachments,
        };
    }

    /** Returns the unread email count for a specific folder or across all folders. */
    getUnreadCount(params: GetUnreadCountParams): { count: number } {
        const { folder_id } = params;
        const count = folder_id != null
            ? this.repository.getUnreadCountByFolder(folder_id)
            : this.repository.getUnreadCount();
        return { count };
    }

    /**
     * Lists attachment metadata for a given email.
     * @param params - Contains the email ID to inspect.
     * @returns Array of attachment info objects.
     * @throws NotFoundError if the email does not exist.
     */
    listAttachments(params: ListAttachmentsParams): AttachmentInfo[] {
        if (this.attachmentReader == null) {
            return [];
        }
        const { email_id } = params;
        const row = this.repository.getEmail(email_id);
        if (row == null) {
            throw new NotFoundError('Email', email_id);
        }
        return this.attachmentReader.listAttachments(email_id);
    }

    /**
     * Saves an email attachment to disk at the specified path.
     * @param params - Contains the email ID, attachment index, and destination path.
     * @returns The saved attachment's name, path, and size.
     * @throws ValidationError if no attachment reader is available or the destination directory is missing.
     * @throws NotFoundError if the email or attachment does not exist.
     * @throws AttachmentTooLargeError if the attachment exceeds the size limit.
     * @throws AttachmentSaveError if the save operation fails.
     */
    downloadAttachment(params: DownloadAttachmentParams): { name: string; savedTo: string; size: number } {
        if (this.attachmentReader == null) {
            throw new ValidationError('Attachment reader not available');
        }
        const { email_id, attachment_index, save_path } = params;
        const row = this.repository.getEmail(email_id);
        if (row == null) {
            throw new NotFoundError('Email', email_id);
        }
        const dir = dirname(save_path);
        if (!existsSync(dir)) {
            throw new ValidationError(`Directory does not exist: ${dir}`);
        }
        const attachments = this.attachmentReader.listAttachments(email_id);
        const attachment = attachments.find(a => a.index === attachment_index);
        if (attachment == null) {
            throw new NotFoundError('Attachment', attachment_index);
        }
        if (attachment.size > MAX_ATTACHMENT_DOWNLOAD_SIZE) {
            throw new AttachmentTooLargeError(attachment.name, attachment.size, MAX_ATTACHMENT_DOWNLOAD_SIZE);
        }
        const result = this.attachmentReader.saveAttachment(email_id, attachment_index, save_path);
        if (!result.success) {
            throw new AttachmentSaveError(attachment.name, result.error ?? 'Unknown error');
        }
        return {
            name: result.name ?? attachment.name,
            savedTo: result.savedTo ?? save_path,
            size: result.fileSize ?? attachment.size,
        };
    }
}

/** Factory that creates a MailTools instance with the given repository, content reader, and optional attachment reader. */
export function createMailTools(repository: IRepository, contentReader: IContentReader = nullContentReader, attachmentReader?: IAttachmentReader): MailTools {
    return new MailTools(repository, contentReader, attachmentReader);
}
