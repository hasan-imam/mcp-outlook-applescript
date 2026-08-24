import { existsSync } from 'node:fs';
import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseSendEmailResult } from './parser.js';
import type { CreateDraftParams } from './scripts.js';
import { AppleScriptError, AttachmentNotFoundError, MailSendError } from '../utils/errors.js';

/** A file attachment identified by its filesystem path and optional display name. */
export interface Attachment {
    readonly path: string;
    readonly name?: string;
}

/** An inline image embedded in an HTML email body, referenced by content ID. */
export interface InlineImage {
    readonly path: string;
    readonly contentId: string;
}

/** Re-export reply parameters used by this sender. */
export type ReplyEmailParams = scripts.ReplyEmailParams;
/** Re-export reply-all parameters used by this sender. */
export type ReplyAllEmailParams = scripts.ReplyAllEmailParams;

/** Input parameters for composing and sending an email. */
export interface MailSenderParams {
    readonly to: readonly string[];
    readonly subject: string;
    readonly body: string;
    readonly bodyType: 'plain' | 'html';
    readonly cc?: readonly string[];
    readonly bcc?: readonly string[];
    readonly replyTo?: string;
    readonly attachments?: readonly Attachment[];
    readonly inlineImages?: readonly InlineImage[];
    readonly accountId?: number;
}

/** Metadata returned after an email is successfully sent. */
export interface SentEmail {
    readonly messageId: string;
    readonly sentAt: string;
}

/** Contract for sending email through Outlook. */
export interface IMailSender {
    /** Composes and sends an email with the given parameters. */
    sendEmail(params: MailSenderParams): SentEmail;
    /** Composes and sends a reply to an existing email. */
    replyEmail(params: scripts.ReplyEmailParams): SentEmail;
    /** Creates a reply-all draft for an existing email. */
    replyAllEmail(params: scripts.ReplyAllEmailParams): SentEmail;
    /** Creates a new outgoing message draft and saves it without sending. */
    createDraft(params: MailSenderParams): { draftId: string };
}

/** Sends email through Outlook via AppleScript. */
export class AppleScriptMailSender implements IMailSender {
    /**
     * Validates attachment paths, builds the send-email AppleScript, executes
     * it, and parses the result.
     * @param params - Email composition parameters (recipients, body, attachments, etc.).
     * @returns The sent message's ID and timestamp.
     * @throws AttachmentNotFoundError if any attachment or inline image path does not exist.
     * @throws MailSendError if Outlook reports a send failure.
     */
    sendEmail(params: MailSenderParams): SentEmail {
        if (params.attachments != null) {
            for (const attachment of params.attachments) {
                if (!existsSync(attachment.path)) {
                    throw new AttachmentNotFoundError(attachment.path);
                }
            }
        }
        if (params.inlineImages != null) {
            for (const image of params.inlineImages) {
                if (!existsSync(image.path)) {
                    throw new AttachmentNotFoundError(image.path);
                }
            }
        }
        let scriptParams: scripts.SendEmailParams = {
            to: params.to,
            subject: params.subject,
            body: params.body,
            bodyType: params.bodyType,
        };
        if (params.cc != null)
            scriptParams = { ...scriptParams, cc: params.cc };
        if (params.bcc != null)
            scriptParams = { ...scriptParams, bcc: params.bcc };
        if (params.replyTo != null)
            scriptParams = { ...scriptParams, replyTo: params.replyTo };
        if (params.attachments != null)
            scriptParams = { ...scriptParams, attachments: params.attachments };
        if (params.inlineImages != null)
            scriptParams = { ...scriptParams, inlineImages: params.inlineImages };
        if (params.accountId != null)
            scriptParams = { ...scriptParams, accountId: params.accountId };
        const script = scripts.sendEmail(scriptParams);
        const output = executeAppleScriptOrThrow(script);
        const result = parseSendEmailResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse send email response');
        }
        if (!result.success) {
            throw new MailSendError(result.error ?? 'Unknown error');
        }
        return {
            messageId: result.messageId ?? '',
            sentAt: result.sentAt ?? '',
        };
    }

    /**
     * Sends a reply to an existing email via AppleScript.
     * @param params - Contains the source message ID and reply body/type.
     * @returns The sent reply's message ID and timestamp.
     * @throws MailSendError if Outlook reports a send failure.
     */
    replyEmail(params: scripts.ReplyEmailParams): SentEmail {
        const script = scripts.replyEmail(params);
        const output = executeAppleScriptOrThrow(script);
        const result = parseSendEmailResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse reply email response');
        }
        if (!result.success) {
            throw new MailSendError(result.error ?? 'Unknown error');
        }
        return {
            messageId: result.messageId ?? '',
            sentAt: result.sentAt ?? '',
        };
    }

    /**
     * Creates a reply-all draft for an existing email via AppleScript.
     * @param params - Contains the source message ID and reply body/type.
     * @returns The saved draft message ID.
     * @throws MailSendError if Outlook reports a failure or the draft cannot be saved.
     */
    replyAllEmail(params: scripts.ReplyAllEmailParams): SentEmail {
        const script = scripts.replyAllEmail(params);
        const output = executeAppleScriptOrThrow(script);
        const result = parseSendEmailResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse reply-all draft response');
        }
        if (!result.success) {
            throw new MailSendError(result.error ?? 'Unknown error');
        }
        return {
            messageId: result.messageId ?? '',
            sentAt: result.sentAt ?? '',
        };
    }

    /**
     * Creates a new outgoing message draft and saves it to the Drafts folder.
     * @param params - Email composition parameters (same as sendEmail).
     * @returns The saved draft's message ID.
     * @throws AttachmentNotFoundError if any attachment or inline image path does not exist.
     * @throws MailSendError if Outlook reports a failure saving the draft.
     */
    createDraft(params: MailSenderParams): { draftId: string } {
        if (params.attachments != null) {
            for (const attachment of params.attachments) {
                if (!existsSync(attachment.path)) {
                    throw new AttachmentNotFoundError(attachment.path);
                }
            }
        }
        if (params.inlineImages != null) {
            for (const image of params.inlineImages) {
                if (!existsSync(image.path)) {
                    throw new AttachmentNotFoundError(image.path);
                }
            }
        }
        let scriptParams: scripts.CreateDraftParams = {
            to: params.to,
            subject: params.subject,
            body: params.body,
            bodyType: params.bodyType,
        };
        if (params.cc != null)
            scriptParams = { ...scriptParams, cc: params.cc };
        if (params.bcc != null)
            scriptParams = { ...scriptParams, bcc: params.bcc };
        if (params.replyTo != null)
            scriptParams = { ...scriptParams, replyTo: params.replyTo };
        if (params.attachments != null)
            scriptParams = { ...scriptParams, attachments: params.attachments };
        if (params.inlineImages != null)
            scriptParams = { ...scriptParams, inlineImages: params.inlineImages };
        if (params.accountId != null)
            scriptParams = { ...scriptParams, accountId: params.accountId };
        const script = scripts.createDraft(scriptParams);
        const output = executeAppleScriptOrThrow(script);
        const result = parseSendEmailResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse create draft response');
        }
        if (!result.success) {
            throw new MailSendError(result.error ?? 'Unknown error');
        }
        return {
            draftId: result.messageId ?? '',
        };
    }
}

/** Creates a new AppleScriptMailSender instance. */
export function createMailSender(): IMailSender {
    return new AppleScriptMailSender();
}
