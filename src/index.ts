#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
    createAppleScriptRepository,
    createAppleScriptContentReaders,
    createAccountRepository,
    createCalendarWriter,
    createCalendarManager,
    createMailSender,
    isOutlookRunning,
    type IAccountRepository,
    type AppleScriptContentReaders,
    type ICalendarWriter,
    type ICalendarManager,
    type IMailSender,
    type CalendarWriterParams,
    type RecurrenceConfig,
    type EventUpdates,
    type MailSenderParams,
} from './applescript/index.js';
import { createMailTools, type MailTools } from './tools/mail.js';
import { createCalendarTools, type CalendarTools } from './tools/calendar.js';
import { createContactsTools, type ContactsTools } from './tools/contacts.js';
import { createTasksTools, type TasksTools } from './tools/tasks.js';
import { createNotesTools, type NotesTools } from './tools/notes.js';
import {
    createMailboxOrganizationTools,
    type MailboxOrganizationTools,
    ListAccountsInput,
    ListEmailsInput,
    SearchEmailsInput,
    GetEmailInput,
    GetUnreadCountInput,
    ListAttachmentsInput,
    DownloadAttachmentInput,
    ListCalendarsInput,
    ListEventsInput,
    GetEventInput,
    SearchEventsInput,
    CreateEventInput,
    RespondToEventInput,
    DeleteEventInput,
    UpdateEventInput,
    ListContactsInput,
    SearchContactsInput,
    GetContactInput,
    ListTasksInput,
    SearchTasksInput,
    GetTaskInput,
    ListNotesInput,
    GetNoteInput,
    SearchNotesInput,
    ListFoldersWithAccountInput,
    SendEmailInput,
    CreateDraftInput,
    ReplyEmailInput,
    ReplyAllEmailInput,
    PrepareDeleteEmailInput,
    ConfirmDeleteEmailInput,
    PrepareMoveEmailInput,
    ConfirmMoveEmailInput,
    PrepareArchiveEmailInput,
    ConfirmArchiveEmailInput,
    PrepareJunkEmailInput,
    ConfirmJunkEmailInput,
    PrepareDeleteFolderInput,
    ConfirmDeleteFolderInput,
    PrepareEmptyFolderInput,
    ConfirmEmptyFolderInput,
    PrepareBatchDeleteEmailsInput,
    PrepareBatchMoveEmailsInput,
    ConfirmBatchOperationInput,
    MarkEmailReadInput,
    MarkEmailUnreadInput,
    SetEmailFlagInput,
    ClearEmailFlagInput,
    SetEmailCategoriesInput,
    CreateFolderInput,
    RenameFolderInput,
    MoveFolderInput,
    type CreateEventParams,
} from './tools/index.js';
import { ApprovalTokenManager } from './approval/index.js';
import {
    wrapError,
    OutlookNotRunningError,
} from './utils/errors.js';
import { extractPlainText } from './parsers/html-stripper.js';

// =============================================================================
// Helper: JSON text result
// =============================================================================

function jsonResult(data: unknown) {
    return { content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }] };
}

function errorResult(message: string) {
    return { content: [{ type: 'text' as const, text: message }], isError: true as const };
}

// =============================================================================
// Server Factory
// =============================================================================

export function createServer(): McpServer {
    const server = new McpServer({
        name: 'outlook-mcp',
        version: '1.1.3',
    });

    const tokenManager = new ApprovalTokenManager();

    let initialized = false;
    let accountRepository: IAccountRepository | null = null;
    let mailTools: MailTools | null = null;
    let calendarTools: CalendarTools | null = null;
    let contactsTools: ContactsTools | null = null;
    let tasksTools: TasksTools | null = null;
    let notesTools: NotesTools | null = null;
    let orgTools: MailboxOrganizationTools | null = null;
    let calendarWriter: ICalendarWriter | null = null;
    let calendarManager: ICalendarManager | null = null;
    let mailSender: IMailSender | null = null;

    function initializeAppleScriptBackend(): void {
        if (!isOutlookRunning()) {
            throw new OutlookNotRunningError();
        }
        const repository = createAppleScriptRepository();
        const contentReaders: AppleScriptContentReaders = createAppleScriptContentReaders();
        accountRepository = createAccountRepository();
        mailTools = createMailTools(repository, contentReaders.email, contentReaders.attachment);
        calendarTools = createCalendarTools(repository, contentReaders.event);
        contactsTools = createContactsTools(repository, contentReaders.contact);
        tasksTools = createTasksTools(repository, contentReaders.task);
        notesTools = createNotesTools(repository, contentReaders.note);
        orgTools = createMailboxOrganizationTools(repository, tokenManager);
        calendarWriter = createCalendarWriter();
        calendarManager = createCalendarManager();
        mailSender = createMailSender();
        initialized = true;
    }

    function ensureInitialized(): void {
        if (initialized)
            return;
        initializeAppleScriptBackend();
    }

    /**
     * Wraps a tool handler with lazy initialization and structured error handling.
     * The generic `A` carries through the Zod-validated arg type from server.tool().
     */
    function handle<A, T>(fn: (args: A) => T | Promise<T>) {
        return async (args: A) => {
            try {
                ensureInitialized();
                return await fn(args);
            }
            catch (error) {
                const wrapped = wrapError(error, 'An error occurred');
                return errorResult(`${wrapped.code}: ${wrapped.message}`);
            }
        };
    }

    // =========================================================================
    // Account & Folder Tools
    // =========================================================================

    server.tool('list_accounts', 'List all mail accounts configured in Microsoft Outlook for Mac. Use this to discover account IDs needed by other tools (e.g., list_folders with account_id filter). Returns an array of accounts, each with id, name, email, and type fields. Note: only Exchange, IMAP, and POP account types are returned; newer Microsoft 365 account types may not appear. Requires Outlook to be running or returns an error.',
        ListAccountsInput.shape,
        handle(() => {
            const accounts = accountRepository!.listAccounts();
            return jsonResult({
                accounts: accounts.map(acc => ({
                    id: acc.id, name: acc.name, email: acc.email, type: acc.type,
                })),
            });
        }),
    );

    server.tool('list_folders', 'List all mail folders in Outlook with their message and unread counts. Use this to discover folder IDs needed by list_emails, search_emails, prepare_move_email, and other folder-based tools. Returns an array of folders, each with id, name, messageCount, and unreadCount. When account_id is provided, returns folders grouped by account. Returns an error if Outlook is not running.',
        ListFoldersWithAccountInput.shape,
        handle((args) => {
            const accountIds = resolveAccountIds(args.account_id, accountRepository!);
            if (accountIds.length > 1 || args.account_id === 'all') {
                const foldersWithAccount = accountRepository!.listMailFoldersByAccounts(accountIds);
                const accounts = accountRepository!.listAccounts();
                const groupedByAccount = accountIds.map(accountId => {
                    const account = accounts.find(a => a.id === accountId);
                    const folders = foldersWithAccount
                        .filter(f => f.accountId === accountId)
                        .map(f => ({
                            id: f.id, name: f.name, unreadCount: f.unreadCount, messageCount: f.messageCount,
                        }));
                    return {
                        account_id: accountId,
                        account_name: account?.name ?? null,
                        account_email: account?.email ?? null,
                        folders,
                    };
                });
                return jsonResult({ accounts: groupedByAccount });
            }
            return jsonResult(mailTools!.listFolders({}));
        }),
    );

    // =========================================================================
    // Mail Tools
    // =========================================================================

    server.tool('list_emails', 'List email summaries in a specific mail folder, ordered newest first. Each result includes a 500-char body preview for browsing without needing get_email. Returns {items, count, hasMore} — when hasMore is true, increment offset by limit to fetch the next page. Use after/before (ISO 8601) to filter by received date. For full email content, use get_email. Requires Outlook to be running.',
        ListEmailsInput.shape,
        handle((args) => jsonResult(mailTools!.listEmails(args))),
    );

    server.tool('search_emails', 'Search emails by matching the query against subject line and sender address. Returns metadata only (subject, sender, date, flags) — no body preview. For email content, call get_email on a matching ID. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. Use after/before (ISO 8601) to filter by received date. Use folder_id to scope to one folder. Pagination is fast at any offset.',
        SearchEmailsInput.shape,
        handle((args) => jsonResult(mailTools!.searchEmails(args))),
    );

    server.tool('get_email', 'Get full details of a single email including its body content. Use this after list_emails or search_emails to read the complete message. Returns all summary fields plus recipients, body (plain text when strip_html is true, the default), attachments list, messageId, and conversationId. Set strip_html to false to also receive the raw HTML in the htmlBody field. Returns an error if the email ID does not exist. Set include_body to false for metadata only.',
        GetEmailInput.shape,
        handle((args) => {
            const result = mailTools!.getEmail(args);
            if (result == null) return errorResult('Email not found');
            return jsonResult(result);
        }),
    );

    server.tool('get_unread_count', 'Get the count of unread emails, either across all folders or for a specific folder. Use this for a quick check without fetching full email lists. Returns an object with a single "count" field (number). Optionally pass folder_id to limit the count to a specific folder (get folder IDs from list_folders). Requires Outlook to be running.',
        GetUnreadCountInput.shape,
        handle((args) => jsonResult(mailTools!.getUnreadCount(args))),
    );

    server.tool('list_attachments', 'List attachment metadata for a specific email. Use this before download_attachment to discover attachment indices and sizes. Returns an array of attachment objects with index, name, size (bytes), and contentType. Returns an error if the email ID does not exist. Returns an empty array if the email has no attachments.',
        ListAttachmentsInput.shape,
        handle((args) => jsonResult(mailTools!.listAttachments(args))),
    );

    server.tool('download_attachment', 'Download and save an email attachment to a file on disk. Use this after list_attachments to get the 1-based attachment index. Returns the saved file path, name, and size in bytes. Writes the file to the specified save_path — the parent directory must already exist. Returns an error if the email or attachment does not exist, the directory is invalid, or the attachment exceeds the size limit.',
        DownloadAttachmentInput.shape,
        handle((args) => jsonResult(mailTools!.downloadAttachment(args))),
    );

    // =========================================================================
    // Calendar Tools
    // =========================================================================

    server.tool('list_calendars', 'List all calendar folders configured in Outlook. Use this to discover calendar IDs for filtering in list_events or specifying a target calendar in create_event. Returns an array of calendar objects with id, name, and accountId. Requires Outlook to be running or returns an error.',
        ListCalendarsInput.shape,
        handle(() => jsonResult(calendarTools!.listCalendars({}))),
    );

    server.tool('list_events', 'List calendar events with optional date range or calendar filtering. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. Filter by date range (start_date + end_date in ISO 8601) and/or calendar_id (from list_calendars). Use get_event for full details including attendees and description.',
        ListEventsInput.shape,
        handle((args) => jsonResult(calendarTools!.listEvents(args))),
    );

    server.tool('get_event', 'Get full details of a single calendar event. Use this after list_events or search_events to see the complete event including location, description, organizer, and attendees list. Returns all summary fields plus location, description, organizer, attendees (with name and status), and recurrence IDs. Returns an error if the event ID does not exist.',
        GetEventInput.shape,
        handle((args) => {
            const result = calendarTools!.getEvent(args);
            if (result == null) return errorResult('Event not found');
            return jsonResult(result);
        }),
    );

    server.tool('search_events', 'Search calendar events by matching the query against event titles. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. Use after/before (ISO 8601) to filter by event start date. For full event details including attendees, call get_event on a matching ID.',
        SearchEventsInput.shape,
        handle((args) => jsonResult(calendarTools!.searchEvents(args))),
    );

    server.tool('create_event', 'Create a new calendar event in Outlook. Use this to schedule meetings or reminders. Returns the created event with id, title, start/end dates, calendar_id, and is_recurring flag. Optionally specify a target calendar_id (from list_calendars), location, description, all-day flag, or recurrence pattern. Returns an error if start_date is not before end_date or if the calendar is unavailable. Use delete_event to remove a created event.',
        CreateEventInput.shape,
        handle((args) => {
            if (calendarWriter == null) return errorResult('Event creation is not available');
            const params = CreateEventInput.parse(args);
            const writerParams = buildCalendarWriterParams(params);
            const created = calendarWriter.createEvent(writerParams);
            return jsonResult({
                id: created.id,
                title: params.title,
                start_date: params.start_date,
                end_date: params.end_date,
                calendar_id: created.calendarId,
                location: params.location ?? null,
                description: params.description ?? null,
                is_all_day: params.is_all_day,
                is_recurring: params.recurrence != null,
            });
        }),
    );

    server.tool('respond_to_event', 'Respond to a meeting invitation with accept, decline, or tentative. Updates your RSVP status in Outlook and by default sends a response notification to the organizer (set send_response to false to suppress). Returns a confirmation message. Use get_event first to review the event details before responding. Returns an error if the event ID does not exist or event response is unavailable.',
        RespondToEventInput.shape,
        handle((args) => {
            if (calendarManager == null) return errorResult('Event response is not available');
            const params = RespondToEventInput.parse(args);
            const result = calendarManager.respondToEvent(params.event_id, params.response, params.send_response, params.comment);
            const responseText = params.response === 'accept' ? 'accepted'
                : params.response === 'decline' ? 'declined'
                    : 'tentatively accepted';
            return { content: [{ type: 'text' as const, text: `Successfully ${responseText} event ${result.eventId}` }] };
        }),
    );

    server.tool('delete_event', 'Delete a calendar event from Outlook. For recurring events, use apply_to to choose between deleting a single instance or the entire series. Returns a confirmation message. This action is permanent and cannot be undone. Returns an error if the event ID does not exist. Use get_event to verify the event before deleting.',
        DeleteEventInput.shape,
        handle((args) => {
            if (calendarManager == null) return errorResult('Event deletion is not available');
            const params = DeleteEventInput.parse(args);
            calendarManager.deleteEvent(params.event_id, params.apply_to);
            const deleteText = params.apply_to === 'all_in_series' ? ' (entire series)' : '';
            return { content: [{ type: 'text' as const, text: `Successfully deleted event ${params.event_id}${deleteText}` }] };
        }),
    );

    server.tool('update_event', 'Update a calendar event in Outlook. Only the fields you specify will be changed — omitted fields remain unchanged. For recurring events, use apply_to to choose between updating a single instance or the entire series. Returns a confirmation with the event ID and list of updated field names. Returns an error if the event ID does not exist or start_date is not before end_date. Use get_event to review current values before updating.',
        UpdateEventInput.shape,
        handle((args) => {
            if (calendarManager == null) return errorResult('Event update is not available');
            const params = UpdateEventInput.parse(args);
            const updates: EventUpdates = {
                ...(params.title != null && { title: params.title }),
                ...(params.start_date != null && { startDate: params.start_date }),
                ...(params.end_date != null && { endDate: params.end_date }),
                ...(params.location != null && { location: params.location }),
                ...(params.description != null && { description: params.description }),
                ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
            };
            const result = calendarManager.updateEvent(params.event_id, updates, params.apply_to);
            const updateText = params.apply_to === 'all_in_series' ? ' (entire series)' : '';
            return { content: [{ type: 'text' as const, text: `Successfully updated event ${result.id}${updateText}. Updated fields: ${result.updatedFields.join(', ')}` }] };
        }),
    );

    // =========================================================================
    // Contact Tools
    // =========================================================================

    server.tool('list_contacts', 'List contact summaries from Outlook with pagination. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. For full details including emails, phones, and addresses, use get_contact. Use search_contacts to find contacts by name.',
        ListContactsInput.shape,
        handle((args) => jsonResult(contactsTools!.listContacts(args))),
    );

    server.tool('search_contacts', 'Search contacts by matching the query against contact names. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. For full contact details, call get_contact on a matching ID.',
        SearchContactsInput.shape,
        handle((args) => jsonResult(contactsTools!.searchContacts(args))),
    );

    server.tool('get_contact', 'Get full details of a single contact. Use this after list_contacts or search_contacts to see complete information. Returns all summary fields plus firstName, lastName, company, jobTitle, department, emails (with type and address), phones (with type and number), addresses, and notes. Returns an error if the contact ID does not exist.',
        GetContactInput.shape,
        handle((args) => {
            const result = contactsTools!.getContact(args);
            if (result == null) return errorResult('Contact not found');
            return jsonResult(result);
        }),
    );

    // =========================================================================
    // Task Tools
    // =========================================================================

    server.tool('list_tasks', 'List task summaries from Outlook with pagination and optional completion filtering. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. Set include_completed to false to see only incomplete tasks. For full details including body and reminder, use get_task.',
        ListTasksInput.shape,
        handle((args) => jsonResult(tasksTools!.listTasks(args))),
    );

    server.tool('search_tasks', 'Search tasks by matching the query against task names. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. For full task details, call get_task on a matching ID.',
        SearchTasksInput.shape,
        handle((args) => jsonResult(tasksTools!.searchTasks(args))),
    );

    server.tool('get_task', 'Get full details of a single task. Use this after list_tasks or search_tasks to see the complete task. Returns all summary fields plus startDate, completedDate, body, hasReminder, reminderDate, and categories. Returns an error if the task ID does not exist.',
        GetTaskInput.shape,
        handle((args) => {
            const result = tasksTools!.getTask(args);
            if (result == null) return errorResult('Task not found');
            return jsonResult(result);
        }),
    );

    // =========================================================================
    // Note Tools
    // =========================================================================

    server.tool('list_notes', 'List note summaries from Outlook with pagination. Returns {items, count, hasMore} — increment offset by limit when hasMore is true. For full details including body content, use get_note. Use search_notes to find notes by title.',
        ListNotesInput.shape,
        handle((args) => jsonResult(notesTools!.listNotes(args))),
    );

    server.tool('get_note', 'Get full details of a single note. Use this after list_notes or search_notes to read the complete note content. Returns all summary fields plus body, createdDate, and categories. Returns an error if the note ID does not exist.',
        GetNoteInput.shape,
        handle((args) => {
            const result = notesTools!.getNote(args);
            if (result == null) return errorResult('Note not found');
            return jsonResult(result);
        }),
    );

    server.tool('search_notes', 'Search notes by matching the query against note titles/names only (does not search body content). Returns {items, count, hasMore} — increment offset by limit when hasMore is true. For full note content, call get_note on a matching ID.',
        SearchNotesInput.shape,
        handle((args) => jsonResult(notesTools!.searchNotes(args))),
    );

    // =========================================================================
    // Send Email
    // =========================================================================

    server.tool('send_email', 'Send an email from Outlook with optional CC, BCC, file attachments, inline images, and HTML formatting. This action sends the email immediately and cannot be undone. Returns the sent message_id and sent_at timestamp. Use list_accounts to find account_id if sending from a non-default account. Returns an error if required fields (to, subject) are missing or if attachment file paths do not exist.',
        SendEmailInput.shape,
        handle((args) => {
            if (mailSender == null) return errorResult('Email sending is not available');
            const params = SendEmailInput.parse(args);
            const sendParams: MailSenderParams = {
                to: params.to,
                subject: params.subject,
                body: params.body,
                bodyType: params.body_type,
                ...(params.cc != null && { cc: params.cc }),
                ...(params.bcc != null && { bcc: params.bcc }),
                ...(params.reply_to != null && { replyTo: params.reply_to }),
                ...(params.attachments != null && { attachments: params.attachments }),
                ...(params.inline_images != null && {
                    inlineImages: params.inline_images.map(img => ({
                        path: img.path,
                        contentId: img.content_id,
                    })),
                }),
                ...(params.account_id != null && { accountId: params.account_id }),
            };
            const sent = mailSender.sendEmail(sendParams);
            return jsonResult({
                message_id: sent.messageId,
                sent_at: sent.sentAt,
                status: 'sent',
            });
        }),
    );

    server.tool('create_draft', 'Create a new email draft in Outlook without sending it. Supports To, CC, BCC, file attachments, inline images, and HTML formatting. The draft is saved to the Drafts folder and can be reviewed and sent from Outlook. Returns the draft message_id. Use list_accounts to find account_id if using a non-default account. Returns an error if required fields (to, subject) are missing or if attachment file paths do not exist.',
        CreateDraftInput.shape,
        handle((args) => {
            if (mailSender == null) return errorResult('Email sending is not available');
            const params = CreateDraftInput.parse(args);
            const draftParams: MailSenderParams = {
                to: params.to,
                subject: params.subject,
                body: params.body,
                bodyType: params.body_type,
                ...(params.cc != null && { cc: params.cc }),
                ...(params.bcc != null && { bcc: params.bcc }),
                ...(params.reply_to != null && { replyTo: params.reply_to }),
                ...(params.attachments != null && { attachments: params.attachments }),
                ...(params.inline_images != null && {
                    inlineImages: params.inline_images.map(img => ({
                        path: img.path,
                        contentId: img.content_id,
                    })),
                }),
                ...(params.account_id != null && { accountId: params.account_id }),
            };
            const saved = mailSender.createDraft(draftParams);
            return jsonResult({
                draft_id: saved.draftId,
                status: 'saved_as_draft',
            });
        }),
    );

    server.tool('reply_to_email', 'Reply to an existing email in Outlook. Creates a threaded reply to the specified message, sets the reply body, and sends it. Returns the sent message_id and sent_at timestamp. Use get_email to read the original message before replying if needed.',
        ReplyEmailInput.shape,
        handle((args) => {
            if (mailSender == null) return errorResult('Email sending is not available');
            const params = ReplyEmailInput.parse(args);
            const sent = mailSender.replyEmail({
                messageId: params.email_id,
                body: params.body,
                bodyType: params.body_type,
            });
            return jsonResult({
                message_id: sent.messageId,
                sent_at: sent.sentAt,
                status: 'sent',
            });
        }),
    );

    server.tool('reply_all_email', 'Create a reply-all draft from an existing email in Outlook. Uses the original To/CC recipients so the draft is a true reply-all, and saves it to the Drafts folder without sending. The reply body is prepended to the quoted original message. HTML bodies are converted to plain text. Returns the draft message_id.',
        ReplyAllEmailInput.shape,
        handle((args) => {
            if (mailSender == null) return errorResult('Email sending is not available');
            const params = ReplyAllEmailInput.parse(args);
            const replyBody = params.body_type === 'html' ? extractPlainText(params.body) : params.body;
            const saved = mailSender.replyAllEmail({
                messageId: params.email_id,
                body: replyBody,
                bodyType: 'plain',
            });
            return jsonResult({
                draft_id: saved.messageId,
                status: 'saved_as_draft',
            });
        }),
    );

    // =========================================================================
    // Mailbox Organization Tools (Destructive — Two-Phase)
    // =========================================================================

    server.tool('prepare_delete_email', 'Prepare to delete an email (move to trash). Returns a preview of the email and an approval token with expiration time. Call confirm_delete_email with the token to execute the deletion. Returns an error if the email ID does not exist. The token expires after 5 minutes.',
        PrepareDeleteEmailInput.shape,
        handle((args) => jsonResult(orgTools!.prepareDeleteEmail(args))),
    );

    server.tool('confirm_delete_email', 'Confirm and execute deletion of an email using a token from prepare_delete_email. Moves the email to the Deleted Items folder. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the email has changed since the prepare step.',
        ConfirmDeleteEmailInput.shape,
        handle((args) => jsonResult(orgTools!.confirmDeleteEmail(args))),
    );

    server.tool('prepare_move_email', 'Prepare to move an email to another folder. Returns a preview of the email, destination folder details, and an approval token. Call confirm_move_email with the token to execute the move. Returns an error if the email or destination folder ID does not exist. The token expires after 5 minutes.',
        PrepareMoveEmailInput.shape,
        handle((args) => jsonResult(orgTools!.prepareMoveEmail(args))),
    );

    server.tool('confirm_move_email', 'Confirm and execute moving an email using a token from prepare_move_email. Moves the email to the destination folder specified during the prepare step. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the email has changed since the prepare step.',
        ConfirmMoveEmailInput.shape,
        handle((args) => jsonResult(orgTools!.confirmMoveEmail(args))),
    );

    server.tool('prepare_archive_email', 'Prepare to archive an email (move to Archive folder). Returns a preview of the email and an approval token. Call confirm_archive_email with the token to execute. Returns an error if the email ID does not exist. The token expires after 5 minutes.',
        PrepareArchiveEmailInput.shape,
        handle((args) => jsonResult(orgTools!.prepareArchiveEmail(args))),
    );

    server.tool('confirm_archive_email', 'Confirm and execute archiving an email using a token from prepare_archive_email. Moves the email to the Archive folder. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the email has changed since the prepare step.',
        ConfirmArchiveEmailInput.shape,
        handle((args) => jsonResult(orgTools!.confirmArchiveEmail(args))),
    );

    server.tool('prepare_junk_email', 'Prepare to mark an email as junk (move to Junk folder). Returns a preview of the email and an approval token. Call confirm_junk_email with the token to execute. Returns an error if the email ID does not exist. The token expires after 5 minutes.',
        PrepareJunkEmailInput.shape,
        handle((args) => jsonResult(orgTools!.prepareJunkEmail(args))),
    );

    server.tool('confirm_junk_email', 'Confirm and execute marking an email as junk using a token from prepare_junk_email. Moves the email to the Junk folder. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the email has changed since the prepare step.',
        ConfirmJunkEmailInput.shape,
        handle((args) => jsonResult(orgTools!.confirmJunkEmail(args))),
    );

    server.tool('prepare_delete_folder', 'Prepare to delete a mail folder and all its messages. Returns a preview of the folder (with message count) and an approval token. Call confirm_delete_folder with the token to execute. Returns an error if the folder ID does not exist. The token expires after 5 minutes.',
        PrepareDeleteFolderInput.shape,
        handle((args) => jsonResult(orgTools!.prepareDeleteFolder(args))),
    );

    server.tool('confirm_delete_folder', 'Confirm and execute deletion of a mail folder using a token from prepare_delete_folder. Permanently deletes the folder and all its messages. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the folder has changed since the prepare step.',
        ConfirmDeleteFolderInput.shape,
        handle((args) => jsonResult(orgTools!.confirmDeleteFolder(args))),
    );

    server.tool('prepare_empty_folder', 'Prepare to empty a mail folder by deleting all its messages. Returns a preview of the folder (with message count) and an approval token. Call confirm_empty_folder with the token to execute. Returns an error if the folder ID does not exist. The token expires after 5 minutes.',
        PrepareEmptyFolderInput.shape,
        handle((args) => jsonResult(orgTools!.prepareEmptyFolder(args))),
    );

    server.tool('confirm_empty_folder', 'Confirm and execute emptying a mail folder using a token from prepare_empty_folder. Deletes all messages in the folder. Returns a success message. Returns an error if the token is invalid, expired, already used, or if the folder has changed since the prepare step.',
        ConfirmEmptyFolderInput.shape,
        handle((args) => jsonResult(orgTools!.confirmEmptyFolder(args))),
    );

    server.tool('prepare_batch_delete_emails', 'Prepare to delete multiple emails (up to 50). Returns individual approval tokens per email so you can selectively confirm. Call confirm_batch_operation with the desired tokens to execute. Returns an error if any email ID does not exist. Tokens expire after 5 minutes.',
        PrepareBatchDeleteEmailsInput.shape,
        handle((args) => jsonResult(orgTools!.prepareBatchDeleteEmails(args))),
    );

    server.tool('prepare_batch_move_emails', 'Prepare to move multiple emails (up to 50) to a destination folder. Returns individual approval tokens per email so you can selectively confirm. Call confirm_batch_operation with the desired tokens to execute. Returns an error if any email or the destination folder ID does not exist. Tokens expire after 5 minutes.',
        PrepareBatchMoveEmailsInput.shape,
        handle((args) => jsonResult(orgTools!.prepareBatchMoveEmails(args))),
    );

    server.tool('confirm_batch_operation', 'Confirm and execute a batch operation using tokens from prepare_batch_delete_emails or prepare_batch_move_emails. You may selectively confirm by including only the desired token/email pairs. Returns per-email results with success/failure status and a summary with total, succeeded, and failed counts. Individual tokens that are invalid, expired, or already used will fail without blocking others.',
        ConfirmBatchOperationInput.shape,
        handle((args) => jsonResult(orgTools!.confirmBatchOperation(args))),
    );

    // =========================================================================
    // Mailbox Organization Tools (Low-Risk — Single Tool)
    // =========================================================================

    server.tool('mark_email_read', 'Mark an email as read in Outlook. Use this to clear unread indicators on a message. Returns a success confirmation. Returns an error if the email ID does not exist. Use mark_email_unread to reverse this action.',
        MarkEmailReadInput.shape,
        handle((args) => jsonResult(orgTools!.markEmailRead(args))),
    );

    server.tool('mark_email_unread', 'Mark an email as unread in Outlook. Use this to restore the unread indicator on a message. Returns a success confirmation. Returns an error if the email ID does not exist. Use mark_email_read to reverse this action.',
        MarkEmailUnreadInput.shape,
        handle((args) => jsonResult(orgTools!.markEmailUnread(args))),
    );

    server.tool('set_email_flag', 'Set the follow-up flag status on an email. Use flag_status 1 for flagged, 2 for completed, or 0 to clear. Returns a success confirmation. For a simpler clear-only operation, use clear_email_flag instead. Returns an error if the email ID does not exist.',
        SetEmailFlagInput.shape,
        handle((args) => jsonResult(orgTools!.setEmailFlag(args))),
    );

    server.tool('clear_email_flag', 'Clear the follow-up flag from an email, setting it to "not flagged". This is a convenience shortcut equivalent to set_email_flag with flag_status 0. Returns a success confirmation. Returns an error if the email ID does not exist.',
        ClearEmailFlagInput.shape,
        handle((args) => jsonResult(orgTools!.clearEmailFlag(args))),
    );

    server.tool('set_email_categories', 'Set categories on an email, replacing any existing categories. Use this to organize emails with color-coded labels. Pass an empty array to clear all categories. Returns a success confirmation. Returns an error if the email ID does not exist. Use get_email to check current categories before modifying.',
        SetEmailCategoriesInput.shape,
        handle((args) => jsonResult(orgTools!.setEmailCategories(args))),
    );

    server.tool('create_folder', 'Create a new mail folder in Outlook. Creates a top-level folder by default, or a subfolder if parent_folder_id is specified. Returns the created folder with id, name, messageCount, and unreadCount. Use list_folders to discover existing folder IDs for the parent. Returns an error if the parent folder does not exist.',
        CreateFolderInput.shape,
        handle((args) => jsonResult(orgTools!.createFolder(args))),
    );

    server.tool('rename_folder', 'Rename a mail folder in Outlook. Changes the display name of the folder without moving it. Returns a success confirmation with the new name. Returns an error if the folder ID does not exist. Use list_folders to find folder IDs.',
        RenameFolderInput.shape,
        handle((args) => jsonResult(orgTools!.renameFolder(args))),
    );

    server.tool('move_folder', 'Move a mail folder under a different parent folder in Outlook. Relocates the folder and all its contents to the new parent. Returns a success confirmation. Returns an error if either the source or destination folder ID does not exist. Use list_folders to find folder IDs.',
        MoveFolderInput.shape,
        handle((args) => jsonResult(orgTools!.moveFolder(args))),
    );

    return server;
}

// =============================================================================
// Helpers
// =============================================================================

/** Maps Zod-validated CreateEventInput fields to the CalendarWriter's internal param shape. */
function buildCalendarWriterParams(params: CreateEventParams): CalendarWriterParams {
    let recurrence: RecurrenceConfig | undefined;
    if (params.recurrence != null) {
        const rec = params.recurrence;
        recurrence = {
            frequency: rec.frequency,
            interval: rec.interval,
            ...(rec.days_of_week != null && { daysOfWeek: rec.days_of_week }),
            ...(rec.day_of_month != null && { dayOfMonth: rec.day_of_month }),
            ...(rec.week_of_month != null && { weekOfMonth: rec.week_of_month }),
            ...(rec.day_of_week_monthly != null && { dayOfWeekMonthly: rec.day_of_week_monthly }),
            ...(rec.end.type === 'end_date' && { endDate: rec.end.date }),
            ...(rec.end.type === 'end_after_count' && { endAfterCount: rec.end.count }),
        };
    }
    return {
        title: params.title,
        startDate: params.start_date,
        endDate: params.end_date,
        ...(params.calendar_id != null && { calendarId: params.calendar_id }),
        ...(params.location != null && { location: params.location }),
        ...(params.description != null && { description: params.description }),
        ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
        ...(recurrence != null && { recurrence }),
    };
}

function resolveAccountIds(accountId: unknown, accountRepository: IAccountRepository): number[] {
    if (accountId === undefined) {
        const defaultId = accountRepository.getDefaultAccountId();
        return defaultId !== null ? [defaultId] : [];
    }
    if (accountId === 'all') {
        const accounts = accountRepository.listAccounts();
        return accounts.map(acc => acc.id);
    }
    if (typeof accountId === 'number') {
        return [accountId];
    }
    if (Array.isArray(accountId)) {
        return accountId as number[];
    }
    const defaultId = accountRepository.getDefaultAccountId();
    return defaultId !== null ? [defaultId] : [];
}

// =============================================================================
// Main
// =============================================================================

async function main(): Promise<void> {
    const server = createServer();
    const transport = new StdioServerTransport();
    await server.connect(transport);
}

const isMainModule = import.meta.url === `file://${process.argv[1]}` ||
    import.meta.url.endsWith('/dist/index.js');
if (isMainModule) {
    main().catch((error: unknown) => {
        console.error('Fatal error:', error);
        process.exit(1);
    });
}
