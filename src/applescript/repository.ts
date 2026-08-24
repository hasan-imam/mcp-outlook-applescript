import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';
import { appleTimestampToIso, isoToAppleTimestamp } from '../utils/dates.js';
import { createEmailPath, createEventPath, createContactPath, createTaskPath, createNotePath, } from './content-readers.js';
import type { IWriteableRepository, FolderRow, EmailRow, EventRow, ContactRow, TaskRow, NoteRow } from '../database/repository.js';
import type { AppleScriptFolderRow, AppleScriptCalendarRow, AppleScriptEmailRow, AppleScriptEventRow, AppleScriptContactRow, AppleScriptTaskRow, AppleScriptNoteRow } from './parser.js';

function priorityToNumber(priority: string): number {
    switch (priority.toLowerCase()) {
        case 'high':
            return 1;
        case 'low':
            return -1;
        default:
            return 0;
    }
}

function toFolderRow(asFolder: AppleScriptFolderRow): FolderRow {
    return {
        id: asFolder.id,
        name: asFolder.name,
        parentId: null,
        specialType: 0,
        folderType: 1,
        accountId: 1,
        messageCount: 0,
        unreadCount: asFolder.unreadCount,
    };
}

function calendarToFolderRow(asCal: AppleScriptCalendarRow): FolderRow {
    return {
        id: asCal.id,
        name: asCal.name,
        parentId: null,
        specialType: 0,
        folderType: 2,
        accountId: 1,
        messageCount: 0,
        unreadCount: 0,
    };
}

function toEmailRow(asEmail: AppleScriptEmailRow): EmailRow {
    return {
        id: asEmail.id,
        folderId: asEmail.folderId ?? 0,
        subject: asEmail.subject,
        sender: asEmail.senderName,
        senderAddress: asEmail.senderEmail,
        recipients: asEmail.toRecipients,
        displayTo: asEmail.toRecipients,
        toAddresses: asEmail.toRecipients,
        ccAddresses: asEmail.ccRecipients,
        preview: asEmail.preview,
        isRead: asEmail.isRead ? 1 : 0,
        timeReceived: isoToAppleTimestamp(asEmail.dateReceived),
        timeSent: isoToAppleTimestamp(asEmail.dateSent),
        hasAttachment: asEmail.attachments.length > 0 ? 1 : 0,
        size: 0,
        priority: priorityToNumber(asEmail.priority),
        flagStatus: asEmail.flagStatus ?? 0,
        categories: null,
        messageId: null,
        conversationId: null,
        dataFilePath: createEmailPath(asEmail.id),
    };
}

function toEventRow(asEvent: AppleScriptEventRow): EventRow {
    return {
        id: asEvent.id,
        folderId: asEvent.calendarId ?? 0,
        startDate: isoToAppleTimestamp(asEvent.startTime),
        endDate: isoToAppleTimestamp(asEvent.endTime),
        isRecurring: asEvent.isRecurring ? 1 : 0,
        hasReminder: 0,
        attendeeCount: asEvent.attendees.length,
        uid: null,
        masterRecordId: null,
        recurrenceId: null,
        dataFilePath: createEventPath(asEvent.id),
    };
}

function toContactRow(asContact: AppleScriptContactRow): ContactRow {
    return {
        id: asContact.id,
        folderId: 0,
        displayName: asContact.displayName,
        sortName: asContact.lastName ?? asContact.displayName,
        contactType: null,
        dataFilePath: createContactPath(asContact.id),
    };
}

function toTaskRow(asTask: AppleScriptTaskRow): TaskRow {
    return {
        id: asTask.id,
        folderId: asTask.folderId ?? 0,
        name: asTask.name,
        isCompleted: asTask.isCompleted ? 1 : 0,
        dueDate: isoToAppleTimestamp(asTask.dueDate),
        startDate: isoToAppleTimestamp(asTask.startDate),
        priority: priorityToNumber(asTask.priority),
        hasReminder: null,
        dataFilePath: createTaskPath(asTask.id),
    };
}

function toNoteRow(asNote: AppleScriptNoteRow): NoteRow {
    return {
        id: asNote.id,
        folderId: asNote.folderId ?? 0,
        modifiedDate: isoToAppleTimestamp(asNote.modifiedDate),
        dataFilePath: createNotePath(asNote.id),
    };
}

/**
 * Deduplicate email rows by ID, preserving first occurrence.
 * searchMessages phases 1 and 2 may return overlapping results — this is
 * intentional: doing dedup in TypeScript (O(n) via Set) instead of AppleScript
 * (O(n × offset) via list scans) eliminates the main performance bottleneck.
 */
export function deduplicateEmailRows(rows: EmailRow[]): EmailRow[] {
    const seen = new Set<number>();
    return rows.filter(r => {
        if (seen.has(r.id)) return false;
        seen.add(r.id);
        return true;
    });
}

/**
 * Calculate timeout for search operations, scaling with offset.
 * Base 140s (phase 1 whose-scan ~30-55s + phase 2 bounded sender scan
 * up to SENDER_SCAN_LIMIT messages at ~0.8s each). +10s per page after
 * the first, capped at 180s.
 */
export function searchTimeoutMs(offset: number): number {
    // Phase 1 (subject whose-scan) alone measured 30-55s on large mailboxes;
    // Phase 2 (bounded sender loop) adds up to SENDER_SCAN_LIMIT * ~0.8s.
    // Base raised from 90s to 140s to give both phases headroom.
    return Math.min(180000, 140000 + Math.floor(offset / 25) * 10000);
}

export class AppleScriptRepository implements IWriteableRepository {
    private readonly folderCache = new Map<number, FolderRow>();
    private folderCacheExpiry = 0;
    private readonly CACHE_TTL_MS = 30000;

    listFolders(): FolderRow[] {
        const output = executeAppleScriptOrThrow(scripts.LIST_MAIL_FOLDERS);
        const folders = parser.parseFolders(output).map(toFolderRow);
        this.folderCache.clear();
        for (const folder of folders) {
            this.folderCache.set(folder.id, folder);
        }
        this.folderCacheExpiry = Date.now() + this.CACHE_TTL_MS;
        return folders;
    }

    getFolder(id: number): FolderRow | undefined {
        if (Date.now() < this.folderCacheExpiry) {
            const cached = this.folderCache.get(id);
            if (cached != null) {
                return cached;
            }
        }
        const folders = this.listFolders();
        return folders.find((f) => f.id === id);
    }

    listEmails(folderId: number, limit: number, offset: number, after?: string, before?: string): EmailRow[] {
        const script = scripts.listMessages(folderId, limit, offset, false, after, before);
        // Email listing is slow on large folders — use 45s timeout; date-filtered queries may be slower
        const timeoutMs = (after != null || before != null) ? 60000 : 45000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }

    listUnreadEmails(folderId: number, limit: number, offset: number, after?: string, before?: string): EmailRow[] {
        const script = scripts.listMessages(folderId, limit, offset, true, after, before);
        const timeoutMs = (after != null || before != null) ? 60000 : 45000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }

    searchEmails(query: string, limit: number, offset: number, after?: string, before?: string): EmailRow[] {
        const script = scripts.searchMessages(query, null, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
        return deduplicateEmailRows(parser.parseEmails(output).map(toEmailRow));
    }

    searchEmailsInFolder(folderId: number, query: string, limit: number, offset: number, after?: string, before?: string): EmailRow[] {
        const script = scripts.searchMessages(query, folderId, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
        return deduplicateEmailRows(parser.parseEmails(output).map(toEmailRow));
    }

    getEmail(id: number): EmailRow | undefined {
        try {
            const script = scripts.getMessage(id);
            const output = executeAppleScriptOrThrow(script);
            const email = parser.parseEmail(output);
            return email != null ? toEmailRow(email) : undefined;
        }
        catch {
            return undefined;
        }
    }

    getUnreadCount(): number {
        const folders = this.listFolders();
        return folders.reduce((sum, f) => sum + f.unreadCount, 0);
    }

    getUnreadCountByFolder(folderId: number): number {
        try {
            const script = scripts.getUnreadCount(folderId);
            const output = executeAppleScriptOrThrow(script);
            return parser.parseCount(output);
        }
        catch {
            return 0;
        }
    }

    listCalendars(): FolderRow[] {
        const output = executeAppleScriptOrThrow(scripts.LIST_CALENDARS);
        return parser.parseCalendars(output).map(calendarToFolderRow);
    }

    listEvents(limit: number, offset: number = 0): EventRow[] {
        const script = scripts.listEvents(null, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }

    listEventsByFolder(folderId: number, limit: number, offset: number = 0): EventRow[] {
        const script = scripts.listEvents(folderId, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }

    listEventsByDateRange(startDate: number, endDate: number, limit: number, offset: number = 0): EventRow[] {
        // Server-side date filtering via AppleScript whose clause
        const startIso = appleTimestampToIso(startDate);
        const endIso = appleTimestampToIso(endDate);
        if (startIso == null || endIso == null) {
            return this.listEvents(limit, offset);
        }
        const script = scripts.listEvents(null, startIso, endIso, limit, offset);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 60000 });
        return parser.parseEvents(output).map(toEventRow);
    }

    getEvent(id: number): EventRow | undefined {
        try {
            const script = scripts.getEvent(id);
            const output = executeAppleScriptOrThrow(script);
            const event = parser.parseEvent(output);
            return event != null ? toEventRow(event) : undefined;
        }
        catch {
            return undefined;
        }
    }

    listContacts(limit: number, offset: number): ContactRow[] {
        const script = scripts.listContacts(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }

    searchContacts(query: string, limit: number, offset: number = 0): ContactRow[] {
        const script = scripts.searchContacts(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }

    getContact(id: number): ContactRow | undefined {
        try {
            const script = scripts.getContact(id);
            const output = executeAppleScriptOrThrow(script);
            const contact = parser.parseContact(output);
            return contact != null ? toContactRow(contact) : undefined;
        }
        catch {
            return undefined;
        }
    }

    listTasks(limit: number, offset: number): TaskRow[] {
        const script = scripts.listTasks(limit, offset, true);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }

    listIncompleteTasks(limit: number, offset: number): TaskRow[] {
        const script = scripts.listTasks(limit, offset, false);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }

    searchTasks(query: string, limit: number, offset: number = 0): TaskRow[] {
        const script = scripts.searchTasks(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }

    getTask(id: number): TaskRow | undefined {
        try {
            const script = scripts.getTask(id);
            const output = executeAppleScriptOrThrow(script);
            const task = parser.parseTask(output);
            return task != null ? toTaskRow(task) : undefined;
        }
        catch {
            return undefined;
        }
    }

    listNotes(limit: number, offset: number): NoteRow[] {
        const script = scripts.listNotes(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseNotes(output).map(toNoteRow);
    }

    searchNotes(query: string, limit: number, offset: number = 0): NoteRow[] {
        const script = scripts.searchNotes(query, limit, offset);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseNotes(output).map(toNoteRow);
    }

    searchEvents(query: string, limit: number, offset: number = 0, after?: string, before?: string): EventRow[] {
        const script = scripts.searchEvents(query, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseEvents(output).map(toEventRow);
    }

    getNote(id: number): NoteRow | undefined {
        try {
            const script = scripts.getNote(id);
            const output = executeAppleScriptOrThrow(script);
            const note = parser.parseNote(output);
            return note != null ? toNoteRow(note) : undefined;
        }
        catch {
            return undefined;
        }
    }

    moveEmail(emailId: number, destinationFolderId: number): void {
        const script = scripts.moveMessage(emailId, destinationFolderId);
        executeAppleScriptOrThrow(script);
    }

    deleteEmail(emailId: number): void {
        const script = scripts.deleteMessage(emailId);
        executeAppleScriptOrThrow(script);
    }

    archiveEmail(emailId: number): void {
        const script = scripts.archiveMessage(emailId);
        executeAppleScriptOrThrow(script);
    }

    junkEmail(emailId: number): void {
        const script = scripts.junkMessage(emailId);
        executeAppleScriptOrThrow(script);
    }

    markEmailRead(emailId: number, isRead: boolean): void {
        const script = scripts.setMessageReadStatus(emailId, isRead);
        executeAppleScriptOrThrow(script);
    }

    setEmailFlag(emailId: number, flagStatus: number): void {
        const script = scripts.setMessageFlag(emailId, flagStatus);
        executeAppleScriptOrThrow(script);
    }

    setEmailCategories(emailId: number, categories: string[]): void {
        const script = scripts.setMessageCategories(emailId, categories);
        executeAppleScriptOrThrow(script);
    }

    createFolder(name: string, parentFolderId?: number): FolderRow {
        const script = scripts.createMailFolder(name, parentFolderId);
        const output = executeAppleScriptOrThrow(script);
        const newFolderId = parseInt(output.trim(), 10);
        return {
            id: newFolderId,
            name,
            parentId: parentFolderId ?? null,
            specialType: 0,
            folderType: 1,
            accountId: 1,
            messageCount: 0,
            unreadCount: 0,
        };
    }

    deleteFolder(folderId: number): void {
        const script = scripts.deleteMailFolder(folderId);
        executeAppleScriptOrThrow(script);
    }

    renameFolder(folderId: number, newName: string): void {
        const script = scripts.renameMailFolder(folderId, newName);
        executeAppleScriptOrThrow(script);
    }

    moveFolder(folderId: number, destinationParentId: number): void {
        const script = scripts.moveMailFolder(folderId, destinationParentId);
        executeAppleScriptOrThrow(script);
    }

    emptyFolder(folderId: number): void {
        const script = scripts.emptyMailFolder(folderId);
        executeAppleScriptOrThrow(script);
    }
}

export function createAppleScriptRepository(): IWriteableRepository {
    return new AppleScriptRepository();
}
