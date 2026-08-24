import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';
import { appleTimestampToIso, isoToAppleTimestamp } from '../utils/dates.js';
import { createEmailPath, createEventPath, createContactPath, createTaskPath, createNotePath, } from './content-readers.js';
function priorityToNumber(priority) {
    switch (priority.toLowerCase()) {
        case 'high':
            return 1;
        case 'low':
            return -1;
        default:
            return 0;
    }
}
function toFolderRow(asFolder) {
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
function calendarToFolderRow(asCal) {
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
function toEmailRow(asEmail) {
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
function toEventRow(asEvent) {
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
function toContactRow(asContact) {
    return {
        id: asContact.id,
        folderId: 0,
        displayName: asContact.displayName,
        sortName: asContact.lastName ?? asContact.displayName,
        contactType: null,
        dataFilePath: createContactPath(asContact.id),
    };
}
function toTaskRow(asTask) {
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
function toNoteRow(asNote) {
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
export function deduplicateEmailRows(rows) {
    const seen = new Set();
    return rows.filter(r => {
        if (seen.has(r.id))
            return false;
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
export function searchTimeoutMs(offset) {
    // Phase 1 (subject whose-scan) alone measured 30-55s on large mailboxes;
    // Phase 2 (bounded sender loop) adds up to SENDER_SCAN_LIMIT * ~0.8s.
    // Base raised from 90s to 140s to give both phases headroom.
    return Math.min(180000, 140000 + Math.floor(offset / 25) * 10000);
}
export class AppleScriptRepository {
    folderCache = new Map();
    folderCacheExpiry = 0;
    CACHE_TTL_MS = 30000;
    listFolders() {
        const output = executeAppleScriptOrThrow(scripts.LIST_MAIL_FOLDERS);
        const folders = parser.parseFolders(output).map(toFolderRow);
        this.folderCache.clear();
        for (const folder of folders) {
            this.folderCache.set(folder.id, folder);
        }
        this.folderCacheExpiry = Date.now() + this.CACHE_TTL_MS;
        return folders;
    }
    getFolder(id) {
        if (Date.now() < this.folderCacheExpiry) {
            const cached = this.folderCache.get(id);
            if (cached != null) {
                return cached;
            }
        }
        const folders = this.listFolders();
        return folders.find((f) => f.id === id);
    }
    listEmails(folderId, limit, offset, after, before) {
        const script = scripts.listMessages(folderId, limit, offset, false, after, before);
        // Email listing is slow on large folders — use 45s timeout; date-filtered queries may be slower
        const timeoutMs = (after != null || before != null) ? 60000 : 45000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }
    listUnreadEmails(folderId, limit, offset, after, before) {
        const script = scripts.listMessages(folderId, limit, offset, true, after, before);
        const timeoutMs = (after != null || before != null) ? 60000 : 45000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }
    searchEmails(query, limit, offset, after, before) {
        const script = scripts.searchMessages(query, null, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
        return deduplicateEmailRows(parser.parseEmails(output).map(toEmailRow));
    }
    searchEmailsInFolder(folderId, query, limit, offset, after, before) {
        const script = scripts.searchMessages(query, folderId, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
        return deduplicateEmailRows(parser.parseEmails(output).map(toEmailRow));
    }
    getEmail(id) {
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
    getUnreadCount() {
        const folders = this.listFolders();
        return folders.reduce((sum, f) => sum + f.unreadCount, 0);
    }
    getUnreadCountByFolder(folderId) {
        try {
            const script = scripts.getUnreadCount(folderId);
            const output = executeAppleScriptOrThrow(script);
            return parser.parseCount(output);
        }
        catch {
            return 0;
        }
    }
    listCalendars() {
        const output = executeAppleScriptOrThrow(scripts.LIST_CALENDARS);
        return parser.parseCalendars(output).map(calendarToFolderRow);
    }
    listEvents(limit, offset = 0) {
        const script = scripts.listEvents(null, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }
    listEventsByFolder(folderId, limit, offset = 0) {
        const script = scripts.listEvents(folderId, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }
    listEventsByDateRange(startDate, endDate, limit, offset = 0) {
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
    getEvent(id) {
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
    listContacts(limit, offset) {
        const script = scripts.listContacts(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }
    searchContacts(query, limit, offset = 0) {
        const script = scripts.searchContacts(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }
    getContact(id) {
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
    listTasks(limit, offset) {
        const script = scripts.listTasks(limit, offset, true);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    listIncompleteTasks(limit, offset) {
        const script = scripts.listTasks(limit, offset, false);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    searchTasks(query, limit, offset = 0) {
        const script = scripts.searchTasks(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    getTask(id) {
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
    listNotes(limit, offset) {
        const script = scripts.listNotes(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseNotes(output).map(toNoteRow);
    }
    searchNotes(query, limit, offset = 0) {
        const script = scripts.searchNotes(query, limit, offset);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseNotes(output).map(toNoteRow);
    }
    searchEvents(query, limit, offset = 0, after, before) {
        const script = scripts.searchEvents(query, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseEvents(output).map(toEventRow);
    }
    getNote(id) {
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
    moveEmail(emailId, destinationFolderId) {
        const script = scripts.moveMessage(emailId, destinationFolderId);
        executeAppleScriptOrThrow(script);
    }
    deleteEmail(emailId) {
        const script = scripts.deleteMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    archiveEmail(emailId) {
        const script = scripts.archiveMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    junkEmail(emailId) {
        const script = scripts.junkMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    markEmailRead(emailId, isRead) {
        const script = scripts.setMessageReadStatus(emailId, isRead);
        executeAppleScriptOrThrow(script);
    }
    setEmailFlag(emailId, flagStatus) {
        const script = scripts.setMessageFlag(emailId, flagStatus);
        executeAppleScriptOrThrow(script);
    }
    setEmailCategories(emailId, categories) {
        const script = scripts.setMessageCategories(emailId, categories);
        executeAppleScriptOrThrow(script);
    }
    createFolder(name, parentFolderId) {
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
    deleteFolder(folderId) {
        const script = scripts.deleteMailFolder(folderId);
        executeAppleScriptOrThrow(script);
    }
    renameFolder(folderId, newName) {
        const script = scripts.renameMailFolder(folderId, newName);
        executeAppleScriptOrThrow(script);
    }
    moveFolder(folderId, destinationParentId) {
        const script = scripts.moveMailFolder(folderId, destinationParentId);
        executeAppleScriptOrThrow(script);
    }
    emptyFolder(folderId) {
        const script = scripts.emptyMailFolder(folderId);
        executeAppleScriptOrThrow(script);
    }
}
export function createAppleScriptRepository() {
    return new AppleScriptRepository();
}
