/**
 * AppleScript template strings for Outlook operations.
 *
 * All scripts output data in a delimiter-based format for reliable parsing:
 * - Records are separated by {{RECORD}}
 * - Fields are separated by {{FIELD}}
 * - Field names and values are separated by {{=}}
 *
 * Example output: {{RECORD}}id{{=}}1{{FIELD}}name{{=}}Inbox{{FIELD}}unread{{=}}5{{RECORD}}...
 */
import { escapeForAppleScript } from './executor.js';
import { isoToDateComponents } from '../utils/dates.js';
// =============================================================================
// Shared AppleScript Blocks
// =============================================================================
/** AppleScript block: read todo flag of message variable `m`, store in `mFlag` as "0"/"1"/"2" */
const FLAG_STATUS_BLOCK = `
      set mFlag to "0"
      try
        set f to (todo flag of m) as string
        if f is "completed" then
          set mFlag to "2"
        else if f is "not flagged" then
          set mFlag to "0"
        else
          set mFlag to "1"
        end if
      end try`;
// =============================================================================
// Date Construction Helper
// =============================================================================
/**
 * Builds AppleScript statements to construct a date variable from ISO components.
 * Uses component-based date construction for locale safety (same pattern as createEvent).
 */
export function buildAppleScriptDateVar(varName, isoString) {
    const { year, month, day, hours, minutes } = isoToDateComponents(isoString);
    return [
        `set ${varName} to current date`,
        `set day of ${varName} to 1`,
        `set year of ${varName} to ${year}`,
        `set month of ${varName} to ${month}`,
        `set day of ${varName} to ${day}`,
        `set hours of ${varName} to ${hours}`,
        `set minutes of ${varName} to ${minutes}`,
        `set seconds of ${varName} to 0`,
    ].join('\n  ');
}
// =============================================================================
// Delimiters
// =============================================================================
export const DELIMITERS = {
    RECORD: '{{RECORD}}',
    FIELD: '{{FIELD}}',
    EQUALS: '{{=}}',
    NULL: '{{NULL}}',
};
// =============================================================================
// Shared AppleScript Output Templates
// =============================================================================
/**
 * AppleScript output expression for a message summary record.
 * Expects local variables: mId, mSubject, mSender, mSenderName, mDate, mRead, mPreview, mFlag.
 * listMessages also has mPriority — it appends that field separately.
 */
const MESSAGE_SUMMARY_OUTPUT = '"{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDate & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}preview{{=}}" & mPreview & "{{FIELD}}flagStatus{{=}}" & mFlag';
/**
 * AppleScript output expression for a contact summary record.
 * Expects local variables: cId, cDisplay, cFirst, cLast, cCompany, cEmail.
 */
const CONTACT_LIST_OUTPUT = '"{{RECORD}}id{{=}}" & cId & "{{FIELD}}displayName{{=}}" & cDisplay & "{{FIELD}}firstName{{=}}" & cFirst & "{{FIELD}}lastName{{=}}" & cLast & "{{FIELD}}company{{=}}" & cCompany & "{{FIELD}}email{{=}}" & cEmail';
/**
 * AppleScript block for safely extracting a truncated preview from an email.
 * The old pattern failed because `text 1 thru N` throws when content is shorter
 * than N chars, and the fallback returned the FULL plain text content (unbounded).
 * This version always truncates to 500 chars max.
 *
 * Used by listMessages only — searchMessages skips preview for performance
 * (the LLM calls get_email for full content on interesting search results).
 */
const PREVIEW_EXTRACT_BLOCK = `      set mPreview to ""
      try
        set rawContent to plain text content of m
        if (count of rawContent) > 500 then
          set mPreview to text 1 thru 500 of rawContent
        else
          set mPreview to rawContent
        end if
      end try`;
// =============================================================================
// Mail Scripts
// =============================================================================
/**
 * Lists all mail folders with their properties.
 */
export const LIST_MAIL_FOLDERS = `
tell application "Microsoft Outlook"
  set output to ""
  set allFolders to mail folders
  repeat with f in allFolders
    try
      set fId to id of f
      set fName to name of f
      set uCount to unread count of f
      set output to output & "{{RECORD}}id{{=}}" & fId & "{{FIELD}}name{{=}}" & fName & "{{FIELD}}unreadCount{{=}}" & uCount
    end try
  end repeat
  return output
end tell
`;
/**
 * Gets messages from a specific folder.
 * When after/before are provided, uses a `whose` clause with date comparison.
 */
export function listMessages(folderId, limit, offset, unreadOnly, after, before) {
    const totalToFetch = limit + offset;
    const hasDateFilter = after != null || before != null;
    let dateVarBlock = '';
    if (hasDateFilter) {
        if (after != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('afterDate', after)}\n`;
        if (before != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('beforeDate', before)}\n`;
    }
    let fetchBlock;
    if (hasDateFilter || unreadOnly) {
        // Build whose conditions
        const conditions = [];
        if (unreadOnly)
            conditions.push('is read is false');
        if (after != null)
            conditions.push('time received ≥ afterDate');
        if (before != null)
            conditions.push('time received ≤ beforeDate');
        fetchBlock = `  set allMsgs to (messages of targetFolder whose ${conditions.join(' and ')})
  set msgCount to count of allMsgs
  set endIdx to ${totalToFetch}
  if endIdx > msgCount then set endIdx to msgCount`;
    }
    else {
        fetchBlock = `  try
    set allMsgs to messages 1 thru ${totalToFetch} of targetFolder
  on error
    set allMsgs to every message of targetFolder
  end try
  set endIdx to count of allMsgs`;
    }
    return `
tell application "Microsoft Outlook"
  set output to ""
  set targetFolder to mail folder id ${folderId}
${dateVarBlock}${fetchBlock}
  set startIdx to ${offset + 1}
  if startIdx > endIdx then return ""

  repeat with i from startIdx to endIdx
    try
      set m to item i of allMsgs
      set mId to id of m
      set mSubject to subject of m
      set mSender to ""
      try
        set mSender to address of sender of m
      end try
      set mSenderName to ""
      try
        set mSenderName to name of sender of m
      end try
      set mDate to ""
      try
        set mDate to time received of m as «class isot» as string
      end try
      set mRead to is read of m
      set mPriority to "normal"
      try
        set p to priority of m
        if p is priority high then
          set mPriority to "high"
        else if p is priority low then
          set mPriority to "low"
        end if
      end try
${PREVIEW_EXTRACT_BLOCK}
${FLAG_STATUS_BLOCK}

      set output to output & "{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDate & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}priority{{=}}" & mPriority & "{{FIELD}}preview{{=}}" & mPreview & "{{FIELD}}flagStatus{{=}}" & mFlag
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Maximum number of messages to scan for sender-address matches.
 * Keeps execution within the 30s AppleScript timeout.
 */
const SENDER_SCAN_LIMIT = 500;
/**
 * Searches messages by query.
 *
 * Uses a two-phase approach to avoid crashes on emails with unresolvable sender
 * addresses (the AppleScript WHERE clause cannot use try/catch):
 *   Phase 1 — subject matches via a safe WHERE clause
 *   Phase 2 — sender matches via a loop with try/catch protection
 *
 * Performance optimizations:
 *   - No matchedIds list or dedup in AppleScript — dedup by ID happens in TypeScript
 *     (eliminates O(500 × offset) list scans and O(offset) Apple Events)
 *   - Phase 2 is skipped entirely when phase 1 has enough results to cover the page
 *   - No preview extraction — subject/sender/date is sufficient for search results;
 *     the LLM calls get_email for full content on interesting matches
 *
 * Offset is handled natively in AppleScript: phase 1 uses startIdx/endIdx on the
 * whose result list; phase 2 has its own skip counter independent of phase 1.
 * When after/before are provided, date filtering is applied in both phases.
 */
export function searchMessages(query, folderId, limit, offset = 0, after, before) {
    const escapedQuery = escapeForAppleScript(query);
    const folderClause = folderId != null ? `of mail folder id ${folderId}` : '';
    const hasDateFilter = after != null || before != null;
    let dateVarBlock = '';
    if (hasDateFilter) {
        if (after != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('afterDate', after)}\n`;
        if (before != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('beforeDate', before)}\n`;
    }
    // Build additional whose conditions for phase 1
    const whoseConditions = [`subject contains "${escapedQuery}"`];
    if (after != null)
        whoseConditions.push('time received ≥ afterDate');
    if (before != null)
        whoseConditions.push('time received ≤ beforeDate');
    const whoseClause = whoseConditions.join(' and ');
    // Build date check for phase 2 (loop-based sender scan)
    let phase2DateCheck = '';
    if (hasDateFilter) {
        const checks = [];
        if (after != null)
            checks.push('mDateObj < afterDate');
        if (before != null)
            checks.push('mDateObj > beforeDate');
        // If any date check fails, skip this message
        phase2DateCheck = `
          set mDateObj to time received of m
          if ${checks.join(' or ')} then
            -- skip: outside date range
          else`;
    }
    const phase2DateCheckEnd = hasDateFilter ? `
          end if` : '';
    return `
tell application "Microsoft Outlook"
  set output to ""
  set resultCount to 0
  set maxResults to ${limit}
  set skipCount to ${offset}
${dateVarBlock}
  -- Phase 1: Subject matches (safe WHERE clause)
  -- whose is fast (native); we skip the first 'offset' items and only extract what we need
  set subjectMatches to (messages ${folderClause} whose ${whoseClause})
  set phase1Total to count of subjectMatches
  set phase1Start to skipCount + 1
  set phase1End to skipCount + maxResults
  if phase1End > phase1Total then set phase1End to phase1Total

  if phase1Start ≤ phase1Total then
    repeat with i from phase1Start to phase1End
      if resultCount ≥ maxResults then exit repeat
      try
        set m to item i of subjectMatches
        set mId to id of m
        set mSubject to subject of m
        set mSender to ""
        try
          set mSender to address of sender of m
        end try
        set mSenderName to ""
        try
          set mSenderName to name of sender of m
        end try
        set mDate to ""
        try
          set mDate to time received of m as «class isot» as string
        end try
        set mRead to is read of m
        set mPreview to ""
${FLAG_STATUS_BLOCK}
        set output to output & ${MESSAGE_SUMMARY_OUTPUT}
        set resultCount to resultCount + 1
      end try
    end repeat
  end if

  -- Skip phase 2 if phase 1 has enough matches to cover this page + next
  if phase1Total > (skipCount + maxResults) then
    return output
  end if

  -- Phase 2: Sender matches (loop with try/catch for safety)
  -- Uses its own counter (phase2Count) independent of phase 1's resultCount.
  -- This prevents duplicates (messages matching both subject and sender) from
  -- consuming phase 2 slots and causing under-fetch after TS-side dedup.
  -- Note: phase2Skip counts raw sender matches to skip, not sender-only matches.
  -- When subject/sender overlaps exist in the skip window, effective page boundaries
  -- may shift slightly. This is acceptable for search pagination.
  set phase2Skip to skipCount - phase1Total
  if phase2Skip < 0 then set phase2Skip to 0
  set phase2Skipped to 0
  set phase2Count to 0

  if resultCount < maxResults then
    set allMsgs to messages ${folderClause}
    set scanLimit to count of allMsgs
    if scanLimit > ${SENDER_SCAN_LIMIT} then set scanLimit to ${SENDER_SCAN_LIMIT}
    repeat with i from 1 to scanLimit
      if phase2Count ≥ maxResults then exit repeat
      try
        set m to item i of allMsgs
        set mId to id of m
        set mSender to ""
        try
          set mSender to address of sender of m
        end try
        if mSender contains "${escapedQuery}" then${phase2DateCheck}
          if phase2Skipped < phase2Skip then
            set phase2Skipped to phase2Skipped + 1
          else
            set mSubject to subject of m
            set mSenderName to ""
            try
              set mSenderName to name of sender of m
            end try
            set mDate to ""
            try
              set mDate to time received of m as «class isot» as string
            end try
            set mRead to is read of m
            set mPreview to ""
${FLAG_STATUS_BLOCK}
            set output to output & ${MESSAGE_SUMMARY_OUTPUT}
            set phase2Count to phase2Count + 1
          end if${phase2DateCheckEnd}
        end if
      end try
    end repeat
  end if

  return output
end tell
`;
}
/**
 * Gets a single message by ID with full content.
 */
export function getMessage(messageId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set mId to id of m
  set mSubject to subject of m
  set mSender to ""
  try
    set mSender to address of sender of m
  end try
  set mSenderName to ""
  try
    set mSenderName to name of sender of m
  end try
  set mDateReceived to ""
  try
    set mDateReceived to time received of m as «class isot» as string
  end try
  set mDateSent to ""
  try
    set mDateSent to time sent of m as «class isot» as string
  end try
  set mRead to is read of m
  set mPriority to "normal"
  try
    set p to priority of m
    if p is priority high then
      set mPriority to "high"
    else if p is priority low then
      set mPriority to "low"
    end if
  end try
  set mHtml to ""
  try
    set mHtml to content of m
  end try
  set mPlain to ""
  try
    set mPlain to plain text content of m
  end try
  set mHasHtml to has html of m
  set mFolderId to ""
  try
    set mFolderId to id of folder of m
  end try

  -- Get recipients
  set toList to ""
  try
    repeat with r in to recipients of m
      set toList to toList & (address of r) & ","
    end repeat
  end try
  set ccList to ""
  try
    repeat with r in cc recipients of m
      set ccList to ccList & (address of r) & ","
    end repeat
  end try

  -- Get attachments
  set attachList to ""
  set attachDetailList to ""
  try
    set idx to 1
    repeat with a in attachments of m
      set aName to name of a
      set attachList to attachList & aName & ","
      set aSize to 0
      try
        set aSize to file size of a
      end try
      set aType to ""
      try
        set aType to content type of a
      end try
      set attachDetailList to attachDetailList & idx & "|" & aName & "|" & aSize & "|" & aType & ","
      set idx to idx + 1
    end repeat
  end try

${FLAG_STATUS_BLOCK}

  return "{{RECORD}}id{{=}}" & mId & "{{FIELD}}subject{{=}}" & mSubject & "{{FIELD}}senderEmail{{=}}" & mSender & "{{FIELD}}senderName{{=}}" & mSenderName & "{{FIELD}}dateReceived{{=}}" & mDateReceived & "{{FIELD}}dateSent{{=}}" & mDateSent & "{{FIELD}}isRead{{=}}" & mRead & "{{FIELD}}priority{{=}}" & mPriority & "{{FIELD}}htmlContent{{=}}" & mHtml & "{{FIELD}}plainContent{{=}}" & mPlain & "{{FIELD}}hasHtml{{=}}" & mHasHtml & "{{FIELD}}folderId{{=}}" & mFolderId & "{{FIELD}}toRecipients{{=}}" & toList & "{{FIELD}}ccRecipients{{=}}" & ccList & "{{FIELD}}attachments{{=}}" & attachList & "{{FIELD}}attachmentDetails{{=}}" & attachDetailList & "{{FIELD}}flagStatus{{=}}" & mFlag
end tell
`;
}
/**
 * Gets unread count for a folder.
 */
export function getUnreadCount(folderId) {
    return `
tell application "Microsoft Outlook"
  set f to mail folder id ${folderId}
  return unread count of f
end tell
`;
}
// =============================================================================
// Attachment Scripts
// =============================================================================
/**
 * Lists attachment metadata for a message.
 */
export function listAttachments(messageId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set output to ""
  set attachList to attachments of m
  set idx to 1
  repeat with a in attachList
    try
      set aName to name of a
      set aSize to 0
      try
        set aSize to file size of a
      end try
      set aType to ""
      try
        set aType to content type of a
      end try
      set output to output & "{{RECORD}}index{{=}}" & idx & "{{FIELD}}name{{=}}" & aName & "{{FIELD}}fileSize{{=}}" & aSize & "{{FIELD}}contentType{{=}}" & aType
    end try
    set idx to idx + 1
  end repeat
  return output
end tell
`;
}
/**
 * Saves an attachment from a message to a file path.
 * Uses 1-based attachment index.
 */
export function saveAttachment(messageId, attachmentIndex, savePath) {
    const escapedPath = escapeForAppleScript(savePath);
    return `
tell application "Microsoft Outlook"
  try
    set m to message id ${messageId}
    set attachList to attachments of m
    set attachCount to count of attachList
    if ${attachmentIndex} > attachCount then
      return "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Attachment index ${attachmentIndex} out of range (message has " & attachCount & " attachments)"
    end if
    set a to item ${attachmentIndex} of attachList
    set aName to name of a
    set aSize to 0
    try
      set aSize to file size of a
    end try
    save a in POSIX file "${escapedPath}"
    return "{{RECORD}}success{{=}}true{{FIELD}}name{{=}}" & aName & "{{FIELD}}savedTo{{=}}${escapedPath}" & "{{FIELD}}fileSize{{=}}" & aSize
  on error errMsg
    return "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
  end try
end tell
`;
}
// =============================================================================
// Calendar Scripts
// =============================================================================
/**
 * Lists all calendars.
 */
export const LIST_CALENDARS = `
tell application "Microsoft Outlook"
  set output to ""
  set allCalendars to calendars
  repeat with c in allCalendars
    try
      set cId to id of c
      set cName to name of c
      set output to output & "{{RECORD}}id{{=}}" & cId & "{{FIELD}}name{{=}}" & cName
    end try
  end repeat
  return output
end tell
`;
/**
 * Lists events from a calendar with optional date range and offset.
 * When startDate/endDate are provided, uses a `whose` clause for server-side filtering.
 */
export function listEvents(calendarId, startDate, endDate, limit, offset = 0) {
    const calendarClause = calendarId != null ? `of calendar id ${calendarId}` : '';
    const hasDateFilter = startDate != null || endDate != null;
    let dateVarBlock = '';
    if (hasDateFilter) {
        if (startDate != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('afterDate', startDate)}\n`;
        if (endDate != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('beforeDate', endDate)}\n`;
    }
    let fetchBlock;
    if (hasDateFilter) {
        const whoseConditions = [];
        if (startDate != null)
            whoseConditions.push('start time ≥ afterDate');
        if (endDate != null)
            whoseConditions.push('start time ≤ beforeDate');
        fetchBlock = `  set allEvents to (calendar events ${calendarClause} whose ${whoseConditions.join(' and ')})
  set eventCount to count of allEvents`;
    }
    else {
        fetchBlock = `  set allEvents to calendar events ${calendarClause}
  set eventCount to count of allEvents`;
    }
    return `
tell application "Microsoft Outlook"
  set output to ""
${dateVarBlock}${fetchBlock}
  set startIdx to ${offset + 1}
  set endIdx to ${offset + limit}
  if endIdx > eventCount then set endIdx to eventCount
  if startIdx > eventCount then return ""

  repeat with i from startIdx to endIdx
    try
      set e to item i of allEvents
      set eId to id of e
      set eSubject to subject of e
      set eStart to ""
      try
        set eStart to start time of e as «class isot» as string
      end try
      set eEnd to ""
      try
        set eEnd to end time of e as «class isot» as string
      end try
      set eLocation to ""
      try
        set eLocation to location of e
      end try
      set eAllDay to all day flag of e
      set eRecurring to is recurring of e

      set output to output & "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation & "{{FIELD}}isAllDay{{=}}" & eAllDay & "{{FIELD}}isRecurring{{=}}" & eRecurring
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Gets a single event by ID.
 */
export function getEvent(eventId) {
    return `
tell application "Microsoft Outlook"
  set e to calendar event id ${eventId}
  set eId to id of e
  set eSubject to subject of e
  set eStart to ""
  try
    set eStart to start time of e as «class isot» as string
  end try
  set eEnd to ""
  try
    set eEnd to end time of e as «class isot» as string
  end try
  set eLocation to ""
  try
    set eLocation to location of e
  end try
  set eContent to ""
  try
    set eContent to content of e
  end try
  set ePlain to ""
  try
    set ePlain to plain text content of e
  end try
  set eAllDay to all day flag of e
  set eRecurring to is recurring of e
  set eOrganizer to ""
  try
    set eOrganizer to organizer of e
  end try
  set eCalId to ""
  try
    set eCalId to id of calendar of e
  end try

  -- Get attendees
  set attendeeList to ""
  try
    repeat with a in attendees of e
      set aEmail to email address of a
      set aName to name of a
      set attendeeList to attendeeList & aEmail & "|" & aName & ","
    end repeat
  end try

  return "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation & "{{FIELD}}htmlContent{{=}}" & eContent & "{{FIELD}}plainContent{{=}}" & ePlain & "{{FIELD}}isAllDay{{=}}" & eAllDay & "{{FIELD}}isRecurring{{=}}" & eRecurring & "{{FIELD}}organizer{{=}}" & eOrganizer & "{{FIELD}}calendarId{{=}}" & eCalId & "{{FIELD}}attendees{{=}}" & attendeeList
end tell
`;
}
/**
 * Searches events by query with optional offset and date filtering.
 */
export function searchEvents(query, limit, offset, after, before) {
    const escapedQuery = escapeForAppleScript(query);
    const hasDateFilter = after != null || before != null;
    let dateVarBlock = '';
    if (hasDateFilter) {
        if (after != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('afterDate', after)}\n`;
        if (before != null)
            dateVarBlock += `  ${buildAppleScriptDateVar('beforeDate', before)}\n`;
    }
    // Build whose conditions
    const whoseConditions = [`subject contains "${escapedQuery}"`];
    if (after != null)
        whoseConditions.push('start time ≥ afterDate');
    if (before != null)
        whoseConditions.push('start time ≤ beforeDate');
    const whoseClause = whoseConditions.join(' and ');
    return `
tell application "Microsoft Outlook"
  set output to ""
${dateVarBlock}  set searchResults to (calendar events whose ${whoseClause})
  set resultCount to count of searchResults
  set startIdx to ${offset + 1}
  set endIdx to ${offset + limit}
  if endIdx > resultCount then set endIdx to resultCount
  if startIdx > resultCount then return ""

  repeat with i from startIdx to endIdx
    try
      set e to item i of searchResults
      set eId to id of e
      set eSubject to subject of e
      set eStart to ""
      try
        set eStart to start time of e as «class isot» as string
      end try
      set eEnd to ""
      try
        set eEnd to end time of e as «class isot» as string
      end try
      set eLocation to ""
      try
        set eLocation to location of e
      end try

      set output to output & "{{RECORD}}id{{=}}" & eId & "{{FIELD}}subject{{=}}" & eSubject & "{{FIELD}}startTime{{=}}" & eStart & "{{FIELD}}endTime{{=}}" & eEnd & "{{FIELD}}location{{=}}" & eLocation
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Builds AppleScript to set recurrence on a newly created event.
 * Assumes `newEvent` variable is already in scope from createEvent.
 */
function buildRecurrenceScript(params) {
    const isOrdinalMonthly = params.frequency === 'monthly' && params.weekOfMonth != null;
    const frequencyMap = {
        daily: 'daily recurrence',
        weekly: 'weekly recurrence',
        monthly: isOrdinalMonthly ? 'month nth recurrence' : 'monthly recurrence',
        yearly: 'yearly recurrence',
    };
    const recurrenceType = frequencyMap[params.frequency];
    const capitalize = (s) => s.charAt(0).toUpperCase() + s.slice(1);
    let script = `
  set is recurring of newEvent to true
  set theRecurrence to recurrence of newEvent
  set recurrence type of theRecurrence to ${recurrenceType}
  set occurrence interval of theRecurrence to ${params.interval}`;
    // Weekly: days of week mask
    if (params.frequency === 'weekly' && params.daysOfWeek != null) {
        const daysList = params.daysOfWeek.map(capitalize).join(', ');
        script += `\n  set day of week mask of theRecurrence to {${daysList}}`;
    }
    // Monthly by date
    if (params.frequency === 'monthly' && params.dayOfMonth != null && params.weekOfMonth == null) {
        script += `\n  set day of month of theRecurrence to ${params.dayOfMonth}`;
    }
    // Monthly ordinal (e.g., 3rd Thursday)
    if (params.frequency === 'monthly' && params.weekOfMonth != null && params.dayOfWeekMonthly != null) {
        const instanceMap = { first: 1, second: 2, third: 3, fourth: 4, last: 5 };
        const instance = instanceMap[params.weekOfMonth] ?? 1;
        script += `\n  set day of week mask of theRecurrence to {${capitalize(params.dayOfWeekMonthly)}}`;
        script += `\n  set instance of theRecurrence to ${instance}`;
    }
    // End after count
    if (params.endAfterCount != null) {
        script += `\n  set occurrences of theRecurrence to ${params.endAfterCount}`;
    }
    // End by date (component-based for locale safety)
    if (params.endDate != null) {
        script += `
  set theEndRecurrenceDate to current date
  set day of theEndRecurrenceDate to 1
  set year of theEndRecurrenceDate to ${params.endDate.year}
  set month of theEndRecurrenceDate to ${params.endDate.month}
  set day of theEndRecurrenceDate to ${params.endDate.day}
  set hours of theEndRecurrenceDate to ${params.endDate.hours}
  set minutes of theEndRecurrenceDate to ${params.endDate.minutes}
  set seconds of theEndRecurrenceDate to 0
  set pattern end date of theRecurrence to theEndRecurrenceDate`;
    }
    return script;
}
/**
 * Responds to an event invitation (RSVP).
 */
export function respondToEvent(params) {
    const { eventId, response, sendResponse, comment } = params;
    const commandMap = {
        accept: 'accept meeting',
        decline: 'decline meeting',
        tentative: 'accept tentatively meeting',
    };
    const responseCommand = commandMap[response];
    const responseOptions = comment != null
        ? ` comment "${escapeForAppleScript(comment)}" sending response ${sendResponse}`
        : ` sending response ${sendResponse}`;
    // Use Outlook's meeting-response commands so the organizer can be notified.
    // Setting the event response status alone does not reliably send an RSVP.
    const responseLine = `${responseCommand} myEvent${responseOptions}`;
    const statusMap = {
        accept: 'accepted',
        decline: 'declined',
        tentative: 'tentatively accepted',
    };
    const status = statusMap[response];
    const notified = sendResponse ? 'true' : 'false';
    const commentIncluded = comment != null ? 'true' : 'false';
    return `
tell application "Microsoft Outlook"
  try
    set myEvent to calendar event id ${eventId}
    ${responseLine}

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId} & "{{FIELD}}status{{=}}${status}{{FIELD}}sentResponse{{=}}${notified}{{FIELD}}commentIncluded{{=}}${commentIncluded}"
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
/**
 * Deletes an event. For recurring events, can delete single instance or entire series.
 */
export function deleteEvent(params) {
    const { eventId, applyTo } = params;
    const comment = applyTo === 'all_in_series'
        ? '-- Deleting entire series'
        : '-- Deleting single instance';
    return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}
    delete myEvent

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId}
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
/**
 * Updates an event. For recurring events, can update single instance or entire series.
 * All update fields are optional — only specified fields will be changed.
 */
export function updateEvent(params) {
    const { eventId, applyTo, updates } = params;
    const updatedFields = [];
    let updateStatements = '';
    if (updates.title != null) {
        updateStatements += `    set subject of myEvent to "${escapeForAppleScript(updates.title)}"\n`;
        updatedFields.push('title');
    }
    if (updates.location != null) {
        updateStatements += `    set location of myEvent to "${escapeForAppleScript(updates.location)}"\n`;
        updatedFields.push('location');
    }
    if (updates.description != null) {
        updateStatements += `    set content of myEvent to "${escapeForAppleScript(updates.description)}"\n`;
        updatedFields.push('description');
    }
    if (updates.startDate != null) {
        const start = isoToDateComponents(updates.startDate);
        updateStatements += `    set start time of myEvent to date "${start.year}-${start.month}-${start.day} ${start.hours}:${start.minutes}:00"\n`;
        updatedFields.push('startDate');
    }
    if (updates.endDate != null) {
        const end = isoToDateComponents(updates.endDate);
        updateStatements += `    set end time of myEvent to date "${end.year}-${end.month}-${end.day} ${end.hours}:${end.minutes}:00"\n`;
        updatedFields.push('endDate');
    }
    if (updates.isAllDay != null) {
        updateStatements += `    set all day flag of myEvent to ${updates.isAllDay}\n`;
        updatedFields.push('isAllDay');
    }
    const comment = applyTo === 'all_in_series'
        ? '-- Updating entire series'
        : '-- Updating single instance';
    const fieldsOutput = updatedFields.join(',');
    return `
tell application "Microsoft Outlook"
  try
    ${comment}
    set myEvent to calendar event id ${eventId}

${updateStatements}
    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}" & ${eventId} & "{{FIELD}}updatedFields{{=}}${fieldsOutput}"
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
/**
 * Creates a new calendar event.
 * Uses component-based date construction for locale safety.
 */
export function createEvent(params) {
    const escapedTitle = escapeForAppleScript(params.title);
    const escapedLocation = params.location != null ? escapeForAppleScript(params.location) : '';
    const escapedDescription = params.description != null ? escapeForAppleScript(params.description) : '';
    // Build properties list
    let properties = `subject:"${escapedTitle}", start time:theStartDate, end time:theEndDate`;
    if (params.location != null) {
        properties += `, location:"${escapedLocation}"`;
    }
    if (params.isAllDay === true) {
        properties += ', all day flag:true';
    }
    // Build the target clause for calendar
    const targetClause = params.calendarId != null
        ? `at calendar id ${params.calendarId} `
        : '';
    return `
tell application "Microsoft Outlook"
  set theStartDate to current date
  set day of theStartDate to 1
  set year of theStartDate to ${params.startYear}
  set month of theStartDate to ${params.startMonth}
  set day of theStartDate to ${params.startDay}
  set hours of theStartDate to ${params.startHours}
  set minutes of theStartDate to ${params.startMinutes}
  set seconds of theStartDate to 0

  set theEndDate to current date
  set day of theEndDate to 1
  set year of theEndDate to ${params.endYear}
  set month of theEndDate to ${params.endMonth}
  set day of theEndDate to ${params.endDay}
  set hours of theEndDate to ${params.endHours}
  set minutes of theEndDate to ${params.endMinutes}
  set seconds of theEndDate to 0

  set newEvent to make new calendar event ${targetClause}with properties {${properties}}
  ${escapedDescription ? `set plain text content of newEvent to "${escapedDescription}"` : ''}${params.recurrence != null ? buildRecurrenceScript(params.recurrence) : ''}
  set eId to id of newEvent
  set eCalId to ""
  try
    set eCalId to id of calendar of newEvent
  end try
  return "{{RECORD}}id{{=}}" & eId & "{{FIELD}}calendarId{{=}}" & eCalId
end tell
`;
}
// =============================================================================
// Contact Scripts
// =============================================================================
/**
 * Lists all contacts.
 */
export function listContacts(limit, offset) {
    const totalToFetch = limit + offset;
    return `
tell application "Microsoft Outlook"
  set output to ""
  set allContacts to contacts
  set contactCount to count of allContacts
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > contactCount then set endIdx to contactCount
  if startIdx > contactCount then return ""

  repeat with i from startIdx to endIdx
    try
      set c to item i of allContacts
      set cId to id of c
      set cDisplay to display name of c
      set cFirst to ""
      try
        set cFirst to first name of c
      end try
      set cLast to ""
      try
        set cLast to last name of c
      end try
      set cCompany to ""
      try
        set cCompany to company of c
      end try
      set cEmail to ""
      try
        set emailAddrs to email addresses of c
        if (count of emailAddrs) > 0 then
          set cEmail to address of item 1 of emailAddrs
        end if
      end try

      set output to output & ${CONTACT_LIST_OUTPUT}
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Searches contacts by name.
 */
export function searchContacts(query, limit, offset) {
    const escapedQuery = escapeForAppleScript(query);
    return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (contacts whose display name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set startIdx to ${offset + 1}
  set endIdx to ${offset + limit}
  if endIdx > resultCount then set endIdx to resultCount
  if startIdx > resultCount then return ""

  repeat with i from startIdx to endIdx
    try
      set c to item i of searchResults
      set cId to id of c
      set cDisplay to display name of c
      set cFirst to ""
      try
        set cFirst to first name of c
      end try
      set cLast to ""
      try
        set cLast to last name of c
      end try
      set cCompany to ""
      try
        set cCompany to company of c
      end try
      set cEmail to ""
      try
        set emailAddrs to email addresses of c
        if (count of emailAddrs) > 0 then
          set cEmail to address of item 1 of emailAddrs
        end if
      end try

      set output to output & ${CONTACT_LIST_OUTPUT}
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Gets a single contact by ID with full details.
 */
export function getContact(contactId) {
    return `
tell application "Microsoft Outlook"
  set c to contact id ${contactId}
  set cId to id of c
  set cDisplay to display name of c
  set cFirst to ""
  try
    set cFirst to first name of c
  end try
  set cLast to ""
  try
    set cLast to last name of c
  end try
  set cMiddle to ""
  try
    set cMiddle to middle name of c
  end try
  set cNickname to ""
  try
    set cNickname to nickname of c
  end try
  set cCompany to ""
  try
    set cCompany to company of c
  end try
  set cTitle to ""
  try
    set cTitle to job title of c
  end try
  set cDept to ""
  try
    set cDept to department of c
  end try
  set cNotes to ""
  try
    set cNotes to description of c
  end try

  -- Phones
  set cHomePhone to ""
  try
    set cHomePhone to home phone number of c
  end try
  set cWorkPhone to ""
  try
    set cWorkPhone to business phone number of c
  end try
  set cMobile to ""
  try
    set cMobile to mobile number of c
  end try

  -- Emails
  set emailList to ""
  try
    repeat with e in email addresses of c
      set emailList to emailList & (address of e) & ","
    end repeat
  end try

  -- Address
  set cHomeStreet to ""
  try
    set cHomeStreet to home street address of c
  end try
  set cHomeCity to ""
  try
    set cHomeCity to home city of c
  end try
  set cHomeState to ""
  try
    set cHomeState to home state of c
  end try
  set cHomeZip to ""
  try
    set cHomeZip to home zip of c
  end try
  set cHomeCountry to ""
  try
    set cHomeCountry to home country of c
  end try

  return "{{RECORD}}id{{=}}" & cId & "{{FIELD}}displayName{{=}}" & cDisplay & "{{FIELD}}firstName{{=}}" & cFirst & "{{FIELD}}lastName{{=}}" & cLast & "{{FIELD}}middleName{{=}}" & cMiddle & "{{FIELD}}nickname{{=}}" & cNickname & "{{FIELD}}company{{=}}" & cCompany & "{{FIELD}}jobTitle{{=}}" & cTitle & "{{FIELD}}department{{=}}" & cDept & "{{FIELD}}notes{{=}}" & cNotes & "{{FIELD}}homePhone{{=}}" & cHomePhone & "{{FIELD}}workPhone{{=}}" & cWorkPhone & "{{FIELD}}mobilePhone{{=}}" & cMobile & "{{FIELD}}emails{{=}}" & emailList & "{{FIELD}}homeStreet{{=}}" & cHomeStreet & "{{FIELD}}homeCity{{=}}" & cHomeCity & "{{FIELD}}homeState{{=}}" & cHomeState & "{{FIELD}}homeZip{{=}}" & cHomeZip & "{{FIELD}}homeCountry{{=}}" & cHomeCountry
end tell
`;
}
// =============================================================================
// Task Scripts
// =============================================================================
/**
 * Lists all tasks.
 */
export function listTasks(limit, offset, includeCompleted) {
    const totalToFetch = limit + offset;
    const completedFilter = includeCompleted ? '' : ' whose todo flag is not completed';
    return `
tell application "Microsoft Outlook"
  set output to ""
  set allTasks to (tasks${completedFilter})
  set taskCount to count of allTasks
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > taskCount then set endIdx to taskCount
  if startIdx > taskCount then return ""

  repeat with i from startIdx to endIdx
    try
      set t to item i of allTasks
      set tId to id of t
      set tName to name of t
      set tDue to ""
      try
        set tDue to due date of t as «class isot» as string
      end try
      set tCompleted to (todo flag of t is completed)
      set tPriority to "normal"
      try
        set p to priority of t
        if p is priority high then
          set tPriority to "high"
        else if p is priority low then
          set tPriority to "low"
        end if
      end try

      set output to output & "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}isCompleted{{=}}" & tCompleted & "{{FIELD}}priority{{=}}" & tPriority
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Searches tasks by name.
 */
export function searchTasks(query, limit, offset) {
    const escapedQuery = escapeForAppleScript(query);
    return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (tasks whose name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set startIdx to ${offset + 1}
  set endIdx to ${offset + limit}
  if endIdx > resultCount then set endIdx to resultCount
  if startIdx > resultCount then return ""

  repeat with i from startIdx to endIdx
    try
      set t to item i of searchResults
      set tId to id of t
      set tName to name of t
      set tDue to ""
      try
        set tDue to due date of t as «class isot» as string
      end try
      set tCompleted to (todo flag of t is completed)

      set output to output & "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}isCompleted{{=}}" & tCompleted
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Gets a single task by ID.
 */
export function getTask(taskId) {
    return `
tell application "Microsoft Outlook"
  set t to task id ${taskId}
  set tId to id of t
  set tName to name of t
  set tContent to ""
  try
    set tContent to content of t
  end try
  set tPlain to ""
  try
    set tPlain to plain text content of t
  end try
  set tDue to ""
  try
    set tDue to due date of t as «class isot» as string
  end try
  set tStart to ""
  try
    set tStart to start date of t as «class isot» as string
  end try
  set tCompletedDate to ""
  try
    set tCompletedDate to completed date of t as «class isot» as string
  end try
  set tCompleted to (todo flag of t is completed)
  set tPriority to "normal"
  try
    set p to priority of t
    if p is priority high then
      set tPriority to "high"
    else if p is priority low then
      set tPriority to "low"
    end if
  end try
  set tFolderId to ""
  try
    set tFolderId to id of folder of t
  end try

  return "{{RECORD}}id{{=}}" & tId & "{{FIELD}}name{{=}}" & tName & "{{FIELD}}htmlContent{{=}}" & tContent & "{{FIELD}}plainContent{{=}}" & tPlain & "{{FIELD}}dueDate{{=}}" & tDue & "{{FIELD}}startDate{{=}}" & tStart & "{{FIELD}}completedDate{{=}}" & tCompletedDate & "{{FIELD}}isCompleted{{=}}" & tCompleted & "{{FIELD}}priority{{=}}" & tPriority & "{{FIELD}}folderId{{=}}" & tFolderId
end tell
`;
}
// =============================================================================
// Note Scripts
// =============================================================================
/**
 * Lists all notes.
 */
export function listNotes(limit, offset) {
    const totalToFetch = limit + offset;
    return `
tell application "Microsoft Outlook"
  set output to ""
  set allNotes to notes
  set noteCount to count of allNotes
  set startIdx to ${offset + 1}
  set endIdx to ${totalToFetch}
  if endIdx > noteCount then set endIdx to noteCount
  if startIdx > noteCount then return ""

  repeat with i from startIdx to endIdx
    try
      set n to item i of allNotes
      set nId to id of n
      set nName to name of n
      set nCreated to ""
      try
        set nCreated to creation date of n as «class isot» as string
      end try
      set nModified to ""
      try
        set nModified to modification date of n as «class isot» as string
      end try
      set nPreview to ""
      try
        set nPreview to text 1 thru 200 of plain text content of n
      on error
        try
          set nPreview to plain text content of n
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}modifiedDate{{=}}" & nModified & "{{FIELD}}preview{{=}}" & nPreview
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Searches notes by name/title (does not search body content).
 */
export function searchNotes(query, limit, offset) {
    const escapedQuery = escapeForAppleScript(query);
    return `
tell application "Microsoft Outlook"
  set output to ""
  set searchResults to (notes whose name contains "${escapedQuery}")
  set resultCount to count of searchResults
  set startIdx to ${offset + 1}
  set endIdx to ${offset + limit}
  if endIdx > resultCount then set endIdx to resultCount
  if startIdx > resultCount then return ""

  repeat with i from startIdx to endIdx
    try
      set n to item i of searchResults
      set nId to id of n
      set nName to name of n
      set nCreated to ""
      try
        set nCreated to creation date of n as «class isot» as string
      end try
      set nPreview to ""
      try
        set nPreview to text 1 thru 200 of plain text content of n
      on error
        try
          set nPreview to plain text content of n
        end try
      end try

      set output to output & "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}preview{{=}}" & nPreview
    end try
  end repeat
  return output
end tell
`;
}
/**
 * Gets a single note by ID.
 */
export function getNote(noteId) {
    return `
tell application "Microsoft Outlook"
  set n to note id ${noteId}
  set nId to id of n
  set nName to name of n
  set nContent to ""
  try
    set nContent to content of n
  end try
  set nPlain to ""
  try
    set nPlain to plain text content of n
  end try
  set nCreated to ""
  try
    set nCreated to creation date of n as «class isot» as string
  end try
  set nModified to ""
  try
    set nModified to modification date of n as «class isot» as string
  end try
  set nFolderId to ""
  try
    set nFolderId to id of folder of n
  end try

  return "{{RECORD}}id{{=}}" & nId & "{{FIELD}}name{{=}}" & nName & "{{FIELD}}htmlContent{{=}}" & nContent & "{{FIELD}}plainContent{{=}}" & nPlain & "{{FIELD}}createdDate{{=}}" & nCreated & "{{FIELD}}modifiedDate{{=}}" & nModified & "{{FIELD}}folderId{{=}}" & nFolderId
end tell
`;
}
// Write Operation Scripts
// =============================================================================
/**
 * Moves a message to a different folder.
 */
export function moveMessage(messageId, destinationFolderId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set targetFolder to mail folder id ${destinationFolderId}
  move m to targetFolder
  return "ok"
end tell
`;
}
/**
 * Moves a message to the Deleted Items folder.
 */
export function deleteMessage(messageId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set deletedFolder to deleted items
  move m to deletedFolder
  return "ok"
end tell
`;
}
/**
 * Moves a message to the Archive folder.
 * Falls back to an "Archive" named folder if the well-known folder isn't available.
 */
export function archiveMessage(messageId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  try
    set archiveFolder to mail folder "Archive"
    move m to archiveFolder
  on error
    -- Try finding by name
    set allFolders to mail folders
    repeat with f in allFolders
      if name of f is "Archive" then
        move m to f
        return "ok"
      end if
    end repeat
    error "Archive folder not found"
  end try
  return "ok"
end tell
`;
}
/**
 * Moves a message to the Junk folder.
 */
export function junkMessage(messageId) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set junkFolder to junk mail
  move m to junkFolder
  return "ok"
end tell
`;
}
/**
 * Sets the read status of a message.
 */
export function setMessageReadStatus(messageId, isRead) {
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set is read of m to ${isRead}
  return "ok"
end tell
`;
}
/**
 * Sets the flag status of a message.
 * flagStatus: 0 = none, 1 = flagged, 2 = completed
 */
export function setMessageFlag(messageId, flagStatus) {
    let flagValue;
    switch (flagStatus) {
        case 1:
            flagValue = 'not completed';
            break;
        case 2:
            flagValue = 'completed';
            break;
        default:
            flagValue = 'not flagged';
    }
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set todo flag of m to ${flagValue}
  return "ok"
end tell
`;
}
/**
 * Sets categories on a message.
 */
export function setMessageCategories(messageId, categories) {
    const categoryList = categories.map(c => `"${escapeForAppleScript(c)}"`).join(', ');
    return `
tell application "Microsoft Outlook"
  set m to message id ${messageId}
  set category of m to {${categoryList}}
  return "ok"
end tell
`;
}
/**
 * Creates a new mail folder.
 */
export function createMailFolder(name, parentFolderId) {
    const escapedName = escapeForAppleScript(name);
    if (parentFolderId != null) {
        return `
tell application "Microsoft Outlook"
  set parentFolder to mail folder id ${parentFolderId}
  set newFolder to make new mail folder at parentFolder with properties {name:"${escapedName}"}
  return id of newFolder
end tell
`;
    }
    return `
tell application "Microsoft Outlook"
  set newFolder to make new mail folder with properties {name:"${escapedName}"}
  return id of newFolder
end tell
`;
}
/**
 * Deletes a mail folder.
 */
export function deleteMailFolder(folderId) {
    return `
tell application "Microsoft Outlook"
  set f to mail folder id ${folderId}
  delete f
  return "ok"
end tell
`;
}
/**
 * Renames a mail folder.
 */
export function renameMailFolder(folderId, newName) {
    const escapedName = escapeForAppleScript(newName);
    return `
tell application "Microsoft Outlook"
  set f to mail folder id ${folderId}
  set name of f to "${escapedName}"
  return "ok"
end tell
`;
}
/**
 * Moves a mail folder to a new parent.
 */
export function moveMailFolder(folderId, destinationParentId) {
    return `
tell application "Microsoft Outlook"
  set f to mail folder id ${folderId}
  set targetParent to mail folder id ${destinationParentId}
  move f to targetParent
  return "ok"
end tell
`;
}
/**
 * Deletes all messages in a folder (empties the folder).
 */
export function emptyMailFolder(folderId) {
    return `
tell application "Microsoft Outlook"
  set targetFolder to mail folder id ${folderId}
  set msgs to messages of targetFolder
  set deletedFolder to deleted items
  repeat with m in msgs
    try
      move m to deletedFolder
    end try
  end repeat
  return "ok"
end tell
`;
}
/**
 * Sends an email with optional CC, BCC, attachments, and account selection.
 */
export function sendEmail(params) {
    const { to, subject, body, bodyType, cc, bcc, replyTo, attachments, inlineImages, accountId } = params;
    const escapedSubject = escapeForAppleScript(subject);
    const escapedBody = escapeForAppleScript(body);
    const toRecipients = to.map(email => `    make new recipient at newMessage with properties {email address:{address:"${escapeForAppleScript(email)}"}}`).join('\n');
    const ccRecipients = cc != null && cc.length > 0
        ? cc.map(email => `    make new recipient at newMessage with properties {email address:{address:"${escapeForAppleScript(email)}"}, recipient type:recipient cc}`).join('\n')
        : '';
    const bccRecipients = bcc != null && bcc.length > 0
        ? bcc.map(email => `    make new recipient at newMessage with properties {email address:{address:"${escapeForAppleScript(email)}"}, recipient type:recipient bcc}`).join('\n')
        : '';
    const contentProperty = bodyType === 'html'
        ? `html content:"${escapedBody}"`
        : `plain text content:"${escapedBody}"`;
    const replyToStatement = replyTo != null
        ? `    set reply to of newMessage to "${escapeForAppleScript(replyTo)}"`
        : '';
    const attachmentStatements = attachments != null && attachments.length > 0
        ? attachments.map(att => `    make new attachment at newMessage with properties {file:(POSIX file "${escapeForAppleScript(att.path)}")}`).join('\n')
        : '';
    const inlineImageStatements = inlineImages != null && inlineImages.length > 0
        ? inlineImages.map((img, i) => `    set inlineAttach${i} to make new attachment at newMessage with properties {file:(POSIX file "${escapeForAppleScript(img.path)}")}\n` +
            `    try\n` +
            `      set content id of inlineAttach${i} to "${escapeForAppleScript(img.contentId)}"\n` +
            `    end try`).join('\n')
        : '';
    const accountStatement = accountId != null
        ? `    set sending account of newMessage to account id ${accountId}`
        : '';
    return `
tell application "Microsoft Outlook"
  try
    set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", ${contentProperty}}

${toRecipients}
${ccRecipients}
${bccRecipients}
${replyToStatement}
${accountStatement}
${attachmentStatements}
${inlineImageStatements}

    -- Capture ID before send (reference becomes stale after send)
    set msgId to "" & (id of newMessage)
    send newMessage

    set sentTime to current date
    set sentISO to sentTime as «class isot» as string

    -- Return success
    set output to "{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}" & msgId & "{{FIELD}}sentAt{{=}}" & sentISO
    return output
  on error errMsg
    -- Return failure
    set output to "{{RECORD}}success{{=}}false{{FIELD}}error{{=}}" & errMsg
    return output
  end try
end tell
`;
}
