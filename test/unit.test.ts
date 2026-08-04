/**
 * Unit tests for mcp-outlook-applescript.
 *
 * Validates generated AppleScript strings, parser behaviour, pagination helpers,
 * search optimizations, date filtering, deduplication, and security-critical
 * escaping — all WITHOUT requiring Microsoft Outlook to be running.
 *
 * Test groups:
 *   - Bug fixes (setMessageFlag, listTasks, listEventsByDateRange, flagStatus)
 *   - Pagination envelope (paginate helper)
 *   - AppleScript offset support (searchMessages, searchContacts, searchTasks, etc.)
 *   - Date filtering (listMessages, searchMessages, searchEvents)
 *   - Search optimizations (no matchedIds, phase 2 skip, no preview, phase 2 counter)
 *   - Deduplication (deduplicateEmailRows)
 *   - Timeout scaling (searchTimeoutMs)
 *   - Security (escapeForAppleScript, sendEmail template escaping)
 *   - Parsers (parseEmails, parseTasks, parseEvents, mutation results)
 */

import { describe, it, expect } from 'vitest';

// ---------------------------------------------------------------------------
// Imports under test
// ---------------------------------------------------------------------------

// scripts.ts exports pure template-generating functions (no Outlook dependency)
import {
  setMessageFlag,
  listMessages,
  getMessage,
  searchMessages,
  searchContacts,
  listTasks,
  searchTasks,
  getTask,
  searchEvents,
  searchNotes,
  listEvents,
  buildAppleScriptDateVar,
  sendEmail,
  setMessageCategories,
  createMailFolder,
} from '../src/applescript/scripts.js';

// executor.ts -- security-critical escaping
import { escapeForAppleScript } from '../src/applescript/executor.js';

// parser.ts exports pure parsing functions
import {
  parseEmails,
  parseEmail,
  parseTasks,
  parseEvents,
  parseSendEmailResult,
  parseDeleteEventResult,
} from '../src/applescript/parser.js';

// repository.ts -- exported helpers for testing
import { deduplicateEmailRows, searchTimeoutMs } from '../src/applescript/repository.js';
import type { EmailRow } from '../src/database/repository.js';

// date utilities
import { isoToAppleTimestamp, appleTimestampToIso } from '../src/utils/dates.js';

// pagination
import { paginate } from '../src/types/pagination.js';

// =============================================================================
// Bug 1: setMessageFlag -- valid AppleScript enum values
// =============================================================================

describe('Bug 1: setMessageFlag generates valid AppleScript enum values', () => {
  it('flagStatus 1 (flagged) produces "set todo flag of m to not completed"', () => {
    const script = setMessageFlag(123, 1);
    expect(script).toContain('set todo flag of m to not completed');
  });

  it('flagStatus 2 (completed) produces "set todo flag of m to completed"', () => {
    const script = setMessageFlag(123, 2);
    expect(script).toContain('set todo flag of m to completed');
  });

  it('flagStatus 0 (not flagged) produces "set todo flag of m to not flagged"', () => {
    const script = setMessageFlag(123, 0);
    expect(script).toContain('set todo flag of m to not flagged');
  });

  it('default/unknown flagStatus (e.g. 99) falls back to "not flagged"', () => {
    const script = setMessageFlag(123, 99);
    expect(script).toContain('set todo flag of m to not flagged');
  });

  it('uses the correct message id in the script', () => {
    const script = setMessageFlag(456, 1);
    expect(script).toContain('set m to message id 456');
  });

  it('does NOT contain old broken enum values like "flag marked"', () => {
    for (const status of [0, 1, 2]) {
      const script = setMessageFlag(100, status);
      expect(script).not.toContain('flag marked');
      expect(script).not.toContain('flag complete');
      expect(script).not.toContain('flag not flagged');
    }
  });

  it('does NOT contain bare numeric todo flag values', () => {
    // Verify we are using symbolic names, not bare numbers like "set todo flag of m to 1"
    for (const status of [0, 1, 2]) {
      const script = setMessageFlag(100, status);
      // The script should NOT have patterns like "to 0", "to 1", "to 2" as the flag value
      // (they should be symbolic: "not flagged", "not completed", "completed")
      const flagLine = script.split('\n').find(l => l.includes('set todo flag of m to'));
      expect(flagLine).toBeDefined();
      // After "to " should come a symbolic name, not a bare digit
      const afterTo = flagLine.trim().replace(/.*set todo flag of m to /, '');
      expect(afterTo).not.toMatch(/^\d+$/);
    }
  });

  it('returns a well-formed AppleScript tell block', () => {
    const script = setMessageFlag(123, 1);
    expect(script).toContain('tell application "Microsoft Outlook"');
    expect(script).toContain('end tell');
    expect(script).toContain('return "ok"');
  });
});

// =============================================================================
// Bug 2: listTasks -- correct whose clause and isCompleted read-back
// =============================================================================

describe('Bug 2: listTasks generates correct whose clause', () => {
  it('includeCompleted=false uses "whose todo flag is not completed"', () => {
    const script = listTasks(10, 0, false);
    expect(script).toContain('whose todo flag is not completed');
  });

  it('includeCompleted=false does NOT produce double "is" pattern', () => {
    const script = listTasks(10, 0, false);
    // The old bug: "whose is completed is false"
    expect(script).not.toContain('whose is completed is false');
    expect(script).not.toContain('whose is completed = false');
  });

  it('includeCompleted=true produces no whose filter', () => {
    const script = listTasks(10, 0, true);
    expect(script).not.toContain('whose');
  });

  it('uses "(todo flag of t is completed)" for isCompleted property read', () => {
    const scriptFiltered = listTasks(10, 0, false);
    const scriptAll = listTasks(10, 0, true);
    // Both should read completion status the same way
    expect(scriptFiltered).toContain('(todo flag of t is completed)');
    expect(scriptAll).toContain('(todo flag of t is completed)');
  });

  it('does NOT use old "is completed of t" syntax for property read', () => {
    const script = listTasks(10, 0, false);
    expect(script).not.toContain('is completed of t');
  });

  it('respects limit and offset parameters', () => {
    const script = listTasks(5, 3, false);
    // startIdx should be offset+1 = 4, endIdx should be limit+offset = 8
    expect(script).toContain('set startIdx to 4');
    expect(script).toContain('set endIdx to 8');
  });
});

describe('Bug 2: searchTasks uses correct isCompleted read-back', () => {
  it('contains "(todo flag of t is completed)" for property read', () => {
    const script = searchTasks('test query', 10, 0);
    expect(script).toContain('(todo flag of t is completed)');
  });

  it('does NOT contain "is completed of t"', () => {
    const script = searchTasks('test query', 10, 0);
    expect(script).not.toContain('is completed of t');
  });

  it('does NOT use a whose clause that filters by completion', () => {
    // searchTasks filters by name, not completion status
    const script = searchTasks('test query', 10, 0);
    expect(script).toContain('whose name contains');
    expect(script).not.toContain('whose todo flag');
    expect(script).not.toContain('whose is completed');
  });
});

describe('Bug 2: getTask uses correct isCompleted read-back', () => {
  it('contains "(todo flag of t is completed)" for property read', () => {
    const script = getTask(42);
    expect(script).toContain('(todo flag of t is completed)');
  });

  it('does NOT contain "is completed of t"', () => {
    const script = getTask(42);
    expect(script).not.toContain('is completed of t');
  });

  it('uses the correct task id', () => {
    const script = getTask(99);
    expect(script).toContain('set t to task id 99');
  });
});

// =============================================================================
// Bug 3: listEventsByDateRange -- now uses server-side AppleScript whose clause
// =============================================================================

describe('Bug 3: listEventsByDateRange uses server-side date filtering', () => {
  it('generates AppleScript with whose start time clause for date range', () => {
    // listEvents with date params generates a whose clause
    const script = listEvents(null, '2025-06-01T00:00:00Z', '2025-12-31T23:59:59Z', 10, 0);
    expect(script).toContain('whose start time');
    expect(script).toContain('afterDate');
    expect(script).toContain('beforeDate');
  });

  it('does NOT generate whose clause when no date filter', () => {
    const script = listEvents(null, null, null, 10, 0);
    expect(script).not.toContain('whose');
    expect(script).not.toContain('afterDate');
  });

  it('generates correct date variable construction', () => {
    const script = listEvents(null, '2025-06-15T14:30:00Z', null, 10, 0);
    expect(script).toContain('set year of afterDate to 2025');
    expect(script).toContain('set month of afterDate to 6');
    expect(script).toContain('set day of afterDate to 15');
    expect(script).toContain('set hours of afterDate to 14');
    expect(script).toContain('set minutes of afterDate to 30');
  });

  it('supports offset in date-filtered queries', () => {
    const script = listEvents(null, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z', 10, 5);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });
});

// =============================================================================
// Bug 4: flagStatus read-back in email scripts and parser
// =============================================================================

describe('Bug 4: listMessages includes flagStatus in generated AppleScript', () => {
  it('generates the flag reading block with set mFlag to "0"', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('set mFlag to "0"');
  });

  it('generates the todo flag conversion logic', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('set f to (todo flag of m) as string');
    expect(script).toContain('if f is "completed"');
    expect(script).toContain('set mFlag to "2"');
    expect(script).toContain('if f is "not flagged"');
    // Default else branch sets to "1" (flagged/marked)
    expect(script).toContain('set mFlag to "1"');
  });

  it('includes flagStatus in the output record', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });

  it('unreadOnly flag still works correctly', () => {
    const script = listMessages(125, 5, 0, true);
    expect(script).toContain('whose is read is false');
    // flagStatus should still be present
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });
});

describe('Bug 4: getMessage includes flagStatus', () => {
  it('generates the flag reading block', () => {
    const script = getMessage(123);
    expect(script).toContain('set mFlag to "0"');
    expect(script).toContain('set f to (todo flag of m) as string');
    expect(script).toContain('if f is "completed"');
    expect(script).toContain('set mFlag to "2"');
  });

  it('includes flagStatus in the output record', () => {
    const script = getMessage(123);
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });

  it('uses the correct message id', () => {
    const script = getMessage(789);
    expect(script).toContain('set m to message id 789');
  });
});

describe('Bug 4: searchMessages includes flagStatus in both phases', () => {
  it('includes flagStatus in Phase 1 (subject matches)', () => {
    const script = searchMessages('test', null, 5);
    // Phase 1 is the first block; it should have the flag logic
    const phase1Marker = '-- Phase 1: Subject matches';
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase1Section = script.slice(
      script.indexOf(phase1Marker),
      script.indexOf(phase2Marker)
    );
    expect(phase1Section).toContain('set mFlag to "0"');
    expect(phase1Section).toContain('flagStatus{{=}}" & mFlag');
  });

  it('includes flagStatus in Phase 2 (sender matches)', () => {
    const script = searchMessages('test', null, 5);
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase2Section = script.slice(script.indexOf(phase2Marker));
    expect(phase2Section).toContain('set mFlag to "0"');
    expect(phase2Section).toContain('flagStatus{{=}}" & mFlag');
  });

  it('includes todo flag conversion logic in both phases', () => {
    const script = searchMessages('test', null, 5);
    // Should have the todo flag logic appearing twice (once per phase)
    const matches = script.match(/set f to \(todo flag of m\) as string/g);
    expect(matches).not.toBeNull();
    expect(matches!.length).toBe(2);
  });

  it('works with a folder ID specified', () => {
    const script = searchMessages('query', 42, 10);
    expect(script).toContain('of mail folder id 42');
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });
});

// =============================================================================
// search_emails performance optimizations
// =============================================================================

describe('searchMessages optimization: no matchedIds dedup in AppleScript', () => {
  it('does NOT contain matchedIds list or "does not contain" check', () => {
    const script = searchMessages('test', null, 10, 0);
    expect(script).not.toContain('set matchedIds to');
    expect(script).not.toContain('does not contain mId');
  });

  it('does NOT collect skipped IDs for dedup', () => {
    const script = searchMessages('test', null, 10, 50);
    // The old code had: "repeat with i from 1 to preEnd" to collect skipped IDs
    expect(script).not.toContain('set end of matchedIds to id of item i of subjectMatches');
  });
});

describe('searchMessages optimization: phase 2 skip when phase 1 sufficient', () => {
  it('contains early return when phase 1 has enough results', () => {
    const script = searchMessages('test', null, 25, 0);
    expect(script).toContain('if phase1Total > (skipCount + maxResults)');
    expect(script).toContain('return output');
  });

  it('uses skipCount and maxResults in the phase 2 skip check', () => {
    const script = searchMessages('test', null, 10, 50);
    expect(script).toContain('set maxResults to 10');
    expect(script).toContain('set skipCount to 50');
    expect(script).toContain('if phase1Total > (skipCount + maxResults)');
  });
});

describe('searchMessages optimization: no preview in search results', () => {
  it('uses set mPreview to "" instead of PREVIEW_EXTRACT_BLOCK', () => {
    const script = searchMessages('test', null, 10, 0);
    // Should NOT contain the "plain text content of m" pattern used by PREVIEW_EXTRACT_BLOCK
    expect(script).not.toContain('plain text content of m');
    // Should contain the empty preview assignment
    expect(script).toContain('set mPreview to ""');
  });

  it('still includes preview field in output record', () => {
    const script = searchMessages('test', null, 10, 0);
    expect(script).toContain('preview{{=}}" & mPreview');
  });
});

describe('listMessages still includes preview (500 chars)', () => {
  it('includes PREVIEW_EXTRACT_BLOCK with plain text content of m', () => {
    const script = listMessages(100, 10, 0, false);
    expect(script).toContain('plain text content of m');
  });

  it('uses 500 char limit', () => {
    const script = listMessages(100, 10, 0, false);
    expect(script).toContain('> 500');
    expect(script).toContain('text 1 thru 500');
  });
});

describe('Bug 4: parseEmails extracts flagStatus from AppleScript output', () => {
  it('parses flagStatus as a number from delimited output', () => {
    const output = '{{RECORD}}id{{=}}1{{FIELD}}subject{{=}}Test Email{{FIELD}}flagStatus{{=}}2';
    const results = parseEmails(output);
    expect(results).toHaveLength(1);
    expect(results[0].flagStatus).toBe(2);
  });

  it('parses flagStatus 0 (not flagged)', () => {
    const output = '{{RECORD}}id{{=}}10{{FIELD}}subject{{=}}Hello{{FIELD}}flagStatus{{=}}0';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBe(0);
  });

  it('parses flagStatus 1 (flagged)', () => {
    const output = '{{RECORD}}id{{=}}20{{FIELD}}subject{{=}}Important{{FIELD}}flagStatus{{=}}1';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBe(1);
  });

  it('returns null when flagStatus is missing', () => {
    const output = '{{RECORD}}id{{=}}30{{FIELD}}subject{{=}}No Flag';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBeNull();
  });

  it('returns null when flagStatus is empty string', () => {
    const output = '{{RECORD}}id{{=}}40{{FIELD}}subject{{=}}Empty Flag{{FIELD}}flagStatus{{=}}';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBeNull();
  });

  it('parses multiple records each with their own flagStatus', () => {
    const output =
      '{{RECORD}}id{{=}}1{{FIELD}}subject{{=}}A{{FIELD}}flagStatus{{=}}0' +
      '{{RECORD}}id{{=}}2{{FIELD}}subject{{=}}B{{FIELD}}flagStatus{{=}}1' +
      '{{RECORD}}id{{=}}3{{FIELD}}subject{{=}}C{{FIELD}}flagStatus{{=}}2';
    const results = parseEmails(output);
    expect(results).toHaveLength(3);
    expect(results[0].flagStatus).toBe(0);
    expect(results[1].flagStatus).toBe(1);
    expect(results[2].flagStatus).toBe(2);
  });
});

describe('Bug 4: parseEmail (single) extracts flagStatus', () => {
  it('parses a single email with flagStatus', () => {
    const output = '{{RECORD}}id{{=}}99{{FIELD}}subject{{=}}Single{{FIELD}}flagStatus{{=}}1';
    const result = parseEmail(output);
    expect(result).not.toBeNull();
    expect(result.flagStatus).toBe(1);
  });

  it('returns null for empty input', () => {
    const result = parseEmail('');
    expect(result).toBeNull();
  });
});

// =============================================================================
// Response Envelope: paginate() helper
// =============================================================================

describe('paginate() helper', () => {
  it('returns hasMore=true when items exceed limit', () => {
    const items = [1, 2, 3, 4, 5, 6]; // 6 items, limit is 5
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(true);
    expect(result.count).toBe(5);
    expect(result.items).toHaveLength(5);
    expect(result.items).toEqual([1, 2, 3, 4, 5]);
  });

  it('returns hasMore=false when items are at or below limit', () => {
    const items = [1, 2, 3];
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(3);
    expect(result.items).toHaveLength(3);
  });

  it('returns hasMore=false when items equal limit exactly', () => {
    const items = [1, 2, 3, 4, 5];
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(5);
  });

  it('handles empty array', () => {
    const result = paginate([], 10);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(0);
    expect(result.items).toEqual([]);
  });
});

// =============================================================================
// Offset in search AppleScript templates
// =============================================================================

describe('Offset support in search AppleScript templates', () => {
  it('searchContacts generates startIdx/endIdx from offset', () => {
    const script = searchContacts('john', 10, 5);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });

  it('searchTasks generates startIdx/endIdx from offset', () => {
    const script = searchTasks('review', 10, 3);
    expect(script).toContain('set startIdx to 4');
    expect(script).toContain('set endIdx to 13');
  });

  it('searchNotes generates startIdx/endIdx from offset', () => {
    const script = searchNotes('meeting', 5, 10);
    expect(script).toContain('set startIdx to 11');
    expect(script).toContain('set endIdx to 15');
  });

  it('searchEvents generates startIdx/endIdx from offset', () => {
    const script = searchEvents('standup', 25, 0);
    expect(script).toContain('set startIdx to 1');
    expect(script).toContain('set endIdx to 25');
  });

  it('searchEvents with offset=25 generates startIdx 26', () => {
    const script = searchEvents('standup', 25, 25);
    expect(script).toContain('set startIdx to 26');
    expect(script).toContain('set endIdx to 50');
  });
});

// =============================================================================
// buildAppleScriptDateVar helper
// =============================================================================

describe('buildAppleScriptDateVar', () => {
  it('generates correct date components from ISO string', () => {
    const result = buildAppleScriptDateVar('afterDate', '2025-06-15T14:30:00Z');
    expect(result).toContain('set afterDate to current date');
    expect(result).toContain('set year of afterDate to 2025');
    expect(result).toContain('set month of afterDate to 6');
    expect(result).toContain('set day of afterDate to 15');
    expect(result).toContain('set hours of afterDate to 14');
    expect(result).toContain('set minutes of afterDate to 30');
    expect(result).toContain('set seconds of afterDate to 0');
  });

  it('handles midnight correctly', () => {
    const result = buildAppleScriptDateVar('beforeDate', '2025-01-01T00:00:00Z');
    expect(result).toContain('set year of beforeDate to 2025');
    expect(result).toContain('set month of beforeDate to 1');
    expect(result).toContain('set day of beforeDate to 1');
    expect(result).toContain('set hours of beforeDate to 0');
    expect(result).toContain('set minutes of beforeDate to 0');
  });

  it('sets day to 1 first to avoid month overflow', () => {
    const result = buildAppleScriptDateVar('d', '2025-12-31T23:59:00Z');
    const lines = result.split('\n');
    // "set day of d to 1" should come before "set year of d to ..."
    const dayTo1Idx = lines.findIndex(l => l.includes('set day of d to 1'));
    const yearIdx = lines.findIndex(l => l.includes('set year of d to'));
    expect(dayTo1Idx).toBeLessThan(yearIdx);
  });
});

// =============================================================================
// Date filtering in email AppleScript
// =============================================================================

describe('Date filtering in listMessages', () => {
  it('generates whose clause with time received >= afterDate', () => {
    const script = listMessages(100, 10, 0, false, '2025-06-01T00:00:00Z');
    expect(script).toContain('time received ≥ afterDate');
    expect(script).toContain('set year of afterDate to 2025');
    expect(script).toContain('set month of afterDate to 6');
  });

  it('generates whose clause with time received <= beforeDate', () => {
    const script = listMessages(100, 10, 0, false, undefined, '2025-12-31T23:59:59Z');
    expect(script).toContain('time received ≤ beforeDate');
    expect(script).toContain('set year of beforeDate to 2025');
    expect(script).toContain('set month of beforeDate to 12');
  });

  it('combines after + before + unreadOnly in whose clause', () => {
    const script = listMessages(100, 10, 0, true, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('is read is false and time received ≥ afterDate and time received ≤ beforeDate');
  });

  it('does NOT use whose clause when no date filter and not unread', () => {
    const script = listMessages(100, 10, 0, false);
    expect(script).not.toContain('whose');
    expect(script).toContain('messages 1 thru');
  });
});

describe('Date filtering in searchMessages', () => {
  it('extends whose clause with date filter in phase 1', () => {
    const script = searchMessages('test', null, 10, 0, '2025-06-01T00:00:00Z');
    // Phase 1 whose clause should include both subject and date
    expect(script).toContain('subject contains "test" and time received ≥ afterDate');
  });

  it('includes date check in phase 2 sender loop', () => {
    const script = searchMessages('test', null, 10, 0, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('mDateObj < afterDate');
    expect(script).toContain('mDateObj > beforeDate');
  });

  it('works without date filter (backward compatible)', () => {
    const script = searchMessages('test', null, 10);
    expect(script).toContain('whose subject contains "test"');
    expect(script).not.toContain('afterDate');
    expect(script).not.toContain('beforeDate');
  });

  it('uses native offset in AppleScript (startIdx/endIdx)', () => {
    const script = searchMessages('test', null, 10, 25);
    expect(script).toContain('set skipCount to 25');
    expect(script).toContain('set phase1Start to skipCount + 1');
    expect(script).toContain('set phase2Skip to skipCount - phase1Total');
  });
});

describe('Date filtering in searchEvents', () => {
  it('extends whose clause with after filter', () => {
    const script = searchEvents('standup', 10, 0, '2025-06-01T00:00:00Z');
    expect(script).toContain('subject contains "standup" and start time ≥ afterDate');
  });

  it('extends whose clause with both after and before', () => {
    const script = searchEvents('standup', 10, 0, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('subject contains "standup" and start time ≥ afterDate and start time ≤ beforeDate');
  });
});

// =============================================================================
// search_notes description fix
// =============================================================================

describe('search_notes description fix', () => {
  it('searchNotes AppleScript searches by name (not content)', () => {
    const script = searchNotes('meeting', 10, 0);
    expect(script).toContain('whose name contains');
    // Should NOT contain "content contains" or "plain text content contains"
    expect(script).not.toContain('content contains');
  });
});

// =============================================================================
// deduplicateEmailRows
// =============================================================================

describe('deduplicateEmailRows', () => {
  function makeRow(id: number, subject: string): EmailRow {
    return {
      id,
      folderId: 0,
      subject,
      sender: '',
      senderAddress: '',
      recipients: '',
      displayTo: '',
      toAddresses: '',
      ccAddresses: '',
      preview: '',
      isRead: 0,
      timeReceived: 0,
      timeSent: 0,
      hasAttachment: 0,
      size: 0,
      priority: 0,
      flagStatus: 0,
      categories: null,
      messageId: null,
      conversationId: null,
      dataFilePath: '',
    };
  }

  it('removes duplicate IDs, preserving first occurrence', () => {
    const rows = [
      makeRow(1, 'First'),
      makeRow(2, 'Second'),
      makeRow(1, 'First duplicate'),
      makeRow(3, 'Third'),
      makeRow(2, 'Second duplicate'),
    ];
    const result = deduplicateEmailRows(rows);
    expect(result).toHaveLength(3);
    expect(result[0].subject).toBe('First');
    expect(result[1].subject).toBe('Second');
    expect(result[2].subject).toBe('Third');
  });

  it('returns all rows when no duplicates', () => {
    const rows = [makeRow(1, 'A'), makeRow(2, 'B'), makeRow(3, 'C')];
    const result = deduplicateEmailRows(rows);
    expect(result).toHaveLength(3);
  });

  it('handles empty array', () => {
    const result = deduplicateEmailRows([]);
    expect(result).toHaveLength(0);
  });

  it('handles single row', () => {
    const result = deduplicateEmailRows([makeRow(42, 'Solo')]);
    expect(result).toHaveLength(1);
    expect(result[0].id).toBe(42);
  });

  it('preserves order of first occurrences', () => {
    const rows = [
      makeRow(3, 'Third-first'),
      makeRow(1, 'First-first'),
      makeRow(3, 'Third-dupe'),
      makeRow(2, 'Second-first'),
      makeRow(1, 'First-dupe'),
    ];
    const result = deduplicateEmailRows(rows);
    expect(result.map(r => r.id)).toEqual([3, 1, 2]);
  });
});

// =============================================================================
// searchTimeoutMs
// =============================================================================

describe('searchTimeoutMs', () => {
  it('returns 90s base for offset=0', () => {
    expect(searchTimeoutMs(0)).toBe(90000);
  });

  it('returns 100s for offset=25 (page 2)', () => {
    expect(searchTimeoutMs(25)).toBe(100000);
  });

  it('returns 110s for offset=50 (page 3)', () => {
    expect(searchTimeoutMs(50)).toBe(110000);
  });

  it('returns 120s for offset=75 (page 4)', () => {
    expect(searchTimeoutMs(75)).toBe(120000);
  });

  it('caps at 150s for high offsets', () => {
    expect(searchTimeoutMs(150)).toBe(150000);
    expect(searchTimeoutMs(300)).toBe(150000);
    expect(searchTimeoutMs(1000)).toBe(150000);
  });

  it('handles fractional pages (offset not multiple of 25)', () => {
    // offset=24 → Math.floor(24/25) = 0 → 90s
    expect(searchTimeoutMs(24)).toBe(90000);
    // offset=26 → Math.floor(26/25) = 1 → 100s
    expect(searchTimeoutMs(26)).toBe(100000);
  });
});

// =============================================================================
// Phase 2 independent counter (over-counting fix)
// =============================================================================

describe('searchMessages phase 2 uses independent counter', () => {
  it('uses phase2Count instead of resultCount in phase 2 loop exit', () => {
    const script = searchMessages('test', null, 10, 0);
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase2Section = script.slice(script.indexOf(phase2Marker));
    // Phase 2 should use its own counter
    expect(phase2Section).toContain('set phase2Count to 0');
    expect(phase2Section).toContain('if phase2Count ≥ maxResults then exit repeat');
    expect(phase2Section).toContain('set phase2Count to phase2Count + 1');
  });

  it('does NOT use shared resultCount in phase 2 loop body', () => {
    const script = searchMessages('test', null, 10, 0);
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase2Section = script.slice(script.indexOf(phase2Marker));
    // Phase 2 should NOT increment the shared resultCount
    expect(phase2Section).not.toContain('set resultCount to resultCount + 1');
  });

  it('still uses resultCount gate to skip phase 2 when phase 1 filled the page', () => {
    const script = searchMessages('test', null, 10, 0);
    expect(script).toContain('if resultCount < maxResults then');
  });

  it('clamps phase2Skip to 0 when negative', () => {
    const script = searchMessages('test', null, 10, 25);
    expect(script).toContain('if phase2Skip < 0 then set phase2Skip to 0');
  });
});

// =============================================================================
// Security: escapeForAppleScript
// =============================================================================

describe('escapeForAppleScript', () => {
  it('returns plain text unchanged', () => {
    expect(escapeForAppleScript('hello world')).toBe('hello world');
  });

  it('escapes double quotes', () => {
    expect(escapeForAppleScript('say "hello"')).toBe('say \\"hello\\"');
  });

  it('escapes backslashes', () => {
    expect(escapeForAppleScript('path\\to\\file')).toBe('path\\\\to\\\\file');
  });

  it('converts \\n to AppleScript linefeed concatenation', () => {
    // \n becomes " & linefeed & " — designed to be embedded inside a quoted string
    expect(escapeForAppleScript('line1\nline2')).toBe('line1" & linefeed & "line2');
  });

  it('converts \\r\\n to AppleScript return concatenation', () => {
    expect(escapeForAppleScript('line1\r\nline2')).toBe('line1" & return & "line2');
  });

  it('converts \\r to AppleScript return concatenation', () => {
    expect(escapeForAppleScript('line1\rline2')).toBe('line1" & return & "line2');
  });

  it('handles combined quotes and backslashes', () => {
    const input = 'He said "C:\\Users"';
    const result = escapeForAppleScript(input);
    expect(result).toContain('\\"C:\\\\Users\\"');
  });

  it('handles empty string', () => {
    expect(escapeForAppleScript('')).toBe('');
  });
});

// =============================================================================
// Security: sendEmail template escaping
// =============================================================================

describe('sendEmail template escaping', () => {
  it('escapes subject with quotes', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'He said "hello"',
      body: 'body',
      bodyType: 'plain',
    });
    // escapeForAppleScript turns " into \" — so the AppleScript string contains \"
    expect(script).toContain('He said \\"hello\\"');
  });

  it('escapes all recipient email addresses', () => {
    const script = sendEmail({
      to: ['to@example.com'],
      cc: ['cc@example.com'],
      bcc: ['bcc@example.com'],
      subject: 'test',
      body: 'body',
      bodyType: 'plain',
    });
    expect(script).toContain('to@example.com');
    expect(script).toContain('cc@example.com');
    expect(script).toContain('bcc@example.com');
  });

  it('escapes attachment paths', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'test',
      body: 'body',
      bodyType: 'plain',
      attachments: [{ path: '/Users/test/file with spaces.pdf' }],
    });
    expect(script).toContain('POSIX file');
    expect(script).toContain('file with spaces.pdf');
  });

  it('escapes replyTo address', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'test',
      body: 'body',
      bodyType: 'plain',
      replyTo: 'reply@example.com',
    });
    expect(script).toContain('reply to');
    expect(script).toContain('reply@example.com');
  });

  it('uses content property for html bodyType', () => {
    // Outlook's AppleScript dictionary names the HTML body property "content"
    // (see `message` class: `content` -- "The HTML content of a message").
    // "html content" is not a valid property and causes a -2741 syntax error
    // ("Expected ':' but found property") when compiled by osascript.
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'test',
      body: '<h1>Hello</h1>',
      bodyType: 'html',
    });
    expect(script).toContain('content:');
    expect(script).not.toContain('html content:');
  });

  it('uses plain text content property for plain bodyType', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'test',
      body: 'Hello',
      bodyType: 'plain',
    });
    expect(script).toContain('plain text content:');
  });
});

// =============================================================================
// setMessageCategories template
// =============================================================================

describe('setMessageCategories', () => {
  it('converts array to AppleScript list', () => {
    const script = setMessageCategories(42, ['Work', 'Urgent']);
    expect(script).toContain('{\"Work\", \"Urgent\"}');
    expect(script).toContain('message id 42');
  });

  it('escapes category names with quotes', () => {
    const script = setMessageCategories(1, ['Category "A"']);
    expect(script).toContain('\\"A\\"');
  });
});

// =============================================================================
// createMailFolder template
// =============================================================================

describe('createMailFolder', () => {
  it('creates at root when no parent specified', () => {
    const script = createMailFolder('New Folder', undefined);
    expect(script).toContain('make new mail folder with properties');
    expect(script).not.toContain('mail folder id');
  });

  it('creates inside parent when parentFolderId specified', () => {
    const script = createMailFolder('Subfolder', 42);
    expect(script).toContain('mail folder id 42');
    expect(script).toContain('make new mail folder at parentFolder');
  });
});

// =============================================================================
// Parser: parseTasks
// =============================================================================

describe('parseTasks', () => {
  it('parses task with isCompleted true', () => {
    const output = '{{RECORD}}id{{=}}1{{FIELD}}name{{=}}Buy milk{{FIELD}}dueDate{{=}}2025-06-15T00:00:00{{FIELD}}isCompleted{{=}}true';
    const tasks = parseTasks(output);
    expect(tasks).toHaveLength(1);
    expect(tasks[0].id).toBe(1);
    expect(tasks[0].name).toBe('Buy milk');
    expect(tasks[0].isCompleted).toBe(true);
  });

  it('parses task with isCompleted false', () => {
    const output = '{{RECORD}}id{{=}}2{{FIELD}}name{{=}}Do laundry{{FIELD}}isCompleted{{=}}false';
    const tasks = parseTasks(output);
    expect(tasks[0].isCompleted).toBe(false);
  });

  it('parses multiple tasks', () => {
    const output = '{{RECORD}}id{{=}}1{{FIELD}}name{{=}}Task A{{RECORD}}id{{=}}2{{FIELD}}name{{=}}Task B';
    const tasks = parseTasks(output);
    expect(tasks).toHaveLength(2);
  });

  it('returns empty array for empty input', () => {
    expect(parseTasks('')).toHaveLength(0);
  });
});

// =============================================================================
// Parser: parseEvents
// =============================================================================

describe('parseEvents', () => {
  it('parses event with all fields', () => {
    const output = '{{RECORD}}id{{=}}100{{FIELD}}subject{{=}}Team Standup{{FIELD}}startTime{{=}}2025-06-15T10:00:00{{FIELD}}endTime{{=}}2025-06-15T10:30:00{{FIELD}}location{{=}}Room 5{{FIELD}}isAllDay{{=}}false{{FIELD}}isRecurring{{=}}true';
    const events = parseEvents(output);
    expect(events).toHaveLength(1);
    expect(events[0].id).toBe(100);
    expect(events[0].subject).toBe('Team Standup');
    expect(events[0].isRecurring).toBe(true);
  });

  it('returns empty array for empty input', () => {
    expect(parseEvents('')).toHaveLength(0);
  });
});

// =============================================================================
// Parser: mutation results
// =============================================================================

describe('parseSendEmailResult', () => {
  it('parses success result', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}messageId{{=}}12345{{FIELD}}sentAt{{=}}2025-06-15T10:00:00';
    const result = parseSendEmailResult(output);
    expect(result).not.toBeNull();
    expect(result!.success).toBe(true);
    expect(result!.messageId).toBe('12345');
  });

  it('parses failure result', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Network error';
    const result = parseSendEmailResult(output);
    expect(result).not.toBeNull();
    expect(result!.success).toBe(false);
    expect(result!.error).toBe('Network error');
  });
});

describe('parseDeleteEventResult', () => {
  it('parses success result', () => {
    const output = '{{RECORD}}success{{=}}true{{FIELD}}eventId{{=}}999';
    const result = parseDeleteEventResult(output);
    expect(result).not.toBeNull();
    expect(result!.success).toBe(true);
  });

  it('parses failure result', () => {
    const output = '{{RECORD}}success{{=}}false{{FIELD}}error{{=}}Event not found';
    const result = parseDeleteEventResult(output);
    expect(result!.success).toBe(false);
    expect(result!.error).toBe('Event not found');
  });
});

// =============================================================================
// Date utilities
// =============================================================================

describe('isoToAppleTimestamp / appleTimestampToIso', () => {
  it('round-trips a known date', () => {
    const iso = '2025-06-15T14:30:00.000Z';
    const apple = isoToAppleTimestamp(iso);
    expect(apple).toBeGreaterThan(0);
    const backToIso = appleTimestampToIso(apple);
    expect(backToIso).toBe(iso);
  });

  it('returns null for null/undefined/empty ISO input', () => {
    expect(isoToAppleTimestamp(null as unknown as string)).toBeNull();
    expect(isoToAppleTimestamp(undefined as unknown as string)).toBeNull();
    expect(isoToAppleTimestamp('')).toBeNull();
  });

  it('returns Apple epoch (2001-01-01) for timestamp 0', () => {
    // 0 in Apple epoch = 2001-01-01T00:00:00.000Z, not null
    expect(appleTimestampToIso(0)).toBe('2001-01-01T00:00:00.000Z');
  });

  it('returns null for null/undefined apple timestamp', () => {
    expect(appleTimestampToIso(null)).toBeNull();
    expect(appleTimestampToIso(undefined)).toBeNull();
  });
});
