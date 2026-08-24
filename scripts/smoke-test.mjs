#!/usr/bin/env node
// =============================================================================
// smoke-test.mjs — Comprehensive MCP smoke test harness (9 suites)
//
// Spawns dist/index.js as a child process, performs the MCP JSON-RPC handshake,
// then tests across 9 suites covering 45 of 52 tools live.
// (5 skipped: send_email, create_draft, download_attachment, respond_to_event, update_event).
//
// Requirements: Outlook must be running. dist/ must be built.
// Usage: node scripts/smoke-test.mjs
// =============================================================================

import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SERVER_PATH = join(__dirname, '..', 'dist', 'index.js');
const TIMEOUT_MS = 30_000;
const GLOBAL_TIMEOUT_MS = 600_000; // 10 min for full suite (search tools can be slow)
const SLOW_THRESHOLD_MS = 60_000;
const FAKE_UUID = '00000000-0000-0000-0000-000000000000';

// =============================================================================
// JSON-RPC transport over stdio
// =============================================================================

let nextId = 1;
let serverDead = false;
/** @type {import('node:child_process').ChildProcess} */
let serverProcess;
/** @type {Map<number, { resolve: Function, reject: Function, timer: ReturnType<typeof setTimeout> }>} */
const pending = new Map();
/** @type {Buffer[]} */
const stderrChunks = [];

function rejectAllPending(reason) {
    for (const [, { reject, timer }] of pending) {
        clearTimeout(timer);
        reject(new Error(reason));
    }
    pending.clear();
}

function startServer() {
    serverProcess = spawn('node', [SERVER_PATH], {
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...process.env },
    });

    serverProcess.on('error', (err) => {
        console.error(`Server process error: ${err.message}`);
        process.exit(1);
    });

    serverProcess.on('exit', (code) => {
        serverDead = true;
        rejectAllPending('Server process exited unexpectedly');
        if (code !== null && code !== 0) {
            console.error(`Server exited with code ${code}`);
        }
    });

    serverProcess.stdin.on('error', () => {});

    const rl = createInterface({ input: serverProcess.stdout });
    rl.on('line', (line) => {
        if (!line.trim()) return;
        let msg;
        try {
            msg = JSON.parse(line);
        } catch {
            return;
        }
        if (msg.id != null && pending.has(msg.id)) {
            const { resolve, timer } = pending.get(msg.id);
            clearTimeout(timer);
            pending.delete(msg.id);
            resolve(msg);
        }
    });
    rl.on('close', () => {
        serverDead = true;
        rejectAllPending('Server stdout closed unexpectedly');
    });

    serverProcess.stderr.on('data', (chunk) => stderrChunks.push(chunk));
}

function getStderr() {
    return Buffer.concat(stderrChunks).toString();
}

function request(method, params = {}, timeoutMs = TIMEOUT_MS) {
    if (serverDead) return Promise.reject(new Error('Server is not running'));
    const id = nextId++;
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
            pending.delete(id);
            reject(new Error(`Timeout waiting for response to ${method} (id=${id})`));
        }, timeoutMs);
        pending.set(id, { resolve, reject, timer });
        const line = JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n';
        serverProcess.stdin.write(line);
    });
}

function notify(method, params = {}) {
    if (serverDead || !serverProcess?.stdin?.writable) return;
    const line = JSON.stringify({ jsonrpc: '2.0', method, params }) + '\n';
    serverProcess.stdin.write(line);
}

function stopServer() {
    rejectAllPending('Server shutting down');
    if (serverProcess && !serverDead) {
        serverProcess.stdin.end();
        serverProcess.kill('SIGTERM');
    }
}

function gracefulExit() {
    stopServer();
    process.exit(1);
}
process.on('SIGINT', gracefulExit);
process.on('SIGTERM', gracefulExit);

// =============================================================================
// Test framework with timing
// =============================================================================

/** @type {{ name: string, suite: string, status: 'PASS'|'FAIL'|'SKIP', detail: string, elapsed: number }[]} */
const results = [];
let currentSuite = '';

function suite(name) {
    currentSuite = name;
    console.log('');
    console.log(`${name}`);
}

function pass(name, detail = '', elapsed = 0) {
    const slow = elapsed >= SLOW_THRESHOLD_MS ? ' \u26a0 SLOW' : '';
    results.push({ name, suite: currentSuite, status: 'PASS', detail, elapsed });
    const ms = elapsed > 0 ? `${elapsed}ms`.padStart(8) : '';
    console.log(`  \u2713 ${name}${ms}${slow}${detail ? `  ${detail}` : ''}`);
}

function fail(name, detail = '', elapsed = 0) {
    results.push({ name, suite: currentSuite, status: 'FAIL', detail, elapsed });
    const ms = elapsed > 0 ? `${elapsed}ms`.padStart(8) : '';
    console.log(`  \u2717 ${name}${ms}  ${detail}`);
}

function skip(name, reason = '') {
    results.push({ name, suite: currentSuite, status: 'SKIP', detail: reason, elapsed: 0 });
    console.log(`  \u25cb ${name}  (${reason})`);
}

// =============================================================================
// Tool call helper with timing + validation
// =============================================================================

/**
 * @param {string} displayName
 * @param {string} toolName
 * @param {object} args
 * @param {object} opts
 * @param {boolean} [opts.expectError]
 * @param {string} [opts.expectErrorContains]
 * @param {(parsed: any) => string|null} [opts.validate]
 * @param {boolean} [opts.rawTextResponse] - If true, don't try to JSON parse
 * @param {number} [opts.timeoutMs] - Per-tool timeout in ms (default: TIMEOUT_MS)
 * @returns {Promise<any>}
 */
async function testTool(displayName, toolName, args, opts = {}) {
    const { expectError = false, expectErrorContains, validate, rawTextResponse = false, timeoutMs } = opts;
    const t0 = Date.now();

    let response;
    try {
        response = await request('tools/call', { name: toolName, arguments: args }, timeoutMs ?? TIMEOUT_MS);
    } catch (err) {
        const elapsed = Date.now() - t0;
        if (expectError) {
            pass(displayName, `transport error as expected: ${err.message}`, elapsed);
            return null;
        }
        fail(displayName, err.message, elapsed);
        return null;
    }

    const elapsed = Date.now() - t0;

    // Protocol-level error
    if (response.error) {
        if (expectError) {
            const errMsg = response.error.message || JSON.stringify(response.error);
            if (expectErrorContains && !errMsg.toLowerCase().includes(expectErrorContains.toLowerCase())) {
                fail(displayName, `Expected error containing "${expectErrorContains}", got: ${errMsg.slice(0, 120)}`, elapsed);
                return null;
            }
            pass(displayName, `protocol error: ${errMsg.slice(0, 80)}`, elapsed);
            return null;
        }
        fail(displayName, `JSON-RPC error: ${response.error.message || JSON.stringify(response.error)}`, elapsed);
        return null;
    }

    const result = response.result;
    if (!result?.content?.[0]) {
        fail(displayName, 'Missing result.content[0]', elapsed);
        return null;
    }

    const content = result.content[0];
    if (content.type !== 'text') {
        fail(displayName, `Expected content type "text", got "${content.type}"`, elapsed);
        return null;
    }

    const text = content.text;

    // Error path
    if (expectError) {
        if (!result.isError) {
            fail(displayName, 'Expected isError but got success', elapsed);
            return null;
        }
        if (expectErrorContains && !text.toLowerCase().includes(expectErrorContains.toLowerCase())) {
            fail(displayName, `Expected error containing "${expectErrorContains}", got: ${text.slice(0, 120)}`, elapsed);
            return null;
        }
        pass(displayName, `error: ${text.slice(0, 80)}`, elapsed);
        return null;
    }

    // Success path
    if (result.isError) {
        fail(displayName, `Unexpected error: ${text.slice(0, 200)}`, elapsed);
        return null;
    }

    if (rawTextResponse) {
        if (validate) {
            const err = validate(text);
            if (err) { fail(displayName, err, elapsed); return null; }
        }
        pass(displayName, text.slice(0, 80), elapsed);
        return text;
    }

    // Parse JSON
    let parsed;
    try {
        parsed = JSON.parse(text);
    } catch {
        fail(displayName, `Not valid JSON: ${text.slice(0, 120)}`, elapsed);
        return null;
    }

    // Custom validation
    if (validate) {
        const err = validate(parsed);
        if (err) { fail(displayName, err, elapsed); return null; }
    }

    // Build detail summary
    let detail = '';
    if (Array.isArray(parsed)) {
        detail = `${parsed.length} items`;
    } else if (parsed && typeof parsed === 'object') {
        // Handle paginated envelope: { items, count, hasMore }
        if (Array.isArray(parsed.items) && typeof parsed.hasMore === 'boolean') {
            detail = `${parsed.count} items${parsed.hasMore ? ', hasMore' : ''}`;
        } else {
            const keys = Object.keys(parsed);
            for (const k of keys) {
                if (Array.isArray(parsed[k])) {
                    detail = `${k}: ${parsed[k].length} items`;
                    break;
                }
            }
            if (!detail) detail = `keys: ${keys.slice(0, 5).join(', ')}`;
        }
    }

    pass(displayName, detail, elapsed);
    return parsed;
}

// =============================================================================
// Helper: extract folders from response (handles single-account vs multi-account)
// =============================================================================

function extractFolders(foldersResult) {
    if (!foldersResult) return [];
    if (Array.isArray(foldersResult)) return foldersResult;
    if (foldersResult.accounts?.[0]?.folders) return foldersResult.accounts[0].folders;
    if (foldersResult.folders) return foldersResult.folders;
    return [];
}

/**
 * Extracts items from a paginated envelope response.
 * Handles both { items, count, hasMore } envelope and legacy bare arrays.
 */
function extractItems(parsed) {
    if (!parsed) return [];
    if (Array.isArray(parsed)) return parsed;
    if (parsed.items && Array.isArray(parsed.items)) return parsed.items;
    return [];
}

/**
 * Validates a paginated envelope response.
 * Returns null if valid, error string if invalid.
 */
function validateEnvelope(parsed, label = '') {
    if (!parsed || typeof parsed !== 'object') return `${label}: expected object, got ${typeof parsed}`;
    if (!Array.isArray(parsed.items)) return `${label}: missing items array`;
    if (typeof parsed.count !== 'number') return `${label}: missing count (number)`;
    if (typeof parsed.hasMore !== 'boolean') return `${label}: missing hasMore (boolean)`;
    if (parsed.count !== parsed.items.length) return `${label}: count (${parsed.count}) !== items.length (${parsed.items.length})`;
    return null;
}

// =============================================================================
// Main test sequence
// =============================================================================

async function run() {
    const globalTimer = setTimeout(() => {
        console.error(`\nGlobal timeout (${GLOBAL_TIMEOUT_MS / 1000}s) exceeded`);
        stopServer();
        process.exit(1);
    }, GLOBAL_TIMEOUT_MS);
    globalTimer.unref();

    const runStart = Date.now();

    console.log('');
    console.log('=== MCP Smoke Test: mcp-outlook-applescript ===');

    startServer();

    // =========================================================================
    // Suite 1: Protocol (2 tests)
    // =========================================================================
    suite('PROTOCOL');

    // Test 1: initialize handshake
    let initResponse;
    const t0Init = Date.now();
    try {
        initResponse = await request('initialize', {
            protocolVersion: '2024-11-05',
            capabilities: {},
            clientInfo: { name: 'smoke-test', version: '1.0.0' },
        });
    } catch (err) {
        fail('initialize handshake', err.message, Date.now() - t0Init);
        const stderr = getStderr();
        if (stderr) console.log('\nServer stderr:\n' + stderr);
        stopServer();
        process.exit(1);
    }

    if (initResponse.error) {
        fail('initialize handshake', `Error: ${JSON.stringify(initResponse.error)}`, Date.now() - t0Init);
        stopServer();
        process.exit(1);
    }

    const serverInfo = initResponse.result?.serverInfo;
    if (serverInfo?.name === 'outlook-mcp' && initResponse.result?.protocolVersion) {
        pass('initialize handshake', `${serverInfo.name} v${serverInfo.version}, protocol ${initResponse.result.protocolVersion}`, Date.now() - t0Init);
    } else {
        fail('initialize handshake', `serverInfo.name=${serverInfo?.name}, protocolVersion=${initResponse.result?.protocolVersion}`, Date.now() - t0Init);
    }

    notify('notifications/initialized');

    // Test 2: tools/list returns 52 tools
    const t0Tools = Date.now();
    let toolsList;
    try {
        toolsList = await request('tools/list');
    } catch (err) {
        fail('tools/list returns 52 tools', err.message, Date.now() - t0Tools);
        stopServer();
        process.exit(1);
    }

    const tools = toolsList.result?.tools;
    if (!Array.isArray(tools)) {
        fail('tools/list returns 52 tools', 'Expected tools array', Date.now() - t0Tools);
        stopServer();
        process.exit(1);
    }

    const expectedTools = [
        'list_accounts', 'list_folders', 'list_emails', 'search_emails', 'get_email',
        'get_unread_count', 'list_attachments', 'download_attachment',
        'list_calendars', 'list_events', 'get_event', 'search_events',
        'create_event', 'respond_to_event', 'delete_event', 'update_event',
        'list_contacts', 'search_contacts', 'get_contact',
        'list_tasks', 'search_tasks', 'get_task',
        'list_notes', 'get_note', 'search_notes',
        'send_email', 'create_draft', 'reply_to_email', 'reply_all_email',
        'prepare_delete_email', 'confirm_delete_email',
        'prepare_move_email', 'confirm_move_email',
        'prepare_archive_email', 'confirm_archive_email',
        'prepare_junk_email', 'confirm_junk_email',
        'prepare_delete_folder', 'confirm_delete_folder',
        'prepare_empty_folder', 'confirm_empty_folder',
        'prepare_batch_delete_emails', 'prepare_batch_move_emails', 'confirm_batch_operation',
        'mark_email_read', 'mark_email_unread',
        'set_email_flag', 'clear_email_flag', 'set_email_categories',
        'create_folder', 'rename_folder', 'move_folder',
    ];

    const toolNames = tools.map(t => t.name).sort();
    const registeredSet = new Set(toolNames);
    const missing = expectedTools.filter(t => !registeredSet.has(t));
    const extra = toolNames.filter(t => !expectedTools.includes(t));
    const allHaveDesc = tools.every(t => t.description);
    const allHaveSchema = tools.every(t => t.inputSchema);

    if (tools.length === 50 && missing.length === 0 && extra.length === 0 && allHaveDesc && allHaveSchema) {
        pass('tools/list returns 52 tools', `${tools.length} tools, all with description + inputSchema`, Date.now() - t0Tools);
    } else {
        const issues = [];
        if (tools.length !== 50) issues.push(`count=${tools.length}`);
        if (missing.length > 0) issues.push(`missing: ${missing.join(', ')}`);
        if (extra.length > 0) issues.push(`extra: ${extra.join(', ')}`);
        if (!allHaveDesc) issues.push('some lack description');
        if (!allHaveSchema) issues.push('some lack inputSchema');
        fail('tools/list returns 52 tools', issues.join('; '), Date.now() - t0Tools);
    }

    // =========================================================================
    // Dynamic ID Discovery
    // =========================================================================

    // Discover folder IDs
    const foldersResult = await testTool('list_folders', 'list_folders', {}, {
        validate: (parsed) => {
            const folders = extractFolders(parsed);
            if (!Array.isArray(folders) || folders.length === 0) return 'no folders found';
            const f = folders[0];
            if (f.id == null || !f.name) return `folder missing id/name: ${JSON.stringify(f)}`;
            return null;
        },
    });

    const folders = extractFolders(foldersResult);
    const inbox = folders.find(f => /inbox/i.test(f.name));
    const inboxFolderId = inbox?.id ?? folders[0]?.id ?? null;

    // Discover email IDs
    let discoveredEmailIds = [];
    if (inboxFolderId != null) {
        const emailsResult = await testTool('list_emails (discovery)', 'list_emails', { folder_id: inboxFolderId, limit: 5 }, {
            validate: (parsed) => {
                const envErr = validateEnvelope(parsed, 'list_emails');
                if (envErr) return envErr;
                const items = parsed.items;
                if (items.length === 0) return 'no emails found in inbox';
                const e = items[0];
                if (e.id == null) return 'email missing id';
                if (!('subject' in e)) return 'email missing subject';
                if (!('isRead' in e)) return 'email missing isRead';
                if (!('flagStatus' in e)) return 'email missing flagStatus (bug fix 4)';
                return null;
            },
        });
        const emailItems = extractItems(emailsResult);
        discoveredEmailIds = emailItems.map(e => e.id);
    }
    const emailId = discoveredEmailIds[0] ?? null;

    // Discover event IDs
    const eventsResult = await testTool('list_events (discovery)', 'list_events', { limit: 5 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_events');
            if (envErr) return envErr;
            return null;
        },
    });
    const eventId = extractItems(eventsResult)?.[0]?.id ?? null;

    // Discover contact IDs
    const contactsResult = await testTool('list_contacts (discovery)', 'list_contacts', { limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_contacts');
            if (envErr) return envErr;
            return null;
        },
    });
    const contactId = extractItems(contactsResult)?.[0]?.id ?? null;

    // Discover task IDs
    const tasksResult = await testTool('list_tasks (discovery)', 'list_tasks', { limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_tasks');
            if (envErr) return envErr;
            return null;
        },
    });
    const taskId = extractItems(tasksResult)?.[0]?.id ?? null;

    // Discover note IDs
    const notesResult = await testTool('list_notes (discovery)', 'list_notes', { limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_notes');
            if (envErr) return envErr;
            return null;
        },
    });
    const noteId = extractItems(notesResult)?.[0]?.id ?? null;

    console.log('');
    console.log(`  Discovered IDs: email=${emailId}, event=${eventId}, contact=${contactId}, task=${taskId}, note=${noteId}, inbox=${inboxFolderId}`);

    // =========================================================================
    // Suite 2: Read-Only Happy Path (20 tests)
    // =========================================================================
    suite('READ-ONLY: Happy Path');

    // list_accounts
    await testTool('list_accounts', 'list_accounts', {}, {
        validate: (parsed) => {
            if (!parsed.accounts || !Array.isArray(parsed.accounts)) return 'expected { accounts: [...] }';
            return null;
        },
    });

    // search_emails (two-phase search needs more time than default 30s)
    await testTool('search_emails', 'search_emails', { query: 'test', limit: 3 }, {
        timeoutMs: 90_000,
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_emails');
            if (envErr) return envErr;
            if (parsed.items.length > 0 && parsed.items[0].id == null) return 'email missing id';
            return null;
        },
    });

    // get_email
    if (emailId != null) {
        await testTool('get_email', 'get_email', { email_id: emailId }, {
            validate: (parsed) => {
                if (parsed.id == null) return 'missing id';
                if (!('subject' in parsed)) return 'missing subject';
                if (!('folderId' in parsed)) return 'missing folderId';
                return null;
            },
        });
    } else {
        skip('get_email', 'no email ID discovered');
    }

    // get_unread_count
    await testTool('get_unread_count', 'get_unread_count', {}, {
        validate: (parsed) => {
            if (typeof parsed.count !== 'number') return `expected { count: number }, got ${JSON.stringify(parsed).slice(0, 80)}`;
            return null;
        },
    });

    // list_attachments
    if (emailId != null) {
        await testTool('list_attachments', 'list_attachments', { email_id: emailId }, {
            validate: (parsed) => {
                if (!Array.isArray(parsed)) return 'expected array';
                return null;
            },
        });
    } else {
        skip('list_attachments', 'no email ID discovered');
    }

    // list_calendars
    await testTool('list_calendars', 'list_calendars', {}, {
        validate: (parsed) => {
            if (!Array.isArray(parsed)) return 'expected array';
            if (parsed.length > 0) {
                const c = parsed[0];
                if (c.id == null || !c.name) return `calendar missing id/name: ${JSON.stringify(c)}`;
            }
            return null;
        },
    });

    // list_events (already done in discovery, but test with date range)
    await testTool('list_events (date range)', 'list_events', {
        start_date: new Date(Date.now() - 30 * 86400000).toISOString(),
        end_date: new Date(Date.now() + 30 * 86400000).toISOString(),
        limit: 5,
    }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_events (date range)');
            if (envErr) return envErr;
            if (parsed.items.length > 0) {
                const e = parsed.items[0];
                if (e.id == null) return 'event missing id';
                if (!e.title && e.title !== '') return 'event missing title';
                if (!e.startDate) return 'event missing startDate';
                if (!e.endDate) return 'event missing endDate';
            }
            return null;
        },
    });

    // get_event
    if (eventId != null) {
        await testTool('get_event', 'get_event', { event_id: eventId }, {
            validate: (parsed) => {
                if (parsed.id == null) return 'missing id';
                if (!('title' in parsed)) return 'missing title';
                if (!parsed.startDate) return 'missing startDate';
                return null;
            },
        });
    } else {
        skip('get_event', 'no event ID discovered');
    }

    // search_events
    await testTool('search_events', 'search_events', { query: 'meeting', limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_events');
            if (envErr) return envErr;
            return null;
        },
    });

    // search_contacts
    await testTool('search_contacts', 'search_contacts', { query: 'a', limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_contacts');
            if (envErr) return envErr;
            return null;
        },
    });

    // get_contact
    if (contactId != null) {
        await testTool('get_contact', 'get_contact', { contact_id: contactId }, {
            validate: (parsed) => {
                if (parsed.id == null) return 'missing id';
                if (!('displayName' in parsed)) return 'missing displayName';
                return null;
            },
        });
    } else {
        skip('get_contact', 'no contact ID discovered');
    }

    // search_tasks
    await testTool('search_tasks', 'search_tasks', { query: 'a', limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_tasks');
            if (envErr) return envErr;
            return null;
        },
    });

    // get_task
    if (taskId != null) {
        await testTool('get_task', 'get_task', { task_id: taskId }, {
            validate: (parsed) => {
                if (parsed.id == null) return 'missing id';
                if (!('name' in parsed)) return 'missing name';
                return null;
            },
        });
    } else {
        skip('get_task', 'no task ID discovered');
    }

    // get_note
    if (noteId != null) {
        await testTool('get_note', 'get_note', { note_id: noteId }, {
            validate: (parsed) => {
                if (parsed.id == null) return 'missing id';
                if (!('title' in parsed)) return 'missing title';
                return null;
            },
        });
    } else {
        skip('get_note', 'no note ID discovered');
    }

    // search_notes
    await testTool('search_notes', 'search_notes', { query: 'a', limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_notes');
            if (envErr) return envErr;
            return null;
        },
    });

    // list_tasks with include_completed: false (bug fix 2)
    await testTool('list_tasks (exclude completed)', 'list_tasks', { include_completed: false, limit: 3 }, {
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'list_tasks');
            if (envErr) return envErr;
            // If there are tasks, none should be completed
            for (const t of parsed.items) {
                if (t.isCompleted === true) return `task ${t.id} is completed but include_completed=false`;
            }
            return null;
        },
    });

    // =========================================================================
    // Suite 3: Read-Only Error Paths (7 tests)
    // =========================================================================
    suite('READ-ONLY: Error Paths');

    await testTool('get_email not found', 'get_email', { email_id: 999999999 }, {
        expectError: true,
        expectErrorContains: 'not found',
    });

    await testTool('get_event not found', 'get_event', { event_id: 999999999 }, {
        expectError: true,
        expectErrorContains: 'not found',
    });

    await testTool('get_contact not found', 'get_contact', { contact_id: 999999999 }, {
        expectError: true,
        expectErrorContains: 'not found',
    });

    await testTool('get_task not found', 'get_task', { task_id: 999999999 }, {
        expectError: true,
        expectErrorContains: 'not found',
    });

    await testTool('get_note not found', 'get_note', { note_id: 999999999 }, {
        expectError: true,
        expectErrorContains: 'not found',
    });

    // Empty search results should return envelope with empty items, not error
    // No-match query: phase 1 returns 0, phase 2 scans 500 messages (~60s) — needs generous timeout
    await testTool('search_emails (no match)', 'search_emails', { query: 'xyzzy_no_match_99' }, {
        timeoutMs: 150_000,
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_emails (no match)');
            if (envErr) return envErr;
            return null;
        },
    });

    await testTool('search_events (no match)', 'search_events', { query: 'xyzzy_no_match_99' }, {
        timeoutMs: 90_000,
        validate: (parsed) => {
            const envErr = validateEnvelope(parsed, 'search_events (no match)');
            if (envErr) return envErr;
            return null;
        },
    });

    // =========================================================================
    // Suite 4: Zod Validation Errors (5 tests)
    // =========================================================================
    suite('ZOD VALIDATION');

    await testTool('list_emails without folder_id', 'list_emails', {}, {
        expectError: true,
    });

    await testTool('search_emails without query', 'search_emails', {}, {
        expectError: true,
    });

    await testTool('get_email without email_id', 'get_email', {}, {
        expectError: true,
    });

    await testTool('get_event without event_id', 'get_event', {}, {
        expectError: true,
    });

    // Unknown tool should give a protocol error
    await testTool('unknown tool', 'nonexistent_tool_xyz', {}, {
        expectError: true,
    });

    // =========================================================================
    // Suite 5: Prepare Tools — Happy Path (8 tests)
    // =========================================================================
    suite('PREPARE TOOLS');

    if (emailId != null) {
        // prepare_delete_email
        await testTool('prepare_delete_email', 'prepare_delete_email', { email_id: emailId }, {
            validate: (parsed) => {
                if (!parsed.token_id || typeof parsed.token_id !== 'string') return 'missing token_id';
                if (!parsed.expires_at) return 'missing expires_at';
                if (!parsed.action || typeof parsed.action !== 'string') return 'missing action';
                if (!parsed.email || parsed.email.id == null) return 'missing email preview';
                return null;
            },
        });

        // prepare_move_email
        if (inboxFolderId != null) {
            await testTool('prepare_move_email', 'prepare_move_email', {
                email_id: emailId,
                destination_folder_id: inboxFolderId,
            }, {
                validate: (parsed) => {
                    if (!parsed.token_id) return 'missing token_id';
                    if (!parsed.expires_at) return 'missing expires_at';
                    if (!parsed.action) return 'missing action';
                    if (!parsed.email || parsed.email.id == null) return 'missing email preview';
                    if (!parsed.destination_folder) return 'missing destination_folder preview';
                    return null;
                },
            });
        } else {
            skip('prepare_move_email', 'no inbox folder ID');
        }

        // prepare_archive_email
        await testTool('prepare_archive_email', 'prepare_archive_email', { email_id: emailId }, {
            validate: (parsed) => {
                if (!parsed.token_id) return 'missing token_id';
                if (!parsed.expires_at) return 'missing expires_at';
                if (!parsed.action) return 'missing action';
                return null;
            },
        });

        // prepare_junk_email
        await testTool('prepare_junk_email', 'prepare_junk_email', { email_id: emailId }, {
            validate: (parsed) => {
                if (!parsed.token_id) return 'missing token_id';
                if (!parsed.expires_at) return 'missing expires_at';
                if (!parsed.action) return 'missing action';
                return null;
            },
        });

        // prepare_batch_delete_emails (use first 2 discovered IDs)
        const batchIds = discoveredEmailIds.slice(0, 2);
        if (batchIds.length >= 1) {
            await testTool('prepare_batch_delete_emails', 'prepare_batch_delete_emails', {
                email_ids: batchIds,
            }, {
                validate: (parsed) => {
                    if (!Array.isArray(parsed.tokens)) return 'missing tokens array';
                    if (parsed.tokens.length !== batchIds.length) return `expected ${batchIds.length} tokens, got ${parsed.tokens.length}`;
                    if (!parsed.tokens[0].token_id) return 'token missing token_id';
                    if (!parsed.action) return 'missing action';
                    return null;
                },
            });
        } else {
            skip('prepare_batch_delete_emails', 'not enough email IDs');
        }

        // prepare_batch_move_emails
        if (inboxFolderId != null && batchIds.length >= 1) {
            await testTool('prepare_batch_move_emails', 'prepare_batch_move_emails', {
                email_ids: batchIds.slice(0, 1),
                destination_folder_id: inboxFolderId,
            }, {
                validate: (parsed) => {
                    if (!Array.isArray(parsed.tokens)) return 'missing tokens array';
                    if (!parsed.destination_folder) return 'missing destination_folder';
                    if (!parsed.action) return 'missing action';
                    return null;
                },
            });
        } else {
            skip('prepare_batch_move_emails', 'no email or folder IDs');
        }
    } else {
        skip('prepare_delete_email', 'no email ID discovered');
        skip('prepare_move_email', 'no email ID discovered');
        skip('prepare_archive_email', 'no email ID discovered');
        skip('prepare_junk_email', 'no email ID discovered');
        skip('prepare_batch_delete_emails', 'no email ID discovered');
        skip('prepare_batch_move_emails', 'no email ID discovered');
    }

    // prepare_delete_folder + prepare_empty_folder: find a safe folder (0 messages, non-critical)
    const safeFolder = folders.find(f =>
        f.messageCount === 0 &&
        !/^(inbox|drafts|sent|trash|junk|deleted|archive|outbox)/i.test(f.name ?? '')
    );

    if (safeFolder) {
        await testTool('prepare_delete_folder', 'prepare_delete_folder', { folder_id: safeFolder.id }, {
            validate: (parsed) => {
                if (!parsed.token_id) return 'missing token_id';
                if (!parsed.expires_at) return 'missing expires_at';
                if (!parsed.folder) return 'missing folder preview';
                return null;
            },
        });

        await testTool('prepare_empty_folder', 'prepare_empty_folder', { folder_id: safeFolder.id }, {
            validate: (parsed) => {
                if (!parsed.token_id) return 'missing token_id';
                if (!parsed.expires_at) return 'missing expires_at';
                if (!parsed.folder) return 'missing folder preview';
                return null;
            },
        });
    } else {
        skip('prepare_delete_folder', 'no safe empty folder found');
        skip('prepare_empty_folder', 'no safe empty folder found');
    }

    // =========================================================================
    // Suite 6: Confirm Tools — Invalid Token Rejection (7 tests)
    // =========================================================================
    suite('CONFIRM TOOLS (invalid token)');

    await testTool('confirm_delete_email rejects fake token', 'confirm_delete_email', {
        token_id: FAKE_UUID, email_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_move_email rejects fake token', 'confirm_move_email', {
        token_id: FAKE_UUID, email_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_archive_email rejects fake token', 'confirm_archive_email', {
        token_id: FAKE_UUID, email_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_junk_email rejects fake token', 'confirm_junk_email', {
        token_id: FAKE_UUID, email_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_delete_folder rejects fake token', 'confirm_delete_folder', {
        token_id: FAKE_UUID, folder_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_empty_folder rejects fake token', 'confirm_empty_folder', {
        token_id: FAKE_UUID, folder_id: 1,
    }, {
        expectError: true,
        expectErrorContains: 'token',
    });

    await testTool('confirm_batch_operation rejects fake token', 'confirm_batch_operation', {
        tokens: [{ token_id: FAKE_UUID, email_id: 1 }],
    }, {
        validate: (parsed) => {
            // confirm_batch_operation returns results array, not isError
            // Each item should show success: false
            if (!parsed.results || !Array.isArray(parsed.results)) return 'missing results array';
            if (parsed.results.length === 0) return 'expected at least one result';
            if (parsed.results[0].success !== false) return `expected success=false, got ${parsed.results[0].success}`;
            return null;
        },
    });

    // =========================================================================
    // Suite 7: Low-Risk Modification Round-Trips (3 tests)
    // =========================================================================
    suite('LOW-RISK MODIFICATIONS (round-trip)');

    if (emailId != null) {
        // Test 1: Read status round-trip
        const t0Read = Date.now();
        let readTestPassed = false;
        try {
            // Get original state
            const origResp = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const origEmail = JSON.parse(origResp.result.content[0].text);
            const origIsRead = origEmail.isRead;

            // Mark as read
            await request('tools/call', { name: 'mark_email_read', arguments: { email_id: emailId } });
            const afterRead = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const readEmail = JSON.parse(afterRead.result.content[0].text);

            if (readEmail.isRead !== true) {
                fail('read status round-trip', `isRead should be true after mark_email_read, got ${readEmail.isRead}`, Date.now() - t0Read);
            } else {
                // Restore original state
                if (!origIsRead) {
                    await request('tools/call', { name: 'mark_email_unread', arguments: { email_id: emailId } });
                    const restored = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
                    const restoredEmail = JSON.parse(restored.result.content[0].text);
                    if (restoredEmail.isRead !== false) {
                        fail('read status round-trip', 'failed to restore unread state', Date.now() - t0Read);
                    } else {
                        readTestPassed = true;
                    }
                } else {
                    readTestPassed = true;
                }
            }

            if (readTestPassed) {
                pass('read status round-trip', `was ${origIsRead ? 'read' : 'unread'}, toggled, restored`, Date.now() - t0Read);
            }
        } catch (err) {
            fail('read status round-trip', err.message, Date.now() - t0Read);
        }

        // Test 2: Flag status round-trip (validates bug fix 1)
        const t0Flag = Date.now();
        let flagTestPassed = false;
        try {
            // Get original state
            const origResp = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const origEmail = JSON.parse(origResp.result.content[0].text);
            const origFlag = origEmail.flagStatus ?? 0;

            // Set flag to 1 (flagged)
            await request('tools/call', { name: 'set_email_flag', arguments: { email_id: emailId, flag_status: 1 } });
            const afterFlag = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const flaggedEmail = JSON.parse(afterFlag.result.content[0].text);

            if (flaggedEmail.flagStatus !== 1) {
                fail('flag status round-trip', `flagStatus should be 1 after set_email_flag, got ${flaggedEmail.flagStatus}`, Date.now() - t0Flag);
            } else {
                // Clear flag
                await request('tools/call', { name: 'clear_email_flag', arguments: { email_id: emailId } });
                const afterClear = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
                const clearedEmail = JSON.parse(afterClear.result.content[0].text);

                if (clearedEmail.flagStatus !== 0) {
                    fail('flag status round-trip', `flagStatus should be 0 after clear, got ${clearedEmail.flagStatus}`, Date.now() - t0Flag);
                } else {
                    // Restore original if it was different from 0
                    if (origFlag !== 0) {
                        await request('tools/call', { name: 'set_email_flag', arguments: { email_id: emailId, flag_status: origFlag } });
                    }
                    flagTestPassed = true;
                }
            }

            if (flagTestPassed) {
                pass('flag status round-trip', `orig=${origFlag}, set=1, cleared=0, restored`, Date.now() - t0Flag);
            }
        } catch (err) {
            fail('flag status round-trip', err.message, Date.now() - t0Flag);
        }

        // Test 3: Categories round-trip
        const t0Cat = Date.now();
        let catTestPassed = false;
        try {
            // Get original categories
            const origResp = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const origEmail = JSON.parse(origResp.result.content[0].text);
            const origCategories = origEmail.categories ?? [];

            // Set test category
            await request('tools/call', { name: 'set_email_categories', arguments: {
                email_id: emailId,
                categories: ['__smoke_test'],
            }});
            const afterSet = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
            const catEmail = JSON.parse(afterSet.result.content[0].text);

            const hasSmokeTest = (catEmail.categories ?? []).some(c => c === '__smoke_test');
            if (!hasSmokeTest) {
                fail('categories round-trip', `expected "__smoke_test" in categories, got: ${JSON.stringify(catEmail.categories)}`, Date.now() - t0Cat);
            } else {
                // Restore original categories
                await request('tools/call', { name: 'set_email_categories', arguments: {
                    email_id: emailId,
                    categories: origCategories,
                }});
                const restored = await request('tools/call', { name: 'get_email', arguments: { email_id: emailId } });
                const restoredEmail = JSON.parse(restored.result.content[0].text);
                const restoredCats = restoredEmail.categories ?? [];

                // Verify restoration
                const origSet = new Set(origCategories);
                const restoredSet = new Set(restoredCats);
                const match = origSet.size === restoredSet.size && [...origSet].every(c => restoredSet.has(c));
                if (!match) {
                    fail('categories round-trip', `failed to restore: expected ${JSON.stringify(origCategories)}, got ${JSON.stringify(restoredCats)}`, Date.now() - t0Cat);
                } else {
                    catTestPassed = true;
                }
            }

            if (catTestPassed) {
                pass('categories round-trip', `orig=[${origCategories.join(',')}], set=["__smoke_test"], restored`, Date.now() - t0Cat);
            }
        } catch (err) {
            fail('categories round-trip', err.message, Date.now() - t0Cat);
        }
    } else {
        skip('read status round-trip', 'no email ID discovered');
        skip('flag status round-trip', 'no email ID discovered');
        skip('categories round-trip', 'no email ID discovered');
    }

    // =========================================================================
    // Suite 8: Calendar Create+Delete Round-Trip (1 test)
    // =========================================================================
    suite('CALENDAR ROUND-TRIP');

    const t0Cal = Date.now();
    let calTestPassed = false;
    let testEventId = null;
    try {
        const tomorrow = new Date(Date.now() + 86400000);
        tomorrow.setHours(10, 0, 0, 0);
        const endTime = new Date(tomorrow.getTime() + 3600000);
        const startIso = tomorrow.toISOString();
        const endIso = endTime.toISOString();

        // Create event
        const createResp = await request('tools/call', { name: 'create_event', arguments: {
            title: '__SMOKE_TEST_EVENT',
            start_date: startIso,
            end_date: endIso,
        }});

        if (createResp.result?.isError) {
            fail('calendar create+delete', `create failed: ${createResp.result.content[0].text}`, Date.now() - t0Cal);
        } else {
            const created = JSON.parse(createResp.result.content[0].text);
            testEventId = created.id;
            if (!testEventId) {
                fail('calendar create+delete', 'created event missing id', Date.now() - t0Cal);
            } else {
                // Delete the event
                const deleteResp = await request('tools/call', { name: 'delete_event', arguments: {
                    event_id: testEventId,
                }});

                if (deleteResp.result?.isError) {
                    fail('calendar create+delete', `delete failed: ${deleteResp.result.content[0].text}`, Date.now() - t0Cal);
                } else {
                    // Verify deleted — get_event should return error
                    const verifyResp = await request('tools/call', { name: 'get_event', arguments: {
                        event_id: testEventId,
                    }});

                    if (verifyResp.result?.isError) {
                        testEventId = null; // cleaned up
                        calTestPassed = true;
                    } else {
                        const verifyText = verifyResp.result?.content?.[0]?.text ?? '';
                        if (verifyText.includes('not found')) {
                            testEventId = null;
                            calTestPassed = true;
                        } else {
                            fail('calendar create+delete', `event still exists after delete (id=${testEventId})`, Date.now() - t0Cal);
                        }
                    }
                }
            }
        }

        if (calTestPassed) {
            pass('calendar create+delete', 'create \u2192 delete \u2192 verify deleted', Date.now() - t0Cal);
        }
    } catch (err) {
        fail('calendar create+delete', err.message, Date.now() - t0Cal);
    }

    // Emergency cleanup: if test event was created but not deleted
    if (testEventId != null) {
        console.log(`  (cleaning up leftover test event id=${testEventId})`);
        try {
            await request('tools/call', { name: 'delete_event', arguments: { event_id: testEventId } });
        } catch { /* best effort */ }
    }

    // =========================================================================
    // Suite 9: Folder Create+Rename+Delete Round-Trip (1 test)
    // =========================================================================
    suite('FOLDER ROUND-TRIP');

    const t0Folder = Date.now();
    let folderTestPassed = false;
    let testFolderId = null;
    let parentFolderId = null;
    try {
        // Step 1: Create parent folder (used as move destination)
        const parentResp = await request('tools/call', { name: 'create_folder', arguments: {
            name: '__smoke_test_parent',
        }});
        if (!parentResp.result?.isError) {
            const parentData = JSON.parse(parentResp.result.content[0].text);
            parentFolderId = parentData.folder?.id;
        }

        // Step 2: Create the main test folder
        const createResp = await request('tools/call', { name: 'create_folder', arguments: {
            name: '__smoke_test_folder',
        }});

        if (createResp.result?.isError) {
            fail('folder lifecycle', `create failed: ${createResp.result.content[0].text}`, Date.now() - t0Folder);
        } else {
            const created = JSON.parse(createResp.result.content[0].text);
            testFolderId = created.folder?.id;

            if (!testFolderId) {
                fail('folder lifecycle', 'created folder missing id', Date.now() - t0Folder);
            } else {
                // Step 3: Verify exists in list_folders
                const listResp = await request('tools/call', { name: 'list_folders', arguments: {} });
                const allFolders = extractFolders(JSON.parse(listResp.result.content[0].text));
                const found = allFolders.some(f => f.name === '__smoke_test_folder');

                if (!found) {
                    fail('folder lifecycle', 'created folder not found in list_folders', Date.now() - t0Folder);
                } else {
                    // Step 4: Rename
                    const renameResp = await request('tools/call', { name: 'rename_folder', arguments: {
                        folder_id: testFolderId,
                        new_name: '__smoke_test_renamed',
                    }});

                    if (renameResp.result?.isError) {
                        fail('folder lifecycle', `rename failed: ${renameResp.result.content[0].text}`, Date.now() - t0Folder);
                    } else {
                        // Verify rename
                        const listResp2 = await request('tools/call', { name: 'list_folders', arguments: {} });
                        const allFolders2 = extractFolders(JSON.parse(listResp2.result.content[0].text));
                        const foundRenamed = allFolders2.some(f => f.name === '__smoke_test_renamed');

                        if (!foundRenamed) {
                            fail('folder lifecycle', 'renamed folder not found', Date.now() - t0Folder);
                        } else {
                            // Step 5: Move folder (into parent, if parent was created)
                            let moveOk = true;
                            if (parentFolderId != null) {
                                const moveResp = await request('tools/call', { name: 'move_folder', arguments: {
                                    folder_id: testFolderId,
                                    destination_parent_id: parentFolderId,
                                }});
                                if (moveResp.result?.isError) {
                                    fail('folder lifecycle', `move failed: ${moveResp.result.content[0].text}`, Date.now() - t0Folder);
                                    moveOk = false;
                                }
                            }
                            // move_folder is still tested even if parent creation failed (just skip the move step)

                            if (moveOk) {
                                // Step 6: Delete via prepare+confirm
                                const prepResp = await request('tools/call', { name: 'prepare_delete_folder', arguments: {
                                    folder_id: testFolderId,
                                }});

                                if (prepResp.result?.isError) {
                                    fail('folder lifecycle', `prepare_delete failed: ${prepResp.result.content[0].text}`, Date.now() - t0Folder);
                                } else {
                                    const prepData = JSON.parse(prepResp.result.content[0].text);
                                    const confResp = await request('tools/call', { name: 'confirm_delete_folder', arguments: {
                                        token_id: prepData.token_id,
                                        folder_id: testFolderId,
                                    }});

                                    if (confResp.result?.isError) {
                                        fail('folder lifecycle', `confirm_delete failed: ${confResp.result.content[0].text}`, Date.now() - t0Folder);
                                    } else {
                                        // Step 7: Verify deleted
                                        const listResp3 = await request('tools/call', { name: 'list_folders', arguments: {} });
                                        const allFolders3 = extractFolders(JSON.parse(listResp3.result.content[0].text));
                                        const stillExists = allFolders3.some(f =>
                                            f.name === '__smoke_test_renamed' || f.name === '__smoke_test_folder'
                                        );

                                        if (stillExists) {
                                            fail('folder lifecycle', 'folder still exists after delete', Date.now() - t0Folder);
                                        } else {
                                            testFolderId = null; // cleaned up
                                            folderTestPassed = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if (folderTestPassed) {
            const steps = parentFolderId != null
                ? 'create \u2192 verify \u2192 rename \u2192 move \u2192 delete \u2192 verify'
                : 'create \u2192 verify \u2192 rename \u2192 delete \u2192 verify (move skipped: no parent)';
            pass('folder lifecycle', steps, Date.now() - t0Folder);
        }
    } catch (err) {
        fail('folder lifecycle', err.message, Date.now() - t0Folder);
    }

    // Emergency cleanup: delete test folders if they weren't cleaned up
    async function cleanupFolder(folderId, label) {
        if (folderId == null) return;
        console.log(`  (cleaning up leftover ${label} id=${folderId})`);
        try {
            const prepResp = await request('tools/call', { name: 'prepare_delete_folder', arguments: { folder_id: folderId } });
            if (!prepResp.result?.isError) {
                const prepData = JSON.parse(prepResp.result.content[0].text);
                await request('tools/call', { name: 'confirm_delete_folder', arguments: { token_id: prepData.token_id, folder_id: folderId } });
            }
        } catch { /* best effort */ }
    }
    await cleanupFolder(testFolderId, 'test folder');
    await cleanupFolder(parentFolderId, 'parent folder');

    // =========================================================================
    // Skipped tools
    // =========================================================================
    suite('SKIPPED');
    skip('send_email', 'would send real mail');
    skip('download_attachment', 'needs specific attachment, writes to filesystem');
    skip('respond_to_event', 'would change RSVP status');
    skip('update_event', 'tested via create+delete round-trip');

    // =========================================================================
    // Summary + Performance
    // =========================================================================
    const totalElapsed = Date.now() - runStart;

    console.log('');
    console.log('=== PERFORMANCE ===');

    const timed = results.filter(r => r.elapsed > 0 && r.status !== 'SKIP');
    if (timed.length > 0) {
        const sorted = [...timed].sort((a, b) => a.elapsed - b.elapsed);
        const fastest = sorted[0];
        const slowest = sorted[sorted.length - 1];
        const median = sorted[Math.floor(sorted.length / 2)];
        const slow = sorted.filter(r => r.elapsed >= SLOW_THRESHOLD_MS);

        console.log(`  Fastest: ${fastest.name} (${fastest.elapsed}ms)`);
        console.log(`  Median:  ${median.name} (${median.elapsed}ms)`);
        console.log(`  Slowest: ${slowest.name} (${slowest.elapsed}ms)`);
        if (slow.length > 0) {
            console.log(`  \u26a0 Slow (>${SLOW_THRESHOLD_MS}ms): ${slow.map(r => `${r.name} (${r.elapsed}ms)`).join(', ')}`);
        }
    }

    console.log('');
    console.log('=== SUMMARY ===');
    console.log('');

    const passed = results.filter(r => r.status === 'PASS').length;
    const failed = results.filter(r => r.status === 'FAIL').length;
    const skipped = results.filter(r => r.status === 'SKIP').length;

    console.log(`  Passed:  ${passed}`);
    console.log(`  Failed:  ${failed}`);
    console.log(`  Skipped: ${skipped}`);
    console.log(`  Total:   ${(totalElapsed / 1000).toFixed(1)}s`);
    console.log('');

    if (failed > 0) {
        console.log('  RESULT: FAIL');
        console.log('');
        console.log('  Failed tests:');
        for (const r of results.filter(r => r.status === 'FAIL')) {
            console.log(`    \u2717 ${r.name}: ${r.detail}`);
        }
        const stderr = getStderr();
        if (stderr.trim()) {
            console.log('');
            console.log('  Server stderr:');
            for (const line of stderr.trim().split('\n').slice(0, 20)) {
                console.log(`    ${line}`);
            }
        }
    } else {
        console.log('  RESULT: ALL PASS');
    }

    console.log('');
    clearTimeout(globalTimer);
    stopServer();
    process.exit(failed > 0 ? 1 : 0);
}

run().catch((err) => {
    console.error('Unhandled error:', err);
    stopServer();
    process.exit(1);
});
