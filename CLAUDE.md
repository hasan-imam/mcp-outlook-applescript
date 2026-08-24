# mcp-outlook-applescript

MCP server for Microsoft Outlook on Mac. 52 tools for mail, calendar, contacts, tasks, and notes via AppleScript. TypeScript with strict mode, Zod validation, and security-hardened input escaping.

## Commands

| Command | Purpose |
|---------|---------|
| `npm run build` | Compile TypeScript to `dist/` and set shebang |
| `npm run typecheck` | Type-check without emitting |
| `npm test` | Run 131 unit tests (vitest, no Outlook needed) |
| `bash scripts/audit.sh` | 28 static quality checks (build, security, package, functional) |
| `node scripts/doc-quality.mjs` | Score all 49 tool descriptions against 6-criterion rubric |
| `node scripts/smoke-test.mjs` | Live integration tests (requires Outlook running) |

## Testing

| When | Run | What it checks |
|------|-----|----------------|
| Any code change | `npm run build && npm test && bash scripts/audit.sh` | Types, 131 unit tests, 28 static checks (security, package, bug-fix preservation) |
| Before release, or after changing AppleScript/repository/parser | `node scripts/smoke-test.mjs` (Outlook must be running) | Live integration: 45/52 tools via JSON-RPC against real Outlook |
| After changing tool descriptions or Zod schemas | Give an MCP client a task without naming tools — it must discover and call them from descriptions | Tool discoverability, schema usability, response format |

## Architecture

```
src/                                 # 34 TypeScript files
├── index.ts                         # MCP server entry point, tool dispatch
├── applescript/                     # AppleScript backend
│   ├── executor.ts                  #   osascript runner, escapeForAppleScript()
│   ├── scripts.ts                   #   AppleScript templates (pure functions)
│   ├── parser.ts                    #   delimiter-based output parser
│   ├── repository.ts                #   IRepository implementation
│   ├── account-repository.ts        #   account enumeration
│   ├── account-scripts.ts           #   account-related AppleScript
│   ├── calendar-manager.ts          #   event update/delete/respond
│   ├── calendar-writer.ts           #   event creation
│   ├── content-readers.ts           #   single-item content readers
│   ├── mail-sender.ts               #   send email with attachments
│   └── index.ts                     #   barrel exports
├── approval/                        # Two-step prepare/confirm for destructive ops
│   ├── token-manager.ts             #   single-use token store (5-min TTL)
│   ├── hash.ts                      #   content hashing for change detection
│   └── types.ts                     #   ApprovalToken, OperationType, TargetType
├── database/
│   └── repository.ts                # IRepository + IWriteableRepository interfaces
├── parsers/
│   └── html-stripper.ts             # HTML to plain text
├── tools/                           # Zod input schemas + tool logic
│   ├── mail.ts                      #   mail read tools + send_email
│   ├── calendar.ts                  #   calendar tools
│   ├── contacts.ts                  #   contact tools
│   ├── tasks.ts                     #   task tools
│   ├── notes.ts                     #   note tools
│   ├── mailbox-organization.ts      #   destructive ops + low-risk mutations
│   └── index.ts                     #   barrel exports
├── types/                           # Domain types + pagination envelope
│   └── pagination.ts                #   PaginatedResult<T>, paginate()
└── utils/
    ├── dates.ts                     # Apple epoch (2001) <-> ISO conversion
    └── errors.ts                    # ErrorCode enum, OutlookMcpError hierarchy
```

### Data flow

1. `src/index.ts` registers 52 tools on the MCP server with Zod schemas
2. Each handler calls a tool class (`MailTools`, `CalendarTools`, etc.) which calls `IRepository`
3. `IRepository` (in `applescript/repository.ts`) calls `scripts.ts` to build an AppleScript string
4. `executor.ts` runs it via `osascript` (stdin, not shell) and returns raw output
5. `parser.ts` splits output on `{{RECORD}}` / `{{FIELD}}` / `{{=}}` delimiters into typed rows
6. Tool class maps rows to domain types and wraps in `paginate()` for list/search results

### Lazy initialization

The server does not connect to Outlook on startup. The first tool call triggers `initializeAppleScriptBackend()`, which checks `isOutlookRunning()` and creates all repositories and tool instances.

## Key conventions

### Zod strict schemas

Every tool uses `z.strictObject()` with `.describe()` on every field. No `as any`, no `args as {...}` casts. Schemas live in `src/tools/<category>.ts` and are re-exported through `src/tools/index.ts`.

### Input escaping

All user-provided strings interpolated into AppleScript templates must pass through `escapeForAppleScript()` in `executor.ts`. This escapes backslashes, double quotes, and converts newlines to AppleScript `linefeed`/`return` concatenation. The audit script verifies this for `sendEmail()` specifically (addresses, replyTo, attachment paths, contentId).

### Pagination envelope

All list/search tools return `{ items, count, hasMore }`. The `paginate()` helper in `src/types/pagination.ts` implements this by over-fetching `limit + 1` rows and trimming. Callers page by incrementing `offset` by `limit`.

### Approval system for destructive operations

Destructive operations (delete, move, archive, junk, empty folder) use a two-step flow:

1. `prepare_*` tool: reads the target, hashes it, generates a single-use token (5-minute TTL), returns a preview
2. `confirm_*` tool: validates the token, checks the hash matches current state, executes the operation, consumes the token

Token state is in-memory only (`ApprovalTokenManager`). Expired tokens are garbage-collected when the store exceeds 100 entries.

### Error handling

All errors extend `OutlookMcpError` (in `utils/errors.ts`) which carries a structured `ErrorCode`. The `handle()` wrapper in `index.ts` catches all errors and returns `{ text, isError: true }` with `CODE: message` format. Error categories: `OUTLOOK_NOT_RUNNING`, `APPLESCRIPT_PERMISSION_DENIED`, `APPLESCRIPT_TIMEOUT`, `APPLESCRIPT_ERROR`, `NOT_FOUND`, `VALIDATION_ERROR`, `APPROVAL_*`, `ATTACHMENT_*`, `MAIL_SEND_ERROR`.

### Delimiter-based output protocol

AppleScript outputs are not JSON. Scripts emit records separated by `{{RECORD}}`, fields by `{{FIELD}}`, key-value pairs by `{{=}}`, and null sentinels as `{{NULL}}`. The parser in `parser.ts` splits on these and converts to typed row interfaces.

## How to add a new tool

1. **`src/applescript/scripts.ts`** -- Add a template function. Use `escapeForAppleScript()` on all user inputs. Use shared output constants (`DELIMITERS`, `FLAG_STATUS_BLOCK`, etc.).
2. **`src/database/repository.ts`** -- Add the method to `IRepository` (reads) or `IWriteableRepository` (mutations).
3. **`src/applescript/repository.ts`** -- Implement: call the script, execute via `executeAppleScriptOrThrow`, parse, map to row types.
4. **`src/tools/<category>.ts`** -- Add a `z.strictObject()` schema with `.describe()` on every field. Add the method to the tool class. Use `paginate()` for list/search results.
5. **`src/tools/index.ts`** -- Export the new schema.
6. **`src/index.ts`** -- Register with `server.tool()`. Write a description covering: what it does, when to use it, return format, pagination, related tools.
7. **`test/unit.test.ts`** -- Test the AppleScript template, input escaping, and offset/limit handling.
8. **`scripts/smoke-test.mjs`** -- Update expected tool count and list. Add a happy-path test.
9. **Update tool count** in `README.md`, `CLAUDE.md`, and `scripts/smoke-test.mjs`.

## AppleScript gotchas

- **`whose` clause limitations:** Works on scalar properties (subject, display name) but NOT on collection properties (email addresses, attachments). For collections, iterate with a loop and try/catch. See the `searchMessages` two-phase pattern: phase 1 filters by subject via `whose`, phase 2 scans for sender matches.
- **`plain text content` is slow:** ~100ms per message. Only use in single-item `get_*` tools and `listMessages` (folder listing with preview), never in search results.
- **Date construction:** Use component-based approach (`set year of d to X`, `set month of d to Y`, ...) for locale safety. See `buildAppleScriptDateVar()`. Setting day to 1 first avoids month-overflow bugs (e.g., setting month to February when day is 31).
- **Timeout management:** Default is 30s. Date-range event queries use 60s. Search with large offset may need more. `searchTimeoutMs()` scales timeout based on offset.
- **Flag enums:** Outlook uses `not completed` / `completed` / `not flagged` for todo flag (not `flag marked` / `flag complete`). The `is completed` property does not exist on task objects; use `todo flag is completed` instead.
- **Apple epoch:** Outlook timestamps are seconds since 2001-01-01, not Unix epoch. Use `appleTimestampToIso()` / `isoToAppleTimestamp()`.

## Dependencies

| Dependency | Purpose |
|------------|---------|
| `@modelcontextprotocol/sdk` | MCP server framework (stdio transport) |
| `zod` | Input validation for all 52 tool handlers |
| `typescript` (dev) | Build toolchain (`strict: true`, `ES2022`, `NodeNext`) |
| `vitest` (dev) | Test runner |
| `@types/node` (dev) | Node.js type definitions |

No native dependencies. No network calls. No database.
