# mcp-outlook-applescript

A [Model Context Protocol](https://modelcontextprotocol.io/) server for Microsoft Outlook on Mac. Access your mail, calendar, contacts, tasks, and notes through 50 MCP tools powered by AppleScript.

## Features

- **AppleScript backend** — works with classic Outlook for Mac, no network or authentication required
- **50 tools** — full coverage of mail, calendar, contacts, tasks, and notes
- **Approval system** — destructive operations (delete, move, junk) require a two-step prepare/confirm flow
- **Zero native dependencies** — pure TypeScript + AppleScript, no C++ compilation needed

### Available Tools

| Category | Tools |
|----------|-------|
| **Mail** | `list_folders`, `list_emails`, `search_emails`, `get_email`, `get_unread_count`, `send_email`, `create_draft`, `reply_to_email`, `reply_all_email`, `list_attachments`, `download_attachment` |
| **Mail Organization** | `mark_email_read`, `mark_email_unread`, `set_email_flag`, `clear_email_flag`, `set_email_categories`, `prepare/confirm_delete_email`, `prepare/confirm_move_email`, `prepare/confirm_archive_email`, `prepare/confirm_junk_email`, `prepare/confirm_batch_*` |
| **Folders** | `create_folder`, `rename_folder`, `move_folder`, `prepare/confirm_delete_folder`, `prepare/confirm_empty_folder` |
| **Calendar** | `list_calendars`, `list_events`, `get_event`, `search_events`, `create_event`, `update_event`, `delete_event`, `respond_to_event` |
| **Contacts** | `list_contacts`, `search_contacts`, `get_contact` |
| **Tasks** | `list_tasks`, `search_tasks`, `get_task` |
| **Notes** | `list_notes`, `search_notes`, `get_note` |
| **Accounts** | `list_accounts` |

## Install

```bash
git clone https://github.com/hasan-imam/mcp-outlook-applescript.git
cd mcp-outlook-applescript
npm install && npm run build && npm pack && npm install -g mcp-outlook-applescript-*.tgz
cd .. && rm -rf mcp-outlook-applescript
```

This installs `mcp-outlook-applescript` as a global command. The checkout can be deleted after install.

To update, re-run the commands above.

### Claude Code

```bash
claude mcp add outlook -- mcp-outlook-applescript
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "mcp-outlook-applescript"
    }
  }
}
```

## Requirements

- macOS with Microsoft Outlook (classic) installed and running
- Node.js >= 20
- Automation permission for Outlook (System Settings > Privacy & Security > Automation)

## Development

```bash
npm install
npm run build      # compile TypeScript → dist/
npm run typecheck   # type-check without emitting
npm test            # 121 unit tests
bash scripts/audit.sh  # static quality audit (build, security, package, functional)
```

## License

MIT — see [LICENSE](LICENSE).
