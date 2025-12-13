# Claude Outlook MCP Tool

This is a Model Context Protocol (MCP) tool that allows Claude to interact with Microsoft Outlook for macOS.

<a href="https://glama.ai/mcp/servers/0j71n92wnh">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/0j71n92wnh/badge" alt="Claude Outlook Tool MCP server" />
</a>

## Features

- **Email Management**
  - Read, search, and filter emails by date range
  - Send emails with HTML support and attachments
  - Reply, reply-all, and forward with comments
  - Create drafts for manual editing
  - Save attachments to disk
- **Folder Management**
  - Create, rename, and delete folders
  - Move emails between folders
  - Support for nested folder paths
- **Calendar Operations**
  - List, create, and delete events
  - Respond to meeting invitations (accept/decline/tentative)
  - Propose alternative meeting times
- **Contacts**
  - List and search contacts

## Prerequisites

- macOS with Apple Silicon (M1/M2/M3) or Intel chip
- [Microsoft Outlook for Mac](https://apps.apple.com/us/app/microsoft-outlook/id985367838) installed and configured
- [Bun](https://bun.sh/) installed
- [Claude desktop app](https://claude.ai/desktop) installed

## Installation

1. Clone this repository:

```bash
git clone https://github.com/syedazharmbnr1/claude-outlook-mcp.git
cd claude-outlook-mcp
```

2. Install dependencies:

```bash
bun install
```

3. Make sure the script is executable:

```bash
chmod +x index.ts
```

4. Update your Claude Desktop configuration:

Edit your `claude_desktop_config.json` file (located at `~/Library/Application Support/Claude/claude_desktop_config.json`) to include this tool:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "/Users/YOURUSERNAME/.bun/bin/bun",
      "args": ["run", "/path/to/claude-outlook-mcp/index.ts"]
    }
  }
}
```

Make sure to replace `YOURUSERNAME` with your actual macOS username and adjust the path to where you cloned this repository.

5. Restart Claude Desktop app

6. Grant permissions:
   - Go to System Preferences > Privacy & Security > Privacy
   - Give Terminal (or your preferred terminal app) access to Accessibility features
   - You may see permission prompts when the tool is first used

## Operations Reference

### outlook_mail

| Operation | Parameters | Description |
|-----------|------------|-------------|
| `read` | `folder`, `limit`, `startDate?`, `endDate?` | Read emails with optional date filtering |
| `unread` | `folder`, `limit` | Get unread emails |
| `search` | `searchTerm`, `folder`, `limit`, `startDate?`, `endDate?` | Search by content or subject |
| `send` | `to`, `subject`, `body`, `cc?`, `bcc?`, `isHtml?`, `attachments?` | Send an email |
| `reply` | `messageId`, `body`, `isHtml?` | Reply to an email |
| `replyAll` | `messageId`, `body`, `isHtml?` | Reply-all to an email |
| `forward` | `messageId`, `to`, `cc?`, `comment?`, `attachments?`, `includeOriginalAttachments?` | Forward an email |
| `create_draft` | `to`, `subject`, `body`, `cc?`, `bcc?`, `isHtml?`, `attachments?` | Create draft (opens in Outlook) |
| `move` | `messageId`, `targetFolder` | Move email to folder |
| `count` | `folder` | Count emails in folder |
| `save_attachments` | `messageId`, `destinationFolder` | Save attachments to disk |
| `folders` | — | List all mail folders |
| `create_folder` | `name`, `parent?` | Create a folder |
| `rename_folder` | `path`, `newName` | Rename a folder |
| `delete_folder` | `path` | Delete a folder |

### outlook_calendar

| Operation | Parameters | Description |
|-----------|------------|-------------|
| `list` | `startDate?`, `endDate?`, `limit?` | List calendar events |
| `create` | `subject`, `startDate`, `startTime`, `endDate?`, `endTime?`, `location?`, `body?`, `attendees?`, `isAllDay?` | Create event |
| `delete` | `eventId` | Delete event |
| `accept` | `eventId`, `comment?` | Accept meeting invitation |
| `decline` | `eventId`, `comment?` | Decline meeting invitation |
| `tentative` | `eventId`, `comment?` | Tentatively accept meeting |
| `propose_new_time` | `eventId`, `newStartDate`, `newStartTime`, `newEndDate?`, `newEndTime?`, `comment?` | Propose new meeting time |

### outlook_contacts

| Operation | Parameters | Description |
|-----------|------------|-------------|
| `list` | `limit?` | List contacts |
| `search` | `query`, `limit?` | Search contacts by name |

### Parameter Formats

| Parameter | Format | Example |
|-----------|--------|---------|
| Date | `YYYY-MM-DD` | `2025-12-13` |
| Time | `HH:MM` (24-hour) | `14:30` |
| Folder path | Slash-separated | `Work/Projects/Active` |
| Recipients | Comma-separated | `a@example.com, b@example.com` |
| Attachments | Array of paths | `["/path/to/file.pdf"]` |

> Parameters with `?` are optional.

## Usage

Once installed, you can use the Outlook tool directly from Claude by asking questions like:

- "Can you check my unread emails in Outlook?"
- "Search my Outlook emails for the quarterly report"
- "Send an email to john@example.com with the subject 'Meeting Tomorrow'"
- "What's on my calendar today?"
- "Create a meeting for tomorrow at 2pm"
- "Find the contact information for Jane Smith"

## Examples

### Natural Language (via Claude)

```
Check my unread emails in Outlook
Search my emails for "budget meeting"
Forward the last email from John to my team with a note "FYI"
Move all emails from newsletters@example.com to my Archive folder
What meetings do I have tomorrow?
```

### JSON Examples (direct MCP calls)

**Read emails with date filter:**
```json
{
  "operation": "read",
  "folder": "Inbox",
  "limit": 20,
  "startDate": "2025-12-01",
  "endDate": "2025-12-13"
}
```

**Forward with comment:**
```json
{
  "operation": "forward",
  "messageId": "12345",
  "to": "colleague@example.com",
  "comment": "FYI - see the discussion below"
}
```

**Create a meeting:**
```json
{
  "operation": "create",
  "subject": "Weekly Sync",
  "startDate": "2025-12-15",
  "startTime": "10:00",
  "endTime": "10:30",
  "attendees": "team@example.com",
  "location": "Conference Room A"
}
```

**Move email to nested folder:**
```json
{
  "operation": "move",
  "messageId": "12345",
  "targetFolder": "Work/Projects/Active"
}
```

## Troubleshooting

If you encounter issues with attachments:
- Check if the file exists and is readable
- Use absolute file paths instead of relative paths
- Make sure the user running the process has permission to read the file

If you encounter the error `Cannot find module '@modelcontextprotocol/sdk/server/index.js'`:

1. Make sure you've run `bun install` to install all dependencies
2. Try installing the MCP SDK explicitly:
   ```bash
   bun add @modelcontextprotocol/sdk@^1.5.0
   ```
3. Check if the module exists in your node_modules directory:
   ```bash
   ls -la node_modules/@modelcontextprotocol/sdk/server/
   ```

If the error persists, try creating a new project with Bun:

```bash
mkdir -p ~/yourpath/claude-outlook-mcp
cd ~/yourpath/claude-outlook-mcp
bun init -y
```

Then copy the package.json and index.ts files to the new directory and run:

```bash
bun install
bun run index.ts
```

Update your claude_desktop_config.json to point to the new location.

## License

MIT