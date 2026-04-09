#!/usr/bin/env bun
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { runAppleScript } from "run-applescript";
import {
  readFileSync,
  realpathSync,
  accessSync,
  statSync,
  constants as fsConstants,
} from "node:fs";
import { access, stat } from "node:fs/promises";
import { resolve, dirname } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";

// ====================================================
// 0. Security Helpers & Configuration
// ====================================================

/**
 * Escape a string for safe interpolation into an AppleScript string literal.
 * Escapes backslashes first (so they don't re-escape quotes), then double quotes.
 */
function escapeForAppleScript(str: string): string {
  return str.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

/**
 * Generate AppleScript code to resolve a mail folder by name or slash-delimited path.
 *
 * Supports three forms:
 *   "Inbox"                        → Exchange inbox (special-cased)
 *   "Inbox/AWS Notifications/Sub"  → walk the folder tree from inbox
 *   "AWS Notifications"            → flat name search (backward compatible)
 *
 * The generated code stores the resolved folder reference in `varName`.
 */
function folderResolverScript(
  folder: string,
  varName: string = "targetFolder",
): string {
  const escaped = escapeForAppleScript(folder);
  return `
        -- ── Resolve folder: "${escaped}" ──
        set ${varName} to missing value
        set _folderPath to "${escaped}"
        if _folderPath is "Inbox" then
          try
            set ${varName} to inbox of exchange account 1
          on error
            set ${varName} to inbox
          end try
        else if _folderPath contains "/" then
          -- Path-based resolution: walk the tree
          set saveTID to AppleScript's text item delimiters
          set AppleScript's text item delimiters to "/"
          set pathParts to text items of _folderPath
          set AppleScript's text item delimiters to saveTID

          set firstPart to item 1 of pathParts
          if firstPart is "Inbox" then
            try
              set ${varName} to inbox of exchange account 1
            on error
              set ${varName} to inbox
            end try
          else
            try
              repeat with mf in (mail folders of exchange account 1)
                if name of mf is firstPart then
                  set ${varName} to mf
                  exit repeat
                end if
              end repeat
            on error
              -- Fall back to flat search for first component
              repeat with mf in (every mail folder)
                if name of mf is firstPart then
                  set ${varName} to mf
                  exit repeat
                end if
              end repeat
            end try
          end if

          if ${varName} is not missing value and (count of pathParts) > 1 then
            repeat with i from 2 to count of pathParts
              set subName to item i of pathParts
              set foundSub to false
              repeat with sf in (mail folders of ${varName})
                if name of sf is subName then
                  set ${varName} to sf
                  set foundSub to true
                  exit repeat
                end if
              end repeat
              if not foundSub then
                set ${varName} to missing value
                exit repeat
              end if
            end repeat
          end if
        else
          -- Flat name search (backward compatible)
          repeat with mf in (every mail folder)
            if name of mf is _folderPath then
              set ${varName} to mf
              exit repeat
            end if
          end repeat
        end if
        if ${varName} is missing value then
          error "Folder not found: " & _folderPath
        end if`;
}

/**
 * Generate AppleScript code to extract sender name + address from a message.
 *
 * Tries `name of sender` and `address of sender` first, combining them as
 * "Display Name <addr@example.com>".  Falls back to coercing the sender
 * object to text, then to "unknown".
 */
function senderScript(msgVar: string, resultVar: string): string {
  return `
            set ${resultVar} to "unknown"
            try
              set _senderObj to sender of ${msgVar}
              set _sName to ""
              set _sAddr to ""
              try
                set _sName to name of _senderObj
              end try
              try
                set _sAddr to address of _senderObj
              end try
              if _sName is not "" and _sAddr is not "" then
                set ${resultVar} to _sName & " <" & _sAddr & ">"
              else if _sName is not "" then
                set ${resultVar} to _sName
              else if _sAddr is not "" then
                set ${resultVar} to _sAddr
              end if
            on error
              -- Never fall back to "sender as text" — it produces binary garbage
              -- for Exchange users. Just leave as "unknown".
              set ${resultVar} to "unknown"
            end try`;
}

/**
 * Load attachment allowlist from config.json.
 * Falls back to sensible defaults if the config file is missing or malformed.
 */
let _cachedAllowlist: string[] | null = null;

function loadAttachmentAllowlist(): string[] {
  if (_cachedAllowlist) return _cachedAllowlist;

  const defaults = ["~/Downloads", "~/Documents", "~/Desktop"];
  try {
    const configPath = resolve(
      dirname(fileURLToPath(import.meta.url)),
      "config.json",
    );
    const raw = readFileSync(configPath, "utf-8");
    const config = JSON.parse(raw);
    const dirs = config?.attachments?.allowedDirectories;
    if (Array.isArray(dirs) && dirs.length > 0) {
      _cachedAllowlist = dirs;
      return dirs;
    }
    console.error(
      "[config] No allowedDirectories in config.json, using defaults",
    );
    _cachedAllowlist = defaults;
    return defaults;
  } catch (err) {
    console.error(
      "[config] Could not load config.json, using defaults:",
      (err as Error).message,
    );
    _cachedAllowlist = defaults;
    return defaults;
  }
}

/** Resolve ~ to the actual home directory */
function expandTilde(p: string): string {
  if (p.startsWith("~/") || p === "~") {
    return p.replace("~", homedir());
  }
  return p;
}

/**
 * Validate that a file path is within the attachment allowlist.
 * Resolves symlinks and relative segments before checking.
 * Throws with a descriptive message if the path is outside all allowed directories.
 */
function validateAttachmentPath(filePath: string): string {
  // Resolve to absolute path
  let absolutePath = filePath;
  if (!filePath.startsWith("/")) {
    absolutePath = resolve(process.cwd(), filePath);
  }

  // Canonicalize: resolve symlinks and .. segments
  let canonicalPath: string;
  try {
    // realpathSync requires the file to exist
    canonicalPath = realpathSync(absolutePath);
  } catch {
    // File may not exist yet (unlikely for attachments, but handle gracefully)
    canonicalPath = resolve(absolutePath);
  }

  // Check against allowlist
  const allowedDirs = loadAttachmentAllowlist();
  const expandedAllowed = allowedDirs.map((d) => expandTilde(d));

  for (const allowedDir of expandedAllowed) {
    let canonicalAllowed: string;
    try {
      canonicalAllowed = realpathSync(allowedDir);
    } catch {
      canonicalAllowed = resolve(allowedDir);
    }
    // Ensure trailing slash for prefix check to prevent /Users/x/DesktopEvil matching /Users/x/Desktop
    const prefix = canonicalAllowed.endsWith("/")
      ? canonicalAllowed
      : canonicalAllowed + "/";
    if (
      canonicalPath === canonicalAllowed ||
      canonicalPath.startsWith(prefix)
    ) {
      return canonicalPath;
    }
  }

  // Rejected — build a helpful error message
  const displayAllowed = allowedDirs.join(", ");
  throw new Error(
    `SECURITY: Attachment path "${filePath}" (resolved to "${canonicalPath}") is outside allowed directories [${displayAllowed}]. ` +
      `To allow this path, add its parent directory to the "attachments.allowedDirectories" array in config.json ` +
      `(located at ${resolve(dirname(fileURLToPath(import.meta.url)), "config.json")}).`,
  );
}

// ====================================================
// 1. Tool Definitions
// ====================================================

// Define Outlook Mail tool
const OUTLOOK_MAIL_TOOL: Tool = {
  name: "outlook_mail",
  description:
    "Interact with Microsoft Outlook for macOS - read, search, send, delete, move, archive, forward, and manage emails",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description:
          "Operation to perform: 'unread', 'search', 'send', 'folders', 'read', 'delete', 'move', 'mark_read', 'count', 'get_message', 'forward', or 'archive'",
        enum: [
          "unread",
          "search",
          "send",
          "folders",
          "read",
          "delete",
          "move",
          "mark_read",
          "count",
          "get_message",
          "forward",
          "archive",
        ],
      },
      folder: {
        type: "string",
        description:
          "Email folder name or slash-delimited path (e.g. 'Inbox', 'Inbox/AWS Notifications/Pfizer CPDB ETL', '_Projects'). Paths walk the folder tree; plain names do a flat search. Defaults to Inbox.",
      },
      limit: {
        type: "number",
        description:
          "Number of emails to retrieve (optional, for unread, read, and search operations)",
      },
      searchTerm: {
        type: "string",
        description:
          "Text to search for in emails (required for search operation)",
      },
      to: {
        type: "string",
        description: "Recipient email address (required for send operation)",
      },
      subject: {
        type: "string",
        description: "Email subject (required for send operation)",
      },
      body: {
        type: "string",
        description: "Email body content (required for send operation)",
      },
      isHtml: {
        type: "boolean",
        description:
          "Whether the body content is HTML (optional for send operation, default: false)",
      },
      cc: {
        type: "string",
        description: "CC email address (optional for send operation)",
      },
      bcc: {
        type: "string",
        description: "BCC email address (optional for send operation)",
      },
      attachments: {
        type: "array",
        description:
          "File paths to attach to the email (optional for send operation)",
        items: {
          type: "string",
        },
      },
      subjectFilter: {
        type: "string",
        description:
          "Filter messages by subject containing this text (for delete, move, mark_read, count, and archive operations)",
      },
      destinationFolder: {
        type: "string",
        description: "Destination folder name (required for move operation)",
      },
      messageIndex: {
        type: "number",
        description:
          "1-based index of the message in the folder (required for get_message and forward operations)",
      },
      markAsRead: {
        type: "boolean",
        description:
          "Whether to mark as read (true) or unread (false). Default is true. (for mark_read operation)",
      },
      comment: {
        type: "string",
        description:
          "Optional comment to prepend when forwarding an email (for forward operation)",
      },
    },
    required: ["operation"],
  },
};

// Define Outlook Calendar tool
const OUTLOOK_CALENDAR_TOOL: Tool = {
  name: "outlook_calendar",
  description:
    "Interact with Microsoft Outlook for macOS calendar - view, create, and manage events",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description:
          "Operation to perform: 'today', 'upcoming', 'search', or 'create'",
        enum: ["today", "upcoming", "search", "create"],
      },
      searchTerm: {
        type: "string",
        description:
          "Text to search for in events (required for search operation)",
      },
      limit: {
        type: "number",
        description:
          "Number of events to retrieve (optional, for today and upcoming operations)",
      },
      days: {
        type: "number",
        description:
          "Number of days to look ahead (optional, for upcoming operation, default: 7)",
      },
      subject: {
        type: "string",
        description: "Event subject/title (required for create operation)",
      },
      start: {
        type: "string",
        description: "Start time in ISO format (required for create operation)",
      },
      end: {
        type: "string",
        description: "End time in ISO format (required for create operation)",
      },
      location: {
        type: "string",
        description: "Event location (optional for create operation)",
      },
      body: {
        type: "string",
        description: "Event description/body (optional for create operation)",
      },
      attendees: {
        type: "string",
        description:
          "Comma-separated list of attendee email addresses (optional for create operation)",
      },
    },
    required: ["operation"],
  },
};

// Define Outlook Contacts tool
const OUTLOOK_CONTACTS_TOOL: Tool = {
  name: "outlook_contacts",
  description: "Search and retrieve contacts from Microsoft Outlook for macOS",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description: "Operation to perform: 'list' or 'search'",
        enum: ["list", "search"],
      },
      searchTerm: {
        type: "string",
        description:
          "Text to search for in contacts (required for search operation)",
      },
      limit: {
        type: "number",
        description: "Number of contacts to retrieve (optional)",
      },
    },
    required: ["operation"],
  },
};

// ====================================================
// 2. Server Setup
// ====================================================

console.error("Starting Outlook MCP server...");

const server = new Server(
  {
    name: "Outlook MCP Tool",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  },
);

// ====================================================
// 3. Core Functions
// ====================================================

// Check if Outlook is installed and running
async function checkOutlookAccess(): Promise<boolean> {
  console.error("[checkOutlookAccess] Checking if Outlook is accessible...");
  try {
    const isInstalled = await runAppleScript(`
      tell application "System Events"
        set outlookExists to exists application process "Microsoft Outlook"
        return outlookExists
      end tell
    `);

    if (isInstalled !== "true") {
      console.error(
        "[checkOutlookAccess] Microsoft Outlook is not installed or running",
      );
      throw new Error(
        "Microsoft Outlook is not installed or running on this system",
      );
    }

    const isRunning = await runAppleScript(`
      tell application "System Events"
        set outlookRunning to application process "Microsoft Outlook" exists
        return outlookRunning
      end tell
    `);

    if (isRunning !== "true") {
      console.error(
        "[checkOutlookAccess] Microsoft Outlook is not running, attempting to launch...",
      );
      try {
        await runAppleScript(`
          tell application "Microsoft Outlook" to activate
          delay 2
        `);
        console.error("[checkOutlookAccess] Launched Outlook successfully");
      } catch (activateError) {
        console.error(
          "[checkOutlookAccess] Error activating Microsoft Outlook:",
          activateError,
        );
        throw new Error(
          "Could not activate Microsoft Outlook. Please start it manually.",
        );
      }
    } else {
      console.error(
        "[checkOutlookAccess] Microsoft Outlook is already running",
      );
    }

    return true;
  } catch (error) {
    console.error("[checkOutlookAccess] Outlook access check failed:", error);
    throw new Error(
      `Cannot access Microsoft Outlook. Please make sure Outlook is installed and properly configured. Error: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

// ====================================================
// 4. EMAIL FUNCTIONS
// ====================================================

// Function to get unread emails
async function getUnreadEmails(
  folder: string = "Inbox",
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[getUnreadEmails] Getting unread emails from folder: ${folder}, limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "theFolder")}

        set RS to ASCII character 30
        set US to ASCII character 31

        -- Build a flat list of folders to scan: root + subfolders up to 3 levels deep
        set foldersToScan to {theFolder}
        try
          repeat with sf1 in (mail folders of theFolder)
            set end of foldersToScan to (contents of sf1)
            try
              repeat with sf2 in (mail folders of sf1)
                set end of foldersToScan to (contents of sf2)
                try
                  repeat with sf3 in (mail folders of sf2)
                    set end of foldersToScan to (contents of sf3)
                  end repeat
                end try
              end repeat
            end try
          end repeat
        end try

        set totalFound to 0
        set totalUnread to 0
        set msgLimit to ${limit}
        set output to ""

        -- Scan each folder for unread messages
        repeat with scanFolder in foldersToScan
          set folderName to name of scanFolder
          set unreadMsgs to (every message of scanFolder whose is read is false)
          set unreadCount to count of unreadMsgs
          set totalUnread to totalUnread + unreadCount

          if unreadCount > 0 and totalFound < msgLimit then
            set remaining to msgLimit - totalFound
            if unreadCount < remaining then set remaining to unreadCount

            repeat with i from 1 to remaining
              set theMsg to item i of unreadMsgs

              set msgSubject to subject of theMsg
              ${senderScript("theMsg", "senderInfo")}
              set msgDate to time sent of theMsg as text
              set msgId to id of theMsg as text

              -- Try to get content preview
              set msgContent to "[Content not available]"
              try
                set rawContent to content of theMsg
                if length of rawContent > 500 then
                  set msgContent to (text 1 thru 500 of rawContent) & "..."
                else
                  set msgContent to rawContent
                end if
              end try

              set msgLine to msgSubject & US & senderInfo & US & msgDate & US & msgContent & US & msgId & US & folderName
              if totalFound > 0 then
                set output to output & RS & msgLine
              else
                set output to msgLine
              end if
              set totalFound to totalFound + 1
            end repeat
          end if

          if totalFound >= msgLimit then exit repeat
        end repeat

        return (totalUnread as text) & US & output
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getUnreadEmails] Raw result length: ${result.length}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    // First field before the first US is the total unread count, rest is message data
    const headerSplit = result.indexOf("\x1F");
    if (headerSplit === -1) {
      // Only got the count, no messages (e.g. "0" with empty output)
      console.error("[getUnreadEmails] Found 0 unread emails");
      return [];
    }

    const totalUnread = parseInt(result.substring(0, headerSplit), 10) || 0;
    const messagesStr = result.substring(headerSplit + 1);

    const emails: any[] = [];
    if (messagesStr.trim() === "") {
      console.error(
        `[getUnreadEmails] ${totalUnread} total unread, 0 returned`,
      );
      return emails;
    }

    const messageChunks = messagesStr.split("\x1E"); // RS between messages
    for (const chunk of messageChunks) {
      if (!chunk.trim()) continue;
      const parts = chunk.split("\x1F"); // US between fields
      emails.push({
        subject: parts[0] || "No subject",
        sender: parts[1] || "Unknown sender",
        dateSent: parts[2] || new Date().toString(),
        content: parts[3] || "[Content not available]",
        id: parts[4] || "",
        folder: parts[5] || folder,
      });
    }

    console.error(
      `[getUnreadEmails] Found ${emails.length} unread emails (${totalUnread} total unread across all subfolders)`,
    );
    return emails;
  } catch (error) {
    console.error("[getUnreadEmails] Error getting unread emails:", error);
    throw error;
  }
}

// Function to search emails
async function searchEmails(
  searchTerm: string,
  folder: string = "Inbox",
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[searchEmails] Searching for "${searchTerm}" in folder: ${folder}, limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "theFolder")}

        set RS to ASCII character 30
        set US to ASCII character 31
        set searchString to "${escapeForAppleScript(searchTerm)}"
        set allMessages to messages of theFolder
        set i to 0
        set output to ""

        repeat with theMessage in allMessages
          try
            if (subject of theMessage contains searchString) or (content of theMessage contains searchString) then
              set i to i + 1

              set msgSubject to subject of theMessage
              ${senderScript("theMessage", "senderInfo")}
              set msgDate to time sent of theMessage as text
              set msgId to id of theMessage as text

              -- Try to get content preview
              set msgContent to "[Content not available]"
              try
                set rawContent to content of theMessage
                if length of rawContent > 500 then
                  set msgContent to (text 1 thru 500 of rawContent) & "..."
                else
                  set msgContent to rawContent
                end if
              end try

              set msgLine to msgSubject & US & senderInfo & US & msgDate & US & msgContent & US & msgId
              if i > 1 then
                set output to output & RS & msgLine
              else
                set output to msgLine
              end if

              if i >= ${limit} then exit repeat
            end if
          on error
            -- Skip problematic messages
          end try
        end repeat

        return (i as text) & US & output
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[searchEmails] Raw result length: ${result.length}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    // First field before the first US is the match count, rest is message data
    const headerSplit = result.indexOf("\x1F");
    if (headerSplit === -1) {
      console.error("[searchEmails] Found 0 matching emails");
      return [];
    }

    const matchCount = parseInt(result.substring(0, headerSplit), 10) || 0;
    const messagesStr = result.substring(headerSplit + 1);

    const emails: any[] = [];
    if (messagesStr.trim() === "") {
      console.error(`[searchEmails] ${matchCount} matches, 0 returned`);
      return emails;
    }

    const messageChunks = messagesStr.split("\x1E"); // RS between messages
    for (const chunk of messageChunks) {
      if (!chunk.trim()) continue;
      const parts = chunk.split("\x1F"); // US between fields
      emails.push({
        subject: parts[0] || "No subject",
        sender: parts[1] || "Unknown sender",
        dateSent: parts[2] || new Date().toString(),
        content: parts[3] || "[Content not available]",
        id: parts[4] || "",
      });
    }

    console.error(`[searchEmails] Found ${emails.length} matching emails`);
    return emails;
  } catch (error) {
    console.error("[searchEmails] Error searching emails:", error);
    throw error;
  }
}

async function checkAttachmentPath(filePath: string): Promise<string> {
  try {
    // Convert to absolute path if relative
    let fullPath = filePath;
    if (!filePath.startsWith("/")) {
      const cwd = process.cwd();
      fullPath = `${cwd}/${filePath}`;
    }

    // Security: validate against attachment allowlist
    fullPath = validateAttachmentPath(fullPath);

    // Check if the file exists and is readable
    try {
      await access(fullPath, fsConstants.R_OK);
      const stats = await stat(fullPath);

      return `File exists and is readable: ${fullPath}\nSize: ${stats.size} bytes\nPermissions: ${stats.mode.toString(8)}\nLast modified: ${stats.mtime}`;
    } catch (err) {
      return `ERROR: Cannot access file: ${fullPath}\nError details: ${(err as Error).message}`;
    }
  } catch (error) {
    return `Failed to check attachment path: ${(error as Error).message}`;
  }
}

// Add a debug version of sending email with attachment to test if files are accessible
async function debugSendEmailWithAttachment(
  to: string,
  subject: string,
  body: string,
  attachmentPath: string,
): Promise<string> {
  // First check if the file exists and is readable
  const fileStatus = await checkAttachmentPath(attachmentPath);
  console.error(`[debugSendEmail] Attachment status: ${fileStatus}`);

  // Create a simple AppleScript that just attempts to open the file
  const script = `
    set theFile to POSIX file "${escapeForAppleScript(attachmentPath)}"
    try
      tell application "Finder"
        set fileExists to exists file theFile
        set fileInfo to info for file theFile
        return "File exists: " & fileExists & ", size: " & (size of fileInfo)
      end tell
    on error errMsg
      return "Error accessing file: " & errMsg
    end try
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[debugSendEmail] AppleScript file check: ${result}`);

    // Now try to actually create a draft with the attachment
    const emailScript = `
      tell application "Microsoft Outlook"
        try
          set newMessage to make new outgoing message with properties {subject:"DEBUG: ${escapeForAppleScript(subject)}", visible:true}
          set content of newMessage to "${escapeForAppleScript(body)}"
          set to recipients of newMessage to {"${escapeForAppleScript(to)}"}

          try
            set attachmentFile to POSIX file "${escapeForAppleScript(attachmentPath)}"
            make new attachment at newMessage with properties {file:attachmentFile}
            set attachResult to "Successfully attached file"
          on error attachErrMsg
            set attachResult to "Failed to attach file: " & attachErrMsg
          end try

          return attachResult
        on error errMsg
          return "Error creating email: " & errMsg
        end try
      end tell
    `;

    const attachResult = await runAppleScript(emailScript);
    console.error(`[debugSendEmail] Attachment result: ${attachResult}`);

    return `File check: ${fileStatus}\n\nAttachment test: ${attachResult}`;
  } catch (error) {
    console.error("[debugSendEmail] Error during debug:", error);
    return `Debugging error: ${(error as Error).message}\n\nFile check: ${fileStatus}`;
  }
}
// Update the sendEmail function to handle attachments and HTML content
async function sendEmail(
  to: string,
  subject: string,
  body: string,
  cc?: string,
  bcc?: string,
  isHtml: boolean = false,
  attachments?: string[],
): Promise<string> {
  console.error(`[sendEmail] Sending email to: ${to}, subject: "${subject}"`);
  console.error(
    `[sendEmail] Attachments: ${attachments ? JSON.stringify(attachments) : "none"}`,
  );

  await checkOutlookAccess();

  // Extract name from email if possible (for display name)
  const extractNameFromEmail = (email: string): string => {
    const namePart = email.split("@")[0];
    return namePart
      .split(".")
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
      .join(" ");
  };

  // Get name for display
  const toName = extractNameFromEmail(to);
  const ccName = cc ? extractNameFromEmail(cc) : "";
  const bccName = bcc ? extractNameFromEmail(bcc) : "";

  // Escape special characters
  const escapedSubject = escapeForAppleScript(subject);
  const escapedBody = escapeForAppleScript(body).replace(/\n/g, "\\n");

  // Process attachments: Convert to absolute paths if they are relative
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    for (const attachPath of attachments) {
      // validateAttachmentPath resolves to absolute, canonicalizes, and checks allowlist.
      // Throws with a descriptive error if the path is outside allowed directories.
      const validated = validateAttachmentPath(attachPath);
      processedAttachments.push(validated);
    }
    console.error(
      `[sendEmail] Validated attachments: ${processedAttachments.length} file(s) passed allowlist check`,
    );
  }

  // Create attachment script part with better error handling
  const attachmentScript =
    processedAttachments.length > 0
      ? processedAttachments
          .map((filePath) => {
            const escapedPath = escapeForAppleScript(filePath);
            return `
        try
          set attachmentFile to POSIX file "${escapedPath}"
          make new attachment at msg with properties {file:attachmentFile}
          log "Successfully attached file: ${escapedPath}"
        on error errMsg
          log "Failed to attach file: ${escapedPath} - Error: " & errMsg
        end try
      `;
          })
          .join("\n")
      : "";

  // Try approach 1: Using specific syntax for creating a message with attachments
  try {
    const script1 = `
      tell application "Microsoft Outlook"
        try
          set msg to make new outgoing message with properties {subject:"${escapedSubject}"}

          ${
            isHtml
              ? `set content type of msg to HTML
             set content of msg to "${escapedBody}"`
              : `set content of msg to "${escapedBody}"`
          }

          tell msg
            set recipTo to make new to recipient with properties {email address:{name:"${escapeForAppleScript(toName)}", address:"${escapeForAppleScript(to)}"}}
            ${cc ? `set recipCc to make new cc recipient with properties {email address:{name:"${escapeForAppleScript(ccName)}", address:"${escapeForAppleScript(cc)}"}}` : ""}
            ${bcc ? `set recipBcc to make new bcc recipient with properties {email address:{name:"${escapeForAppleScript(bccName)}", address:"${escapeForAppleScript(bcc)}"}}` : ""}

            ${attachmentScript}
          end tell

          -- Delay to allow attachments to be processed
          delay 1

          send msg
          return "Email sent successfully with attachments"
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;

    console.error("[sendEmail] Executing AppleScript method 1");
    const result = await runAppleScript(script1);
    console.error(`[sendEmail] Result (method 1): ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    return result;
  } catch (error1) {
    console.error("[sendEmail] Method 1 failed:", error1);

    // Try approach 2: Using AppleScript's draft window method
    try {
      const script2 = `
        tell application "Microsoft Outlook"
          try
            set newDraft to make new draft window
            set theMessage to item 1 of mail items of newDraft
            set subject of theMessage to "${escapedSubject}"

            ${
              isHtml
                ? `set content type of theMessage to HTML
               set content of theMessage to "${escapedBody}"`
                : `set content of theMessage to "${escapedBody}"`
            }

            set to recipients of theMessage to {"${escapeForAppleScript(to)}"}
            ${cc ? `set cc recipients of theMessage to {"${escapeForAppleScript(cc)}"}` : ""}
            ${bcc ? `set bcc recipients of theMessage to {"${escapeForAppleScript(bcc)}"}` : ""}

            ${processedAttachments
              .map((filePath) => {
                const escapedPath = escapeForAppleScript(filePath);
                return `
                try
                  set attachmentFile to POSIX file "${escapedPath}"
                  make new attachment at theMessage with properties {file:attachmentFile}
                  log "Successfully attached file: ${escapedPath}"
                on error attachErrMsg
                  log "Failed to attach file: ${escapedPath} - Error: " & attachErrMsg
                end try
              `;
              })
              .join("\n")}

            -- Delay to allow attachments to be processed
            delay 1

            send theMessage
            return "Email sent successfully with method 2"
          on error errMsg
            return "Error: " & errMsg
          end try
        end tell
      `;

      console.error("[sendEmail] Executing AppleScript method 2");
      const result = await runAppleScript(script2);
      console.error(`[sendEmail] Result (method 2): ${result}`);

      if (result.startsWith("Error:")) {
        throw new Error(result);
      }

      return result;
    } catch (error2) {
      console.error("[sendEmail] Method 2 failed:", error2);

      // Try approach 3: Create a draft for the user to manually send
      try {
        const script3 = `
          tell application "Microsoft Outlook"
            try
              set newMessage to make new outgoing message with properties {subject:"${escapedSubject}", visible:true}

              ${
                isHtml
                  ? `set content type of newMessage to HTML
                 set content of newMessage to "${escapedBody}"`
                  : `set content of newMessage to "${escapedBody}"`
              }

              set to recipients of newMessage to {"${escapeForAppleScript(to)}"}
              ${cc ? `set cc recipients of newMessage to {"${escapeForAppleScript(cc)}"}` : ""}
              ${bcc ? `set bcc recipients of newMessage to {"${escapeForAppleScript(bcc)}"}` : ""}

              ${processedAttachments
                .map((filePath) => {
                  const escapedPath = escapeForAppleScript(filePath);
                  return `
                  try
                    set attachmentFile to POSIX file "${escapedPath}"
                    make new attachment at newMessage with properties {file:attachmentFile}
                    log "Successfully attached file: ${escapedPath}"
                  on error attachErrMsg
                    log "Failed to attach file: ${escapedPath} - Error: " & attachErrMsg
                  end try
                `;
                })
                .join("\n")}

              -- Display the message
              activate
              return "Email draft created with attachments. Please review and send manually."
            on error errMsg
              return "Error: " & errMsg
            end try
          end tell
        `;

        console.error("[sendEmail] Executing AppleScript method 3");
        const result = await runAppleScript(script3);
        console.error(`[sendEmail] Result (method 3): ${result}`);

        if (result.startsWith("Error:")) {
          throw new Error(result);
        }

        return "A draft has been created in Outlook with the content and attachments. Please review and send it manually.";
      } catch (error3) {
        console.error("[sendEmail] All methods failed:", error3);
        throw new Error(
          `Could not send or create email. Please check if Outlook is properly configured and that you have granted necessary permissions. Error details: ${error3}`,
        );
      }
    }
  }
}
// Function to get mail folders - this works based on your logs
async function getMailFolders(): Promise<string[]> {
  console.error("[getMailFolders] Getting mail folder tree");
  await checkOutlookAccess();

  const script = `
      tell application "Microsoft Outlook"
        set LF to (ASCII character 10)
        set output to ""

        try
          set acct to exchange account 1
          set topFolders to mail folders of acct

          repeat with f1 in topFolders
            set n1 to name of f1
            set output to output & n1 & LF

            -- Level 2: direct subfolders
            try
              set subs1 to mail folders of f1
              repeat with f2 in subs1
                set n2 to name of f2
                set output to output & "  " & n2 & "  ::  " & n1 & "/" & n2 & LF

                -- Level 3
                try
                  set subs2 to mail folders of f2
                  repeat with f3 in subs2
                    set n3 to name of f3
                    set output to output & "    " & n3 & "  ::  " & n1 & "/" & n2 & "/" & n3 & LF

                    -- Level 4
                    try
                      set subs3 to mail folders of f3
                      repeat with f4 in subs3
                        set n4 to name of f4
                        set output to output & "      " & n4 & "  ::  " & n1 & "/" & n2 & "/" & n3 & "/" & n4 & LF
                      end repeat
                    end try
                  end repeat
                end try
              end repeat
            end try
          end repeat

        on error errMsg
          -- Fallback: flat list from every mail folder
          set allFolders to every mail folder
          repeat with mf in allFolders
            set output to output & name of mf & LF
          end repeat
        end try

        return output
      end tell
    `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getMailFolders] Result length: ${result.length}`);
    return result.split("\n").filter((line) => line.trim() !== "");
  } catch (error) {
    console.error("[getMailFolders] Error getting mail folders:", error);
    throw error;
  }
}

// Function to read emails in a folder that uses simple AppleScript
async function readEmails(
  folder: string = "Inbox",
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[readEmails] Reading emails from folder: ${folder}, limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
      tell application "Microsoft Outlook"
        try
          ${folderResolverScript(folder, "targetFolder")}

          set RS to ASCII character 30
          set US to ASCII character 31

          set allMsgs to messages of targetFolder
          set msgCount to 0
          set output to ""

          repeat with i from 1 to (count of allMsgs)
            if msgCount >= ${limit} then exit repeat

            try
              set theMsg to item i of allMsgs

              -- Filter: only process actual mail messages (skip calendar items, etc.)
              set msgClass to class of theMsg
              if msgClass is not incoming message and msgClass is not outgoing message then
                -- skip non-email objects like calendar notifications
              else
                set msgSubject to subject of theMsg
                ${senderScript("theMsg", "senderInfo")}
                set msgDate to time sent of theMsg as text
                set msgId to id of theMsg as text
                set isRead to is read of theMsg as text

                set msgLine to msgSubject & US & senderInfo & US & msgDate & US & msgId & US & isRead
                if msgCount > 0 then
                  set output to output & RS & msgLine
                else
                  set output to msgLine
                end if
                set msgCount to msgCount + 1
              end if
            on error
              -- Skip problematic messages
            end try
          end repeat

          return output
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;

  try {
    const result = await runAppleScript(script);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    if (result.trim() === "") {
      console.error("[readEmails] No emails found");
      return [];
    }

    // Parse using record separator (RS) between messages, unit separator (US) between fields
    const messageChunks = result.split("\x1E"); // RS
    const emails = messageChunks
      .filter((chunk) => chunk.trim() !== "")
      .map((chunk) => {
        const parts = chunk.split("\x1F"); // US
        return {
          subject: parts[0] || "No subject",
          sender: parts[1] || "Unknown sender",
          dateSent: parts[2] || new Date().toString(),
          id: parts[3] || "",
          isRead: parts[4] === "true",
          content: "Use get_message with messageIndex to retrieve full content",
        };
      });

    console.error(`[readEmails] Found ${emails.length} emails`);
    return emails;
  } catch (error) {
    console.error("[readEmails] Error reading emails:", error);
    throw error;
  }
}

// ====================================================
// 4b. EMAIL MANAGEMENT FUNCTIONS
// ====================================================

// Function to delete (soft-delete) emails from a folder
async function deleteEmails(
  folder: string = "Inbox",
  subjectFilter?: string,
  limit: number = 50,
): Promise<number> {
  console.error(
    `[deleteEmails] Deleting emails from folder: ${folder}, filter: "${subjectFilter || "none"}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const escapedFilter = subjectFilter
    ? escapeForAppleScript(subjectFilter)
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        set deletedCount to 0
        set useFilter to ${subjectFilter ? "true" : "false"}
        set filterText to "${escapedFilter}"

        repeat while deletedCount < ${limit}
          set msgList to messages of targetFolder
          if (count of msgList) is 0 then exit repeat

          set foundMatch to false
          set theMsg to item 1 of msgList

          if useFilter then
            -- Search for a matching message
            set foundMatch to false
            repeat with i from 1 to (count of msgList)
              set candidateMsg to item i of msgList
              try
                set msgSubject to subject of candidateMsg
                if msgSubject contains filterText then
                  set theMsg to candidateMsg
                  set foundMatch to true
                  exit repeat
                end if
              end try
            end repeat
            if not foundMatch then exit repeat
          else
            set foundMatch to true
          end if

          if foundMatch then
            try
              delete theMsg
              set deletedCount to deletedCount + 1
            on error errMsg
              exit repeat
            end try
          end if
        end repeat

        return deletedCount as text
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[deleteEmails] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const count = parseInt(result, 10) || 0;
    console.error(`[deleteEmails] Deleted ${count} emails`);
    return count;
  } catch (error) {
    console.error("[deleteEmails] Error deleting emails:", error);
    throw error;
  }
}

// Function to move emails from one folder to another
async function moveEmails(
  sourceFolder: string,
  destinationFolder: string,
  subjectFilter?: string,
  limit: number = 50,
): Promise<number> {
  console.error(
    `[moveEmails] Moving emails from "${sourceFolder}" to "${destinationFolder}", filter: "${subjectFilter || "none"}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const escapedFilter = subjectFilter
    ? escapeForAppleScript(subjectFilter)
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(sourceFolder, "targetFolder")}

        ${folderResolverScript(destinationFolder, "destFolder")}

        set movedCount to 0
        set useFilter to ${subjectFilter ? "true" : "false"}
        set filterText to "${escapedFilter}"

        repeat while movedCount < ${limit}
          set msgList to messages of targetFolder
          if (count of msgList) is 0 then exit repeat

          set foundMatch to false
          set theMsg to item 1 of msgList

          if useFilter then
            set foundMatch to false
            repeat with i from 1 to (count of msgList)
              set candidateMsg to item i of msgList
              try
                set msgSubject to subject of candidateMsg
                if msgSubject contains filterText then
                  set theMsg to candidateMsg
                  set foundMatch to true
                  exit repeat
                end if
              end try
            end repeat
            if not foundMatch then exit repeat
          else
            set foundMatch to true
          end if

          if foundMatch then
            try
              move theMsg to destFolder
              set movedCount to movedCount + 1
            on error errMsg
              exit repeat
            end try
          end if
        end repeat

        return movedCount as text
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[moveEmails] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const count = parseInt(result, 10) || 0;
    console.error(`[moveEmails] Moved ${count} emails`);
    return count;
  } catch (error) {
    console.error("[moveEmails] Error moving emails:", error);
    throw error;
  }
}

// Function to mark emails as read or unread
async function markEmailsRead(
  folder: string = "Inbox",
  subjectFilter?: string,
  limit: number = 50,
  markAsRead: boolean = true,
): Promise<number> {
  console.error(
    `[markEmailsRead] Marking emails in "${folder}" as ${markAsRead ? "read" : "unread"}, filter: "${subjectFilter || "none"}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const escapedFilter = subjectFilter
    ? escapeForAppleScript(subjectFilter)
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        set markedCount to 0
        set useFilter to ${subjectFilter ? "true" : "false"}
        set filterText to "${escapedFilter}"
        set allMsgs to messages of targetFolder
        set msgTotal to count of allMsgs

        repeat with i from 1 to msgTotal
          if markedCount >= ${limit} then exit repeat

          try
            set theMsg to item i of allMsgs
            if useFilter then
              set msgSubject to subject of theMsg
              if msgSubject contains filterText then
                set is read of theMsg to ${markAsRead}
                set markedCount to markedCount + 1
              end if
            else
              set is read of theMsg to ${markAsRead}
              set markedCount to markedCount + 1
            end if
          on error
            -- Skip problematic messages
          end try
        end repeat

        return markedCount as text
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[markEmailsRead] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const count = parseInt(result, 10) || 0;
    console.error(
      `[markEmailsRead] Marked ${count} emails as ${markAsRead ? "read" : "unread"}`,
    );
    return count;
  } catch (error) {
    console.error("[markEmailsRead] Error marking emails:", error);
    throw error;
  }
}

// Function to count emails in a folder
async function countEmails(
  folder: string = "Inbox",
  subjectFilter?: string,
): Promise<{ total: number; matching: number }> {
  console.error(
    `[countEmails] Counting emails in folder: ${folder}, filter: "${subjectFilter || "none"}"`,
  );
  await checkOutlookAccess();

  const escapedFilter = subjectFilter
    ? escapeForAppleScript(subjectFilter)
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        set allMsgs to messages of targetFolder
        set totalCount to count of allMsgs
        set matchCount to 0

        set useFilter to ${subjectFilter ? "true" : "false"}
        set filterText to "${escapedFilter}"

        if useFilter then
          repeat with i from 1 to totalCount
            try
              set theMsg to item i of allMsgs
              set msgSubject to subject of theMsg
              if msgSubject contains filterText then
                set matchCount to matchCount + 1
              end if
            on error
              -- Skip problematic messages
            end try
          end repeat
        else
          set matchCount to totalCount
        end if

        return (totalCount as text) & (ASCII character 31) & (matchCount as text)
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[countEmails] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const parts = result.split("\x1F");
    const total = parseInt(parts[0], 10) || 0;
    const matching = parseInt(parts[1], 10) || 0;

    console.error(`[countEmails] Total: ${total}, Matching: ${matching}`);
    return { total, matching };
  } catch (error) {
    console.error("[countEmails] Error counting emails:", error);
    throw error;
  }
}

// Function to get full details of a specific message by index
async function getMessageDetail(
  folder: string = "Inbox",
  messageIndex: number = 1,
): Promise<any> {
  console.error(
    `[getMessageDetail] Getting message detail from folder: ${folder}, index: ${messageIndex}`,
  );
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        set allMsgs to messages of targetFolder
        set msgTotal to count of allMsgs

        if ${messageIndex} > msgTotal then
          return "Error: Message index ${messageIndex} exceeds message count " & msgTotal
        end if

        set theMsg to item ${messageIndex} of allMsgs

        set msgSubject to subject of theMsg

        ${senderScript("theMsg", "msgSender")}

        set msgDate to time sent of theMsg

        try
          set msgContent to content of theMsg
          if (length of msgContent) > 5000 then
            set msgContent to text 1 thru 5000 of msgContent
          end if
        on error
          set msgContent to "(content unavailable)"
        end try

        set msgIsRead to is read of theMsg

        try
          set msgTodoFlag to todo flag of theMsg
        on error
          set msgTodoFlag to "none"
        end try

        set msgId to id of theMsg

        try
          set msgFolderName to name of targetFolder
        on error
          set msgFolderName to "${escapeForAppleScript(folder)}"
        end try

        set attachCount to 0
        set attachNames to ""
        try
          set attachList to attachments of theMsg
          set attachCount to count of attachList
          repeat with j from 1 to attachCount
            set attachName to name of item j of attachList
            if j > 1 then
              set attachNames to attachNames & ", " & attachName
            else
              set attachNames to attachName
            end if
          end repeat
        on error
          set attachCount to 0
          set attachNames to ""
        end try

        set sep to ASCII character 31
        return msgSubject & sep & msgSender & sep & msgDate & sep & msgContent & sep & (msgIsRead as text) & sep & (msgTodoFlag as text) & sep & (attachCount as text) & sep & attachNames & sep & (msgId as text) & sep & msgFolderName
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getMessageDetail] Raw result length: ${result.length}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const parts = result.split("\x1F");
    if (parts.length < 10) {
      throw new Error(
        `Unexpected result format: expected 10 fields, got ${parts.length}`,
      );
    }

    const detail = {
      subject: parts[0].trim(),
      sender: parts[1].trim(),
      dateSent: parts[2].trim(),
      content: parts[3].trim(),
      isRead: parts[4].trim() === "true",
      todoFlag: parts[5].trim(),
      attachmentCount: parseInt(parts[6].trim(), 10) || 0,
      attachmentNames: parts[7].trim() || "",
      id: parts[8].trim(),
      folder: parts[9].trim(),
    };

    console.error(
      `[getMessageDetail] Retrieved details for: "${detail.subject}"`,
    );
    return detail;
  } catch (error) {
    console.error("[getMessageDetail] Error getting message detail:", error);
    throw error;
  }
}

// Function to forward a specific email
async function forwardEmail(
  folder: string = "Inbox",
  messageIndex: number = 1,
  to: string,
  comment?: string,
): Promise<string> {
  console.error(
    `[forwardEmail] Forwarding message ${messageIndex} from "${folder}" to "${to}", comment: ${comment ? "yes" : "no"}`,
  );
  await checkOutlookAccess();

  const escapedTo = escapeForAppleScript(to);
  const escapedComment = comment
    ? escapeForAppleScript(comment).replace(/\n/g, "\\n")
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        set allMsgs to messages of targetFolder
        set msgTotal to count of allMsgs

        if ${messageIndex} > msgTotal then
          return "Error: Message index ${messageIndex} exceeds message count " & msgTotal
        end if

        set theMsg to item ${messageIndex} of allMsgs
        set msgSubject to subject of theMsg

        set fwdMsg to forward theMsg opening window no

        make new to recipient at fwdMsg with properties {email address:{address:"${escapedTo}"}}

        ${
          comment
            ? `
        try
          set existingContent to content of fwdMsg
          set content of fwdMsg to "${escapedComment}" & return & return & existingContent
        on error
          set content of fwdMsg to "${escapedComment}"
        end try
        `
            : ""
        }

        send fwdMsg

        return "Successfully forwarded \\"" & msgSubject & "\\" to ${escapedTo}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[forwardEmail] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    console.error(`[forwardEmail] ${result}`);
    return result;
  } catch (error) {
    console.error("[forwardEmail] Error forwarding email:", error);
    throw error;
  }
}

// Function to archive emails (move to Archive folder)
async function archiveEmails(
  folder: string = "Inbox",
  subjectFilter?: string,
  limit: number = 50,
): Promise<number> {
  console.error(
    `[archiveEmails] Archiving emails from "${folder}", filter: "${subjectFilter || "none"}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const escapedFilter = subjectFilter
    ? escapeForAppleScript(subjectFilter)
    : "";

  const script = `
    tell application "Microsoft Outlook"
      try
        ${folderResolverScript(folder, "targetFolder")}

        ${folderResolverScript("Archive", "archiveFolder")}

        set archivedCount to 0
        set useFilter to ${subjectFilter ? "true" : "false"}
        set filterText to "${escapedFilter}"

        repeat while archivedCount < ${limit}
          set msgList to messages of targetFolder
          if (count of msgList) is 0 then exit repeat

          set foundMatch to false
          set theMsg to item 1 of msgList

          if useFilter then
            set foundMatch to false
            repeat with i from 1 to (count of msgList)
              set candidateMsg to item i of msgList
              try
                set msgSubject to subject of candidateMsg
                if msgSubject contains filterText then
                  set theMsg to candidateMsg
                  set foundMatch to true
                  exit repeat
                end if
              end try
            end repeat
            if not foundMatch then exit repeat
          else
            set foundMatch to true
          end if

          if foundMatch then
            try
              move theMsg to archiveFolder
              set archivedCount to archivedCount + 1
            on error errMsg
              exit repeat
            end try
          end if
        end repeat

        return archivedCount as text
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[archiveEmails] Raw result: ${result}`);

    if (result.startsWith("Error:")) {
      throw new Error(result);
    }

    const count = parseInt(result, 10) || 0;
    console.error(`[archiveEmails] Archived ${count} emails`);
    return count;
  } catch (error) {
    console.error("[archiveEmails] Error archiving emails:", error);
    throw error;
  }
}

// ====================================================
// 5. CALENDAR FUNCTIONS
// ====================================================

// Function to get today's calendar events
async function getTodayEvents(limit: number = 10): Promise<any[]> {
  console.error(`[getTodayEvents] Getting today's events, limit: ${limit}`);
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      set todayEvents to {}
      -- Find the main calendar (first one named "Calendar" with events, falling back to first calendar)
      set allCalendars to every calendar
      set theCalendar to item 1 of allCalendars
      set bestCount to 0
      repeat with cal in allCalendars
        if name of cal is "Calendar" then
          set evtCount to count of calendar events of cal
          if evtCount > bestCount then
            set theCalendar to cal
            set bestCount to evtCount
          end if
        end if
      end repeat
      set todayDate to current date
      set startOfDay to todayDate - (time of todayDate)
      set dayEnd to startOfDay + 1 * days

      set eventList to calendar events of theCalendar whose start time is greater than or equal to startOfDay and start time is less than dayEnd

      set eventCount to count of eventList
      set limitCount to ${limit}

      if eventCount < limitCount then
        set limitCount to eventCount
      end if

      set eventResults to ""
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set evtSubject to subject of theEvent
        set evtStart to start time of theEvent
        set evtEndTime to end time of theEvent
        try
          set evtLocation to location of theEvent
        on error
          set evtLocation to ""
        end try
        set evtId to id of theEvent

        set eventLine to evtSubject & (ASCII character 31) & evtStart & (ASCII character 31) & evtEndTime & (ASCII character 31) & evtLocation & (ASCII character 31) & evtId
        if i > 1 then
          set eventResults to eventResults & (ASCII character 30) & eventLine
        else
          set eventResults to eventLine
        end if
      end repeat

      return eventResults
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getTodayEvents] Raw result length: ${result.length}`);

    // Parse the results
    const events: any[] = [];

    if (result && !result.startsWith("Error:")) {
      const eventLines = result.split("\x1E");
      for (const line of eventLines) {
        const parts = line.split("\x1F");
        if (parts.length >= 5) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim(),
          });
        }
      }
    }

    console.error(`[getTodayEvents] Found ${events.length} events for today`);
    return events;
  } catch (error) {
    console.error("[getTodayEvents] Error getting today's events:", error);
    throw error;
  }
}

// Function to get upcoming calendar events
async function getUpcomingEvents(
  days: number = 7,
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[getUpcomingEvents] Getting upcoming events for next ${days} days, limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      set upcomingEvents to {}
      -- Find the main calendar (first one named "Calendar" with events, falling back to first calendar)
      set allCalendars to every calendar
      set theCalendar to item 1 of allCalendars
      set bestCount to 0
      repeat with cal in allCalendars
        if name of cal is "Calendar" then
          set evtCount to count of calendar events of cal
          if evtCount > bestCount then
            set theCalendar to cal
            set bestCount to evtCount
          end if
        end if
      end repeat
      set todayDate to current date
      set startOfToday to todayDate - (time of todayDate)
      set dayEndDate to startOfToday + ${days} * days

      set eventList to calendar events of theCalendar whose start time is greater than or equal to todayDate and start time is less than dayEndDate

      set eventCount to count of eventList
      set limitCount to ${limit}

      if eventCount < limitCount then
        set limitCount to eventCount
      end if

      set eventResults to ""
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set evtSubject to subject of theEvent
        set evtStart to start time of theEvent
        set evtEndTime to end time of theEvent
        try
          set evtLocation to location of theEvent
        on error
          set evtLocation to ""
        end try
        set evtId to id of theEvent

        set eventLine to evtSubject & (ASCII character 31) & evtStart & (ASCII character 31) & evtEndTime & (ASCII character 31) & evtLocation & (ASCII character 31) & evtId
        if i > 1 then
          set eventResults to eventResults & (ASCII character 30) & eventLine
        else
          set eventResults to eventLine
        end if
      end repeat

      return eventResults
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getUpcomingEvents] Raw result length: ${result.length}`);

    // Parse the results
    const events: any[] = [];

    if (result && !result.startsWith("Error:")) {
      const eventLines = result.split("\x1E");
      for (const line of eventLines) {
        const parts = line.split("\x1F");
        if (parts.length >= 5) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim(),
          });
        }
      }
    }

    console.error(`[getUpcomingEvents] Found ${events.length} upcoming events`);
    return events;
  } catch (error) {
    console.error("[getUpcomingEvents] Error getting upcoming events:", error);
    throw error;
  }
}

// Function to search calendar events
async function searchEvents(
  searchTerm: string,
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[searchEvents] Searching for events with term: "${searchTerm}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
    tell application "Microsoft Outlook"
      set searchResults to {}
      -- Find the main calendar (first one named "Calendar" with events, falling back to first calendar)
      set allCalendars to every calendar
      set theCalendar to item 1 of allCalendars
      set bestCount to 0
      repeat with cal in allCalendars
        if name of cal is "Calendar" then
          set evtCount to count of calendar events of cal
          if evtCount > bestCount then
            set theCalendar to cal
            set bestCount to evtCount
          end if
        end if
      end repeat
      set allEvents to calendar events of theCalendar
      set i to 0
      set searchString to "${escapeForAppleScript(searchTerm)}"

      set eventResults to ""
      repeat with theEvent in allEvents
        if (subject of theEvent contains searchString) or (location of theEvent contains searchString) then
          set i to i + 1
          set evtSubject to subject of theEvent
          set evtStart to start time of theEvent
          set evtEndTime to end time of theEvent
          try
            set evtLocation to location of theEvent
          on error
            set evtLocation to ""
          end try
          set evtId to id of theEvent

          set eventLine to evtSubject & (ASCII character 31) & evtStart & (ASCII character 31) & evtEndTime & (ASCII character 31) & evtLocation & (ASCII character 31) & evtId
          if i > 1 then
            set eventResults to eventResults & (ASCII character 30) & eventLine
          else
            set eventResults to eventLine
          end if

          -- Stop if we've reached the limit
          if i >= ${limit} then
            exit repeat
          end if
        end if
      end repeat

      return eventResults
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[searchEvents] Raw result length: ${result.length}`);

    // Parse the results
    const events: any[] = [];

    if (result && !result.startsWith("Error:")) {
      const eventLines = result.split("\x1E");
      for (const line of eventLines) {
        const parts = line.split("\x1F");
        if (parts.length >= 5) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim(),
          });
        }
      }
    }

    console.error(`[searchEvents] Found ${events.length} matching events`);
    return events;
  } catch (error) {
    console.error("[searchEvents] Error searching events:", error);
    throw error;
  }
}

// Function to create a calendar event
async function createEvent(
  subject: string,
  start: string,
  end: string,
  location?: string,
  body?: string,
  attendees?: string,
): Promise<string> {
  console.error(
    `[createEvent] Creating event: "${subject}", start: ${start}, end: ${end}`,
  );
  await checkOutlookAccess();

  // Parse the ISO date strings to a format AppleScript can understand
  const startDate = new Date(start);
  const endDate = new Date(end);

  // Format for AppleScript (month/day/year hour:minute:second)
  const formattedStart = `date "${startDate.getMonth() + 1}/${startDate.getDate()}/${startDate.getFullYear()} ${startDate.getHours()}:${startDate.getMinutes()}:${startDate.getSeconds()}"`;
  const formattedEnd = `date "${endDate.getMonth() + 1}/${endDate.getDate()}/${endDate.getFullYear()} ${endDate.getHours()}:${endDate.getMinutes()}:${endDate.getSeconds()}"`;

  // Escape strings for AppleScript
  const escapedSubject = escapeForAppleScript(subject);
  const escapedLocation = location ? escapeForAppleScript(location) : "";
  const escapedBody = body ? escapeForAppleScript(body) : "";

  let script = `
    tell application "Microsoft Outlook"
      -- Find the main calendar (first one named "Calendar" with events, falling back to first calendar)
      set allCalendars to every calendar
      set theCalendar to item 1 of allCalendars
      set bestCount to 0
      repeat with cal in allCalendars
        if name of cal is "Calendar" then
          set evtCount to count of calendar events of cal
          if evtCount > bestCount then
            set theCalendar to cal
            set bestCount to evtCount
          end if
        end if
      end repeat
      set newEvent to make new calendar event at theCalendar with properties {subject:"${escapedSubject}", start time:${formattedStart}, end time:${formattedEnd}
  `;

  if (location) {
    script += `, location:"${escapedLocation}"`;
  }

  if (body) {
    script += `, content:"${escapedBody}"`;
  }

  script += `}
  `;

  // Add attendees if provided
  if (attendees) {
    const attendeeList = attendees.split(",").map((email) => email.trim());

    for (const attendee of attendeeList) {
      const escapedAttendee = escapeForAppleScript(attendee);
      script += `
        make new attendee at newEvent with properties {email address:"${escapedAttendee}"}
      `;
    }
  }

  script += `
      save newEvent
      return "Event created successfully"
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[createEvent] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[createEvent] Error creating event:", error);
    throw error;
  }
}

// ====================================================
// 6. CONTACTS FUNCTIONS
// ====================================================

// Function to list contacts with improved AppleScript syntax
async function listContacts(limit: number = 20): Promise<any[]> {
  console.error(`[listContacts] Listing contacts, limit: ${limit}`);
  await checkOutlookAccess();

  const script = `
      tell application "Microsoft Outlook"
        set contactList to {}
        set allContactsList to contacts
        set contactCount to count of allContactsList
        set limitCount to ${limit}

        if contactCount < limitCount then
          set limitCount to contactCount
        end if

        repeat with i from 1 to limitCount
          try
            set theContact to item i of allContactsList
            set contactName to full name of theContact

            -- Create a basic object with name
            set contactData to {name:contactName}

            -- Try to get email
            try
              set emailList to email addresses of theContact
              if (count of emailList) > 0 then
                set emailAddr to address of item 1 of emailList
                set contactData to contactData & {email:emailAddr}
              else
                set contactData to contactData & {email:"No email"}
              end if
            on error
              set contactData to contactData & {email:"No email"}
            end try

            -- Try to get phone
            try
              set phoneList to phones of theContact
              if (count of phoneList) > 0 then
                set phoneNum to formatted dial string of item 1 of phoneList
                set contactData to contactData & {phone:phoneNum}
              else
                set contactData to contactData & {phone:"No phone"}
              end if
            on error
              set contactData to contactData & {phone:"No phone"}
            end try

            set end of contactList to contactData
          on error
            -- Skip contacts that can't be processed
          end try
        end repeat

        return contactList
      end tell
    `;

  try {
    const result = await runAppleScript(script);
    console.error(`[listContacts] Raw result length: ${result.length}`);

    // Parse the results
    const contacts = [];
    const matches = result.match(/\{([^}]+)\}/g);

    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(",");
          const contact: any = {};

          props.forEach((prop) => {
            const parts = prop.split(":");
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(":").trim();
              contact[key] = value;
            }
          });

          if (contact.name) {
            contacts.push({
              name: contact.name,
              email: contact.email || "No email",
              phone: contact.phone || "No phone",
            });
          }
        } catch (parseError) {
          console.error(
            "[listContacts] Error parsing contact match:",
            parseError,
          );
        }
      }
    }

    console.error(`[listContacts] Found ${contacts.length} contacts`);
    return contacts;
  } catch (error) {
    console.error("[listContacts] Error listing contacts:", error);

    // Try an alternative approach using a simpler script
    try {
      const alternativeScript = `
          tell application "Microsoft Outlook"
            set contactList to {}
            set contactCount to count of contacts
            set limitCount to ${limit}

            if contactCount < limitCount then
              set limitCount to contactCount
            end if

            repeat with i from 1 to limitCount
              try
                set theContact to item i of contacts
                set contactName to full name of theContact
                set end of contactList to contactName
              end try
            end repeat

            return contactList
          end tell
        `;

      const result = await runAppleScript(alternativeScript);

      // Parse the simpler result format (just names)
      const simplifiedContacts = result.split(", ").map((name) => ({
        name: name,
        email: "Not available with simplified method",
        phone: "Not available with simplified method",
      }));

      console.error(
        `[listContacts] Found ${simplifiedContacts.length} contacts using alternative method`,
      );
      return simplifiedContacts;
    } catch (altError) {
      console.error("[listContacts] Alternative method also failed:", altError);
      throw new Error(
        `Error accessing contacts. The error might be related to Outlook permissions or configuration: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }
}

// Function to search contacts
// Function to search contacts with improved AppleScript syntax
async function searchContacts(
  searchTerm: string,
  limit: number = 10,
): Promise<any[]> {
  console.error(
    `[searchContacts] Searching for contacts with term: "${searchTerm}", limit: ${limit}`,
  );
  await checkOutlookAccess();

  const script = `
      tell application "Microsoft Outlook"
        set searchResults to {}
        set allContacts to contacts
        set i to 0
        set searchString to "${escapeForAppleScript(searchTerm)}"

        repeat with theContact in allContacts
          try
            set contactName to full name of theContact

            if contactName contains searchString then
              set i to i + 1

              -- Create basic contact info
              set contactData to {name:contactName}

              -- Try to get email
              try
                set emailList to email addresses of theContact
                if (count of emailList) > 0 then
                  set emailAddr to address of item 1 of emailList
                  set contactData to contactData & {email:emailAddr}
                else
                  set contactData to contactData & {email:"No email"}
                end if
              on error
                set contactData to contactData & {email:"No email"}
              end try

              -- Try to get phone
              try
                set phoneList to phones of theContact
                if (count of phoneList) > 0 then
                  set phoneNum to formatted dial string of item 1 of phoneList
                  set contactData to contactData & {phone:phoneNum}
                else
                  set contactData to contactData & {phone:"No phone"}
                end if
              on error
                set contactData to contactData & {phone:"No phone"}
              end try

              set end of searchResults to contactData

              -- Stop if we've reached the limit
              if i >= ${limit} then
                exit repeat
              end if
            end if
          on error
            -- Skip contacts that can't be processed
          end try
        end repeat

        return searchResults
      end tell
    `;

  try {
    const result = await runAppleScript(script);
    console.error(`[searchContacts] Raw result length: ${result.length}`);

    // Parse the results
    const contacts = [];
    const matches = result.match(/\{([^}]+)\}/g);

    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(",");
          const contact: any = {};

          props.forEach((prop) => {
            const parts = prop.split(":");
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(":").trim();
              contact[key] = value;
            }
          });

          if (contact.name) {
            contacts.push({
              name: contact.name,
              email: contact.email || "No email",
              phone: contact.phone || "No phone",
            });
          }
        } catch (parseError) {
          console.error(
            "[searchContacts] Error parsing contact match:",
            parseError,
          );
        }
      }
    }

    console.error(
      `[searchContacts] Found ${contacts.length} matching contacts`,
    );
    return contacts;
  } catch (error) {
    console.error("[searchContacts] Error searching contacts:", error);

    // Try an alternative approach with a simpler script that just returns names
    try {
      const alternativeScript = `
          tell application "Microsoft Outlook"
            set matchingContacts to {}
            set searchString to "${escapeForAppleScript(searchTerm)}"
            set i to 0

            repeat with theContact in contacts
              try
                set contactName to full name of theContact
                if contactName contains searchString then
                  set i to i + 1
                  set end of matchingContacts to contactName
                  if i >= ${limit} then exit repeat
                end if
              end try
            end repeat

            return matchingContacts
          end tell
        `;

      const result = await runAppleScript(alternativeScript);

      // Parse the simpler result format (just names)
      const simplifiedContacts = result.split(", ").map((name) => ({
        name: name,
        email: "Not available with simplified method",
        phone: "Not available with simplified method",
      }));

      console.error(
        `[searchContacts] Found ${simplifiedContacts.length} contacts using alternative method`,
      );
      return simplifiedContacts;
    } catch (altError) {
      console.error(
        "[searchContacts] Alternative method also failed:",
        altError,
      );
      throw new Error(
        `Error searching contacts. The error might be related to Outlook permissions or configuration: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }
}

// ====================================================
// 7. TYPE GUARDS
// ====================================================

// Type guards for arguments
function isMailArgs(args: unknown): args is {
  operation:
    | "unread"
    | "search"
    | "send"
    | "folders"
    | "read"
    | "delete"
    | "move"
    | "mark_read"
    | "count"
    | "get_message"
    | "forward"
    | "archive";
  folder?: string;
  limit?: number;
  searchTerm?: string;
  to?: string;
  subject?: string;
  body?: string;
  isHtml?: boolean;
  cc?: string;
  bcc?: string;
  attachments?: string[];
  subjectFilter?: string;
  destinationFolder?: string;
  messageIndex?: number;
  markAsRead?: boolean;
  comment?: string;
} {
  if (typeof args !== "object" || args === null) return false;

  const { operation } = args as any;

  if (
    !operation ||
    ![
      "unread",
      "search",
      "send",
      "folders",
      "read",
      "delete",
      "move",
      "mark_read",
      "count",
      "get_message",
      "forward",
      "archive",
    ].includes(operation)
  ) {
    return false;
  }

  // Check required fields based on operation
  switch (operation) {
    case "search":
      if (!(args as any).searchTerm) return false;
      break;
    case "send":
      if (!(args as any).to || !(args as any).subject || !(args as any).body)
        return false;
      break;
    case "move":
      if (!(args as any).folder || !(args as any).destinationFolder)
        return false;
      break;
    case "get_message":
      if (!(args as any).folder || !(args as any).messageIndex) return false;
      break;
    case "forward":
      if (
        !(args as any).folder ||
        !(args as any).messageIndex ||
        !(args as any).to
      )
        return false;
      break;
    case "delete":
    case "mark_read":
    case "count":
    case "archive":
      if (!(args as any).folder) return false;
      break;
  }

  return true;
}

function isCalendarArgs(args: unknown): args is {
  operation: "today" | "upcoming" | "search" | "create";
  searchTerm?: string;
  limit?: number;
  days?: number;
  subject?: string;
  start?: string;
  end?: string;
  location?: string;
  body?: string;
  attendees?: string;
} {
  if (typeof args !== "object" || args === null) return false;

  const { operation } = args as any;

  if (
    !operation ||
    !["today", "upcoming", "search", "create"].includes(operation)
  ) {
    return false;
  }

  // Check required fields based on operation
  switch (operation) {
    case "search":
      if (!(args as any).searchTerm) return false;
      break;
    case "create":
      if (!(args as any).subject || !(args as any).start || !(args as any).end)
        return false;
      break;
  }

  return true;
}

function isContactsArgs(args: unknown): args is {
  operation: "list" | "search";
  searchTerm?: string;
  limit?: number;
} {
  if (typeof args !== "object" || args === null) return false;

  const { operation } = args as any;

  if (!operation || !["list", "search"].includes(operation)) {
    return false;
  }

  // Check required fields based on operation
  if (operation === "search" && !(args as any).searchTerm) {
    return false;
  }

  return true;
}

// ====================================================
// 8. MCP REQUEST HANDLERS
// ====================================================

// Set up request handlers
server.setRequestHandler(ListToolsRequestSchema, async () => {
  console.error("[ListToolsRequest] Returning available tools");
  return {
    tools: [OUTLOOK_MAIL_TOOL, OUTLOOK_CALENDAR_TOOL, OUTLOOK_CONTACTS_TOOL],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;
    console.error(`[CallToolRequest] Received request for tool: ${name}`);

    if (!args) {
      throw new Error("No arguments provided");
    }

    switch (name) {
      case "outlook_mail": {
        if (!isMailArgs(args)) {
          throw new Error("Invalid arguments for outlook_mail tool");
        }

        const { operation } = args;
        console.error(`[CallToolRequest] Mail operation: ${operation}`);

        switch (operation) {
          case "unread": {
            const emails = await getUnreadEmails(args.folder, args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    emails.length > 0
                      ? `Found ${emails.length} unread email(s)${args.folder ? ` in folder "${args.folder}"` : ""}\n\n` +
                        emails
                          .map(
                            (email) =>
                              `[${email.dateSent}] From: ${email.sender}\nFolder: ${email.folder || "Inbox"}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? "..." : ""}`,
                          )
                          .join("\n\n")
                      : `No unread emails found${args.folder ? ` in folder "${args.folder}" or its subfolders` : ""}`,
                },
              ],
              isError: false,
            };
          }

          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const emails = await searchEmails(
              args.searchTerm,
              args.folder,
              args.limit,
            );
            return {
              content: [
                {
                  type: "text",
                  text:
                    emails.length > 0
                      ? `Found ${emails.length} email(s) for "${args.searchTerm}"${args.folder ? ` in folder "${args.folder}"` : ""}\n\n` +
                        emails
                          .map(
                            (email) =>
                              `[${email.dateSent}] From: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? "..." : ""}`,
                          )
                          .join("\n\n")
                      : `No emails found for "${args.searchTerm}"${args.folder ? ` in folder "${args.folder}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          // Update the handler in CallToolRequestSchema
          case "send": {
            if (!args.to || !args.subject || !args.body) {
              throw new Error(
                "Recipient (to), subject, and body are required for send operation",
              );
            }

            // Validate attachments if provided
            if (args.attachments && !Array.isArray(args.attachments)) {
              throw new Error("Attachments must be an array of file paths");
            }

            // Log attachment information for debugging
            console.error(
              `[CallTool] Send email with attachments: ${args.attachments ? JSON.stringify(args.attachments) : "none"}`,
            );

            const result = await sendEmail(
              args.to,
              args.subject,
              args.body,
              args.cc,
              args.bcc,
              args.isHtml || false,
              args.attachments,
            );

            return {
              content: [{ type: "text", text: result }],
              isError: false,
            };
          }

          case "folders": {
            const rawLines = await getMailFolders();
            // Each line is "  Name  ::  Full/Path" — format into a readable tree
            // with the path shown after the name for reference
            const treeLines = rawLines.map((line) => {
              const sepIdx = line.indexOf("  ::  ");
              if (sepIdx === -1) return line; // fallback line without path
              const display = line.substring(0, sepIdx);
              const fullPath = line.substring(sepIdx + 6);
              // Only show path if it differs from name (i.e. has depth)
              if (fullPath.includes("/")) {
                return `${display}  →  ${fullPath}`;
              }
              return display;
            });
            return {
              content: [
                {
                  type: "text",
                  text:
                    treeLines.length > 0
                      ? `Mail folder tree (${treeLines.length} folders):\n\nUse the full path (shown after →) as the 'folder' parameter for nested folders.\n\n${treeLines.join("\n")}`
                      : "No mail folders found. Make sure Outlook is running and properly configured.",
                },
              ],
              isError: false,
            };
          }

          case "read": {
            const emails = await readEmails(args.folder, args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    emails.length > 0
                      ? `Found ${emails.length} email(s)${args.folder ? ` in folder "${args.folder}"` : ""}\n\n` +
                        emails
                          .map(
                            (email) =>
                              `[${email.dateSent}] From: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? "..." : ""}`,
                          )
                          .join("\n\n")
                      : `No emails found${args.folder ? ` in folder "${args.folder}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          case "delete": {
            if (!args.folder) {
              throw new Error("Folder is required for delete operation");
            }
            const deletedCount = await deleteEmails(
              args.folder,
              args.subjectFilter,
              args.limit || 50,
            );
            return {
              content: [
                {
                  type: "text",
                  text: `Deleted ${deletedCount} email(s) from folder "${args.folder}"${args.subjectFilter ? ` matching subject filter "${args.subjectFilter}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          case "move": {
            if (!args.folder || !args.destinationFolder) {
              throw new Error(
                "Source folder and destination folder are required for move operation",
              );
            }
            const movedCount = await moveEmails(
              args.folder,
              args.destinationFolder,
              args.subjectFilter,
              args.limit || 50,
            );
            return {
              content: [
                {
                  type: "text",
                  text: `Moved ${movedCount} email(s) from "${args.folder}" to "${args.destinationFolder}"${args.subjectFilter ? ` matching subject filter "${args.subjectFilter}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          case "mark_read": {
            if (!args.folder) {
              throw new Error("Folder is required for mark_read operation");
            }
            const markRead = args.markAsRead !== false; // default true
            const markedCount = await markEmailsRead(
              args.folder,
              args.subjectFilter,
              args.limit || 50,
              markRead,
            );
            return {
              content: [
                {
                  type: "text",
                  text: `Marked ${markedCount} email(s) as ${markRead ? "read" : "unread"} in folder "${args.folder}"${args.subjectFilter ? ` matching subject filter "${args.subjectFilter}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          case "count": {
            if (!args.folder) {
              throw new Error("Folder is required for count operation");
            }
            const counts = await countEmails(args.folder, args.subjectFilter);
            return {
              content: [
                {
                  type: "text",
                  text: args.subjectFilter
                    ? `Folder "${args.folder}": ${counts.total} total email(s), ${counts.matching} matching subject filter "${args.subjectFilter}"`
                    : `Folder "${args.folder}": ${counts.total} email(s)`,
                },
              ],
              isError: false,
            };
          }

          case "get_message": {
            if (!args.folder || !args.messageIndex) {
              throw new Error(
                "Folder and messageIndex are required for get_message operation",
              );
            }
            const message = await getMessageDetail(
              args.folder,
              args.messageIndex,
            );
            return {
              content: [
                {
                  type: "text",
                  text:
                    `Message #${args.messageIndex} in "${args.folder}":\n\n` +
                    `Subject: ${message.subject}\n` +
                    `From: ${message.sender}\n` +
                    `Date: ${message.dateSent}\n` +
                    `Read: ${message.isRead}\n` +
                    `Flag: ${message.todoFlag}\n` +
                    `Attachments (${message.attachmentCount}): ${message.attachmentNames || "none"}\n` +
                    `Folder: ${message.folder}\n` +
                    `ID: ${message.id}\n\n` +
                    `--- Content ---\n${message.content}`,
                },
              ],
              isError: false,
            };
          }

          case "forward": {
            if (!args.folder || !args.messageIndex || !args.to) {
              throw new Error(
                "Folder, messageIndex, and to are required for forward operation",
              );
            }
            const fwdResult = await forwardEmail(
              args.folder,
              args.messageIndex,
              args.to,
              args.comment,
            );
            return {
              content: [{ type: "text", text: fwdResult }],
              isError: false,
            };
          }

          case "archive": {
            if (!args.folder) {
              throw new Error("Folder is required for archive operation");
            }
            const archivedCount = await archiveEmails(
              args.folder,
              args.subjectFilter,
              args.limit || 50,
            );
            return {
              content: [
                {
                  type: "text",
                  text: `Archived ${archivedCount} email(s) from folder "${args.folder}"${args.subjectFilter ? ` matching subject filter "${args.subjectFilter}"` : ""}`,
                },
              ],
              isError: false,
            };
          }

          default:
            throw new Error(`Unknown mail operation: ${operation}`);
        }
      }

      case "outlook_calendar": {
        if (!isCalendarArgs(args)) {
          throw new Error("Invalid arguments for outlook_calendar tool");
        }

        const { operation } = args;
        console.error(`[CallToolRequest] Calendar operation: ${operation}`);

        switch (operation) {
          case "today": {
            const events = await getTodayEvents(args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    events.length > 0
                      ? `Found ${events.length} event(s) for today:\n\n` +
                        events
                          .map(
                            (event) =>
                              `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`,
                          )
                          .join("\n\n")
                      : "No events found for today",
                },
              ],
              isError: false,
            };
          }

          case "upcoming": {
            const days = args.days || 7;
            const events = await getUpcomingEvents(days, args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    events.length > 0
                      ? `Found ${events.length} upcoming event(s) for the next ${days} days:\n\n` +
                        events
                          .map(
                            (event) =>
                              `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`,
                          )
                          .join("\n\n")
                      : `No upcoming events found for the next ${days} days`,
                },
              ],
              isError: false,
            };
          }

          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const events = await searchEvents(args.searchTerm, args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    events.length > 0
                      ? `Found ${events.length} event(s) matching "${args.searchTerm}":\n\n` +
                        events
                          .map(
                            (event) =>
                              `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`,
                          )
                          .join("\n\n")
                      : `No events found matching "${args.searchTerm}"`,
                },
              ],
              isError: false,
            };
          }

          case "create": {
            if (!args.subject || !args.start || !args.end) {
              throw new Error(
                "Subject, start time, and end time are required for create operation",
              );
            }
            const result = await createEvent(
              args.subject,
              args.start,
              args.end,
              args.location,
              args.body,
              args.attendees,
            );
            return {
              content: [{ type: "text", text: result }],
              isError: false,
            };
          }

          default:
            throw new Error(`Unknown calendar operation: ${operation}`);
        }
      }

      case "outlook_contacts": {
        if (!isContactsArgs(args)) {
          throw new Error("Invalid arguments for outlook_contacts tool");
        }

        const { operation } = args;
        console.error(`[CallToolRequest] Contacts operation: ${operation}`);

        switch (operation) {
          case "list": {
            const contacts = await listContacts(args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    contacts.length > 0
                      ? `Found ${contacts.length} contact(s):\n\n` +
                        contacts
                          .map(
                            (contact) =>
                              `Name: ${contact.name}\nEmail: ${contact.email}\nPhone: ${contact.phone}`,
                          )
                          .join("\n\n")
                      : "No contacts found",
                },
              ],
              isError: false,
            };
          }

          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const contacts = await searchContacts(args.searchTerm, args.limit);
            return {
              content: [
                {
                  type: "text",
                  text:
                    contacts.length > 0
                      ? `Found ${contacts.length} contact(s) matching "${args.searchTerm}":\n\n` +
                        contacts
                          .map(
                            (contact) =>
                              `Name: ${contact.name}\nEmail: ${contact.email}\nPhone: ${contact.phone}`,
                          )
                          .join("\n\n")
                      : `No contacts found matching "${args.searchTerm}"`,
                },
              ],
              isError: false,
            };
          }

          default:
            throw new Error(`Unknown contacts operation: ${operation}`);
        }
      }

      default:
        return {
          content: [{ type: "text", text: `Unknown tool: ${name}` }],
          isError: true,
        };
    }
  } catch (error) {
    console.error("[CallToolRequest] Error:", error);
    return {
      content: [
        {
          type: "text",
          text: `Error: ${error instanceof Error ? error.message : String(error)}`,
        },
      ],
      isError: true,
    };
  }
});

// ====================================================
// 9. START SERVER
// ====================================================

// Start the MCP server
console.error("Initializing Outlook MCP server transport...");
const transport = new StdioServerTransport();

(async () => {
  try {
    console.error("Connecting to transport...");
    await server.connect(transport);
    console.error("Outlook MCP Server running on stdio");
  } catch (error) {
    console.error("Failed to initialize MCP server:", error);
    process.exit(1);
  }
})();
