#!/usr/bin/env bun
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { runAppleScript } from 'run-applescript';
import os from 'os';
import path from 'path';

// ====================================================
// 1. Tool Definitions
// ====================================================

// Define Outlook Mail tool
const OUTLOOK_MAIL_TOOL: Tool = {
  name: "outlook_mail",
  description: "Interact with Microsoft Outlook for macOS - read, search, send, and manage emails",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description: "Operation to perform: 'unread', 'search', 'send', 'folders', or 'read'",
        enum: ["unread", "search", "send", "folders", "read"]
      },
      folder: {
        type: "string",
        description: "Email folder to use (optional - if not provided, uses inbox or searches across all folders)"
      },
      limit: {
        type: "number",
        description: "Number of emails to retrieve (optional, for unread, read, and search operations)"
      },
      searchTerm: {
        type: "string",
        description: "Text to search for in emails (required for search operation)"
      },
      to: {
        type: "string",
        description: "Recipient email address (required for send operation)"
      },
      subject: {
        type: "string",
        description: "Email subject (required for send operation)"
      },
      body: {
        type: "string",
        description: "Email body content (required for send operation)"
      },
      isHtml: {
        type: "boolean",
        description: "Whether the body content is HTML (optional for send operation, default: false)"
      },
      cc: {
        type: "string",
        description: "CC email address (optional for send operation)"
      },
      bcc: {
        type: "string",
        description: "BCC email address (optional for send operation)"
      },
      attachments: {
        type: "array",
        description: "File paths to attach to the email (optional for send operation)",
        items: {
          type: "string"
        }
      }
    },
    required: ["operation"]
  }
};

// Define Outlook Calendar tool
const OUTLOOK_CALENDAR_TOOL: Tool = {
  name: "outlook_calendar",
  description: "Interact with Microsoft Outlook for macOS calendar - view, create, and manage events",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description: "Operation to perform: 'today', 'upcoming', 'search', or 'create'",
        enum: ["today", "upcoming", "search", "create"]
      },
      searchTerm: {
        type: "string",
        description: "Text to search for in events (required for search operation)"
      },
      limit: {
        type: "number",
        description: "Number of events to retrieve (optional, for today and upcoming operations)"
      },
      days: {
        type: "number",
        description: "Number of days to look ahead (optional, for upcoming operation, default: 7)"
      },
      subject: {
        type: "string",
        description: "Event subject/title (required for create operation)"
      },
      start: {
        type: "string",
        description: "Start time in ISO format (required for create operation)"
      },
      end: {
        type: "string",
        description: "End time in ISO format (required for create operation)"
      },
      location: {
        type: "string",
        description: "Event location (optional for create operation)"
      },
      body: {
        type: "string",
        description: "Event description/body (optional for create operation)"
      },
      attendees: {
        type: "string",
        description: "Comma-separated list of attendee email addresses (optional for create operation)"
      }
    },
    required: ["operation"]
  }
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
        enum: ["list", "search"]
      },
      searchTerm: {
        type: "string",
        description: "Text to search for in contacts (required for search operation)"
      },
      limit: {
        type: "number",
        description: "Number of contacts to retrieve (optional)"
      }
    },
    required: ["operation"]
  }
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
  }
);

// ====================================================
// 3. Core Functions
// ====================================================

// Escape a string for safe interpolation inside an AppleScript double-quoted literal.
// Backslashes must be escaped before quotes, otherwise a trailing backslash can
// neutralise the quote escape and allow injection (e.g. the sequence \\" becomes
// a literal backslash followed by an unescaped quote that terminates the string).
function escapeAppleScriptString(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

// Resolve an attachment path and verify it is inside the user's home directory.
// Rejects path-traversal attempts (e.g. ../../etc/passwd) and absolute paths
// outside the home tree.
function safeAttachmentPath(filePath: string): string {
  const homeDir = os.homedir();
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(homeDir + path.sep) && resolved !== homeDir) {
    throw new Error(`Attachment path is outside the home directory and was rejected: ${resolved}`);
  }
  return resolved;
}

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
      console.error("[checkOutlookAccess] Microsoft Outlook is not installed or running");
      throw new Error("Microsoft Outlook is not installed or running on this system");
    }
    
    const isRunning = await runAppleScript(`
      tell application "System Events"
        set outlookRunning to application process "Microsoft Outlook" exists
        return outlookRunning
      end tell
    `);

    if (isRunning !== "true") {
      console.error("[checkOutlookAccess] Microsoft Outlook is not running, attempting to launch...");
      try {
        await runAppleScript(`
          tell application "Microsoft Outlook" to activate
          delay 2
        `);
        console.error("[checkOutlookAccess] Launched Outlook successfully");
      } catch (activateError) {
        console.error("[checkOutlookAccess] Error activating Microsoft Outlook:", activateError);
        throw new Error("Could not activate Microsoft Outlook. Please start it manually.");
      }
    } else {
      console.error("[checkOutlookAccess] Microsoft Outlook is already running");
    }
    
    return true;
  } catch (error) {
    console.error("[checkOutlookAccess] Outlook access check failed:", error);
    throw new Error(
      `Cannot access Microsoft Outlook. Please make sure Outlook is installed and properly configured. Error: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}

// ====================================================
// 4. EMAIL FUNCTIONS
// ====================================================

// Function to get unread emails
async function getUnreadEmails(folder: string = "Inbox", limit: number = 10): Promise<any[]> {
  console.error(`[getUnreadEmails] Getting unread emails from folder: ${folder}, limit: ${limit}`);
  await checkOutlookAccess();
  
  const escapedFolderName = escapeAppleScriptString(folder);
  const script = `
    tell application "Microsoft Outlook"
      try
        -- Look up folder by name to avoid unquoted identifier injection
        set targetFolder to null
        set allFolders to mail folders
        repeat with mailFolder in allFolders
          if name of mailFolder is "${escapedFolderName}" then
            set targetFolder to mailFolder
            exit repeat
          end if
        end repeat
        if targetFolder is null then set targetFolder to inbox
        set theFolder to targetFolder
        set unreadMessages to {}
        set allMessages to messages of theFolder
        set i to 0
        
        repeat with theMessage in allMessages
          if read status of theMessage is false then
            set i to i + 1
            set msgData to {subject:subject of theMessage, sender:sender of theMessage, ¬
                       date:time sent of theMessage, id:id of theMessage}
            
            -- Try to get content
            try
              set msgContent to content of theMessage
              if length of msgContent > 500 then
                set msgContent to (text 1 thru 500 of msgContent) & "..."
              end if
              set msgData to msgData & {content:msgContent}
            on error
              set msgData to msgData & {content:"[Content not available]"}
            end try
            
            set end of unreadMessages to msgData
            
            -- Stop if we've reached the limit
            if i >= ${limit} then
              exit repeat
            end if
          end if
        end repeat
        
        return unreadMessages
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[getUnreadEmails] Raw result length: ${result.length}`);
    
    // Parse the results (AppleScript returns records as text)
    if (result.startsWith("Error:")) {
      throw new Error(result);
    }
    
    // Simple parsing for demonstration
    // In a production environment, you'd want more robust parsing
    const emails = [];
    const matches = result.match(/\{([^}]+)\}/g);
    
    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(',');
          const email: any = {};
          
          props.forEach(prop => {
            const parts = prop.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(':').trim();
              email[key] = value;
            }
          });
          
          if (email.subject || email.sender) {
            emails.push({
              subject: email.subject || "No subject",
              sender: email.sender || "Unknown sender",
              dateSent: email.date || new Date().toString(),
              content: email.content || "[Content not available]",
              id: email.id || ""
            });
          }
        } catch (parseError) {
          console.error('[getUnreadEmails] Error parsing email match:', parseError);
        }
      }
    }
    
    console.error(`[getUnreadEmails] Found ${emails.length} unread emails`);
    return emails;
  } catch (error) {
    console.error("[getUnreadEmails] Error getting unread emails:", error);
    throw error;
  }
}

// Function to search emails
async function searchEmails(searchTerm: string, folder: string = "Inbox", limit: number = 10): Promise<any[]> {
  console.error(`[searchEmails] Searching for "${searchTerm}" in folder: ${folder}, limit: ${limit}`);
  await checkOutlookAccess();
  
  const escapedFolderName = escapeAppleScriptString(folder);
  const script = `
    tell application "Microsoft Outlook"
      try
        -- Look up folder by name to avoid unquoted identifier injection
        set targetFolder to null
        set allFolders to mail folders
        repeat with mailFolder in allFolders
          if name of mailFolder is "${escapedFolderName}" then
            set targetFolder to mailFolder
            exit repeat
          end if
        end repeat
        if targetFolder is null then set targetFolder to inbox
        set theFolder to targetFolder
        set searchResults to {}
        set allMessages to messages of theFolder
        set i to 0
        set searchString to "${escapeAppleScriptString(searchTerm)}"
        
        repeat with theMessage in allMessages
          if (subject of theMessage contains searchString) or (content of theMessage contains searchString) then
            set i to i + 1
            set msgData to {subject:subject of theMessage, sender:sender of theMessage, ¬
                       date:time sent of theMessage, id:id of theMessage}
            
            -- Try to get content
            try
              set msgContent to content of theMessage
              if length of msgContent > 500 then
                set msgContent to (text 1 thru 500 of msgContent) & "..."
              end if
              set msgData to msgData & {content:msgContent}
            on error
              set msgData to msgData & {content:"[Content not available]"}
            end try
            
            set end of searchResults to msgData
            
            -- Stop if we've reached the limit
            if i >= ${limit} then
              exit repeat
            end if
          end if
        end repeat
        
        return searchResults
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[searchEmails] Raw result length: ${result.length}`);
    
    // Parse the results
    if (result.startsWith("Error:")) {
      throw new Error(result);
    }
    
    // Parse the emails similar to unread emails
    const emails = [];
    const matches = result.match(/\{([^}]+)\}/g);
    
    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(',');
          const email: any = {};
          
          props.forEach(prop => {
            const parts = prop.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(':').trim();
              email[key] = value;
            }
          });
          
          if (email.subject || email.sender) {
            emails.push({
              subject: email.subject || "No subject",
              sender: email.sender || "Unknown sender",
              dateSent: email.date || new Date().toString(),
              content: email.content || "[Content not available]",
              id: email.id || ""
            });
          }
        } catch (parseError) {
          console.error('[searchEmails] Error parsing email match:', parseError);
        }
      }
    }
    
    console.error(`[searchEmails] Found ${emails.length} matching emails`);
    return emails;
  } catch (error) {
    console.error("[searchEmails] Error searching emails:", error);
    throw error;
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
  attachments?: string[]
): Promise<string> {
  console.error(`[sendEmail] Sending email to: ${to}, subject: "${subject}"`);
  console.error(`[sendEmail] Attachments: ${attachments ? JSON.stringify(attachments) : 'none'}`);
  
  await checkOutlookAccess();

  // Extract name from email if possible (for display name)
  const extractNameFromEmail = (email: string): string => {
    const namePart = email.split('@')[0];
    return namePart
      .split('.')
      .map(part => part.charAt(0).toUpperCase() + part.slice(1))
      .join(' ');
  };

  // Get name for display
  const toName = extractNameFromEmail(to);
  const ccName = cc ? extractNameFromEmail(cc) : "";
  const bccName = bcc ? extractNameFromEmail(bcc) : "";

  // Escape special characters
  const escapedSubject = escapeAppleScriptString(subject);
  const escapedBody = escapeAppleScriptString(body).replace(/\n/g, '\\n');

  // Process attachments: resolve paths and reject anything outside home directory
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    processedAttachments = attachments.map(p => safeAttachmentPath(p));
    console.error(`[sendEmail] Processed attachments: ${JSON.stringify(processedAttachments)}`);
  }
  
  // Create attachment script part with better error handling
  const attachmentScript = processedAttachments.length > 0
    ? processedAttachments.map(filePath => {
      const escapedPath = escapeAppleScriptString(filePath);
      return `
        try
          set attachmentFile to POSIX file "${escapedPath}"
          make new attachment at msg with properties {file:attachmentFile}
          log "Successfully attached file: ${escapedPath}"
        on error errMsg
          log "Failed to attach file: ${escapedPath} - Error: " & errMsg
        end try
      `;
    }).join('\n')
    : '';

  // Try approach 1: Using specific syntax for creating a message with attachments
  try {
    const script1 = `
      tell application "Microsoft Outlook"
        try
          set msg to make new outgoing message with properties {subject:"${escapedSubject}"}
          
          ${isHtml ? 
            `set content type of msg to HTML
             set content of msg to "${escapedBody}"` 
          : 
            `set content of msg to "${escapedBody}"`
          }
          
          tell msg
            set recipTo to make new to recipient with properties {email address:{name:"${escapeAppleScriptString(toName)}", address:"${escapeAppleScriptString(to)}"}}
            ${cc ? `set recipCc to make new cc recipient with properties {email address:{name:"${escapeAppleScriptString(ccName)}", address:"${escapeAppleScriptString(cc)}"}}` : ''}
            ${bcc ? `set recipBcc to make new bcc recipient with properties {email address:{name:"${escapeAppleScriptString(bccName)}", address:"${escapeAppleScriptString(bcc)}"}}` : ''}
            
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
            
            ${isHtml ? 
              `set content type of theMessage to HTML
               set content of theMessage to "${escapedBody}"` 
            : 
              `set content of theMessage to "${escapedBody}"`
            }
            
            set to recipients of theMessage to {"${escapeAppleScriptString(to)}"}
            ${cc ? `set cc recipients of theMessage to {"${escapeAppleScriptString(cc)}"}` : ''}
            ${bcc ? `set bcc recipients of theMessage to {"${escapeAppleScriptString(bcc)}"}` : ''}

            ${processedAttachments.map(filePath => {
              const escapedPath = escapeAppleScriptString(filePath);
              return `
                try
                  set attachmentFile to POSIX file "${escapedPath}"
                  make new attachment at theMessage with properties {file:attachmentFile}
                  log "Successfully attached file: ${escapedPath}"
                on error attachErrMsg
                  log "Failed to attach file: ${escapedPath} - Error: " & attachErrMsg
                end try
              `;
            }).join('\n')}

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
              
              ${isHtml ? 
                `set content type of newMessage to HTML
                 set content of newMessage to "${escapedBody}"` 
              : 
                `set content of newMessage to "${escapedBody}"`
              }
              
              set to recipients of newMessage to {"${escapeAppleScriptString(to)}"}
              ${cc ? `set cc recipients of newMessage to {"${escapeAppleScriptString(cc)}"}` : ''}
              ${bcc ? `set bcc recipients of newMessage to {"${escapeAppleScriptString(bcc)}"}` : ''}

              ${processedAttachments.map(filePath => {
                const escapedPath = escapeAppleScriptString(filePath);
                return `
                  try
                    set attachmentFile to POSIX file "${escapedPath}"
                    make new attachment at newMessage with properties {file:attachmentFile}
                    log "Successfully attached file: ${escapedPath}"
                  on error attachErrMsg
                    log "Failed to attach file: ${escapedPath} - Error: " & attachErrMsg
                  end try
                `;
              }).join('\n')}
              
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
        throw new Error(`Could not send or create email. Please check if Outlook is properly configured and that you have granted necessary permissions. Error details: ${error3}`);
      }
    }
  }
}
// Function to get mail folders - this works based on your logs
async function getMailFolders(): Promise<string[]> {
    console.error("[getMailFolders] Getting mail folders");
    await checkOutlookAccess();
  
    const script = `
      tell application "Microsoft Outlook"
        set folderNames to {}
        set allFolders to mail folders
        
        repeat with theFolder in allFolders
          set end of folderNames to name of theFolder
        end repeat
        
        return folderNames
      end tell
    `;
  
    try {
      const result = await runAppleScript(script);
      console.error(`[getMailFolders] Result: ${result}`);
      return result.split(", ");
    } catch (error) {
      console.error("[getMailFolders] Error getting mail folders:", error);
      throw error;
    }
  }
  
  // Function to read emails in a folder that uses simple AppleScript
async function readEmails(folder: string = "Inbox", limit: number = 10): Promise<any[]> {
    console.error(`[readEmails] Reading emails from folder: ${folder}, limit: ${limit}`);
    await checkOutlookAccess();
    
    // Use a simplified approach that should be more compatible
    const escapedFolderName = escapeAppleScriptString(folder);
    const script = `
      tell application "Microsoft Outlook"
        try
          -- Get the folder by name safely
          set targetFolder to null
          set allFolders to mail folders
          repeat with mailFolder in allFolders
            if name of mailFolder is "${escapedFolderName}" then
              set targetFolder to mailFolder
              exit repeat
            end if
          end repeat
          
          if targetFolder is null then set targetFolder to inbox
          
          -- Get messages
          set messageList to {}
          set msgCount to 0
          set allMsgs to messages of targetFolder
          
          repeat with i from 1 to (count of allMsgs)
            if msgCount >= ${limit} then exit repeat
            
            try
              set theMsg to item i of allMsgs
              set msgSubject to subject of theMsg
              set msgSender to sender of theMsg
              set msgDate to time sent of theMsg
              
              -- Create a simple text representation for the message
              set msgInfo to msgSubject & " | " & msgSender & " | " & msgDate
              set end of messageList to msgInfo
              set msgCount to msgCount + 1
            on error
              -- Skip problematic messages
            end try
          end repeat
          
          return messageList
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
      
      // Parse the results in a simple format
      const emails = result.split(", ").map(msgInfo => {
        const parts = msgInfo.split(" | ");
        return {
          subject: parts[0] || "No subject",
          sender: parts[1] || "Unknown sender",
          dateSent: parts[2] || new Date().toString(),
          content: "Content not retrieved in simple mode"
        };
      });
      
      console.error(`[readEmails] Found ${emails.length} emails using simplified approach`);
      return emails;
    } catch (error) {
      console.error("[readEmails] Error reading emails:", error);
      throw error;
    }
  }

// ====================================================
// 5. CALENDAR FUNCTIONS
// ====================================================

// Convert an AppleScript date string like "Tuesday, May 5, 2026 at 2:00:00 PM"
// to a comparable minutes-since-midnight value (for same-day ordering only).
function parseStartMinutes(start: string): number {
  const m = start.match(/at (\d+):(\d+):\d+ (AM|PM)/i);
  if (!m) return 0;
  let h = parseInt(m[1]);
  const min = parseInt(m[2]);
  const meridiem = m[3].toUpperCase();
  if (meridiem === 'PM' && h !== 12) h += 12;
  if (meridiem === 'AM' && h === 12) h = 0;
  return h * 60 + min;
}

// Post-process calendar event lists:
// 1. If "Declined: X" exists on the same day, suppress the non-declined "X"
//    (handles the case where Outlook materializes both the master and the declined exception)
// 2. If the same subject appears more than once on the same day, keep only the
//    earliest start time (Outlook sometimes stores both original and modified occurrences)
function deduplicateCalendarEvents(events: any[]): any[] {
  // Build set of declined keys: "<baseSubject>|<datePrefix>"
  const declinedKeys = new Set<string>();
  for (const e of events) {
    if (e.subject.startsWith('Declined: ')) {
      const base = e.subject.slice('Declined: '.length);
      const day = e.start.split(' at ')[0]?.trim() ?? e.start;
      declinedKeys.add(`${base}|${day}`);
    }
  }

  // Build map of subject+day → event with earliest start time
  const earliest = new Map<string, any>();
  for (const e of events) {
    // Don't deduplicate declined/canceled markers against each other
    const day = e.start.split(' at ')[0]?.trim() ?? e.start;
    const key = `${e.subject}|${day}`;
    if (!earliest.has(key)) {
      earliest.set(key, e);
    } else {
      const cur = earliest.get(key);
      if (parseStartMinutes(e.start) < parseStartMinutes(cur.start)) {
        earliest.set(key, e);
      }
    }
  }

  return events.filter(e => {
    // Suppress plain "X" when "Declined: X" exists on the same day
    if (!e.subject.startsWith('Declined: ') && !e.subject.startsWith('Canceled: ')) {
      const day = e.start.split(' at ')[0]?.trim() ?? e.start;
      if (declinedKeys.has(`${e.subject}|${day}`)) {
        return false;
      }
    }
    // Keep only the earliest occurrence for a given subject+day
    const day = e.start.split(' at ')[0]?.trim() ?? e.start;
    const key = `${e.subject}|${day}`;
    return earliest.get(key) === e;
  });
}

// Function to get today's calendar events
async function getTodayEvents(limit: number = 10): Promise<any[]> {
  console.error(`[getTodayEvents] Getting today's events, limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set todayDate to current date
      set startOfDay to todayDate - (time of todayDate)
      set endOfDay to startOfDay + 1 * days

      set eventList to every calendar event whose start time is greater than or equal to startOfDay and start time is less than endOfDay

      set eventCount to count of eventList
      set limitCount to ${limit}
      if eventCount < limitCount then
        set limitCount to eventCount
      end if

      set resultStr to ""
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set evtSubject to subject of theEvent
        set evtStart to (start time of theEvent) as string
        set evtEnd to (end time of theEvent) as string
        -- For recurring masters whose first occurrence falls on today (new series),
        -- read the actual occurrence time. A specific occurrence may have been rescheduled
        -- from its original slot. Only apply to non-declined/canceled events to avoid
        -- calling get occurrence of on exception objects.
        set phaseOneInclude to true
        if is recurring of theEvent and not (evtSubject begins with "Declined:") and not (evtSubject begins with "Canceled:") then
          try
            set occActual to get occurrence of theEvent at (start time of theEvent)
            if occActual is not missing value then
              set occActualStart to start time of occActual
              if occActualStart >= startOfDay and occActualStart < endOfDay then
                set evtStart to occActualStart as string
                set evtEnd to (end time of occActual) as string
                set evtSubject to subject of occActual
              else
                -- Rescheduled to a different day; skip from Phase 1
                set phaseOneInclude to false
              end if
            end if
          on error
            -- Occurrence was deleted by organizer; do not show it
            set phaseOneInclude to false
          end try
        end if
        if phaseOneInclude then
          set evtLoc to location of theEvent
          if evtLoc is missing value then set evtLoc to ""
          set evtId to (id of theEvent) as string

          set evtAttendees to ""
          try
            set attList to attendees of theEvent
            set attCount to count of attList
            repeat with j from 1 to attCount
              try
                set thisAtt to item j of attList
                set addrRec to email address of thisAtt
                set rawEmail to address of addrRec
                set displayName to name of addrRec
                set attEntry to rawEmail
                if displayName is not "" and displayName is not missing value then
                  if displayName contains rawEmail then
                    set attEntry to displayName
                  else
                    set attEntry to displayName & " <" & rawEmail & ">"
                  end if
                end if
                if evtAttendees is "" then
                  set evtAttendees to attEntry
                else
                  set evtAttendees to evtAttendees & ";" & attEntry
                end if
              on error
              end try
            end repeat
          on error
          end try

          set evtRec to evtSubject & "|||" & evtStart & "|||" & evtEnd & "|||" & evtLoc & "|||" & evtId & "|||" & evtAttendees
          if resultStr is "" then
            set resultStr to evtRec
          else
            set resultStr to resultStr & "###" & evtRec
          end if
        end if -- phaseOneInclude
      end repeat

      -- Phase 2: recurring master events that generate an occurrence today.
      -- Outlook only stores the master occurrence for self-created recurring events;
      -- individual instances are not materialized until modified. We compute today's
      -- occurrences from the recurrence rule.
      -- Only synthesize for self-organized events: externally-organized recurring
      -- meetings appear in Phase 1 when active; absence from Phase 1 = canceled.
      -- Self-created events have organizer = missing value; externally-organized
      -- events have organizer set to the other person's record.
      set allRecurring to every calendar event whose is recurring is true
      set recurCount to count of allRecurring
      repeat with rk from 1 to recurCount
        set masterEvent to item rk of allRecurring
        set masterStart to start time of masterEvent

        -- Skip if the master itself falls on today (already captured in Phase 1)
        if masterStart < startOfDay or masterStart >= endOfDay then
          set recData to recurrence of masterEvent
          set recStartDate to start date of recData
          set recStartDay to recStartDate - (time of recStartDate)

          if startOfDay >= recStartDay then
            -- Check that today is before or on the recurrence end date (if any)
            set inRange to true
            try
              set endInfo to end date of recData
              if (end type of endInfo) is end date type then
                set recEndDay to (data of endInfo) - (time of (data of endInfo))
                if startOfDay > recEndDay then
                  set inRange to false
                end if
              end if
            end try

            if inRange then
              -- No weekday/interval pre-filter: use get occurrence of to handle
              -- multi-day recurrence patterns (e.g. Mon+Thu) correctly.
              -- Check if already captured in Phase 1 (exact subject, Canceled:, or Declined: prefix)
              set alreadyCaptured to false
              set masterSubject to subject of masterEvent
              repeat with k from 1 to eventCount
                set kSubject to subject of (item k of eventList)
                if kSubject = masterSubject or kSubject = "Canceled: " & masterSubject or kSubject = "Declined: " & masterSubject then
                  set alreadyCaptured to true
                  exit repeat
                end if
              end repeat

              if not alreadyCaptured then
                -- Verify the occurrence actually exists and is not declined/deleted.
                -- get occurrence of throws an error if the series was ended or the
                -- specific occurrence was deleted by the organizer.
                set occExists to false
                try
                  set masterTimeOfDay to time of masterStart
                  set checkDateTime to startOfDay + masterTimeOfDay
                  set testOcc to get occurrence of masterEvent at checkDateTime
                  if testOcc is not missing value then
                    if not (subject of testOcc begins with "Declined:") then
                      set occExists to true
                    end if
                  end if
                on error
                  set occExists to false
                end try

                if occExists then
                -- Use actual occurrence times (handles rescheduled occurrences)
                set evtStartToday to start time of testOcc
                set evtEndToday to end time of testOcc
                if evtStartToday >= startOfDay and evtStartToday < endOfDay then
                set evtSubject to subject of masterEvent
                set evtStart to evtStartToday as string
                set evtEnd to evtEndToday as string
                set evtLoc to location of masterEvent
                if evtLoc is missing value then set evtLoc to ""
                set evtId to (id of masterEvent) as string

                set evtAttendees to ""
                try
                  set attList to attendees of masterEvent
                  set attCount to count of attList
                  repeat with j from 1 to attCount
                    try
                      set thisAtt to item j of attList
                      set addrRec to email address of thisAtt
                      set rawEmail to address of addrRec
                      set displayName to name of addrRec
                      set attEntry to rawEmail
                      if displayName is not "" and displayName is not missing value then
                        if displayName contains rawEmail then
                          set attEntry to displayName
                        else
                          set attEntry to displayName & " <" & rawEmail & ">"
                        end if
                      end if
                      if evtAttendees is "" then
                        set evtAttendees to attEntry
                      else
                        set evtAttendees to evtAttendees & ";" & attEntry
                      end if
                    on error
                    end try
                  end repeat
                on error
                end try

                set evtRec to evtSubject & "|||" & evtStart & "|||" & evtEnd & "|||" & evtLoc & "|||" & evtId & "|||" & evtAttendees
                if resultStr is "" then
                  set resultStr to evtRec
                else
                  set resultStr to resultStr & "###" & evtRec
                end if
                end if -- occurrence still today
                end if -- occExists
              end if -- not alreadyCaptured
            end if -- inRange
          end if -- startOfDay >= recStartDay
        end if -- masterStart outside today
      end repeat

      return resultStr
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getTodayEvents] Raw result length: ${result.length}`);

    if (!result || result.trim() === '') {
      console.error('[getTodayEvents] No events found for today');
      return [];
    }

    const events = [];
    const records = result.split('###');
    for (const record of records) {
      try {
        const parts = record.split('|||');
        if (parts.length >= 5 && parts[0].trim()) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim(),
            attendees: parts[5]
              ? parts[5].split(';').filter((e: string) => e && e.trim())
              : []
          });
        }
      } catch (parseError) {
        console.error('[getTodayEvents] Error parsing event record:', parseError);
      }
    }

    const deduped = deduplicateCalendarEvents(events);
    console.error(`[getTodayEvents] Found ${deduped.length} events for today (${events.length} before dedup)`);
    return deduped;
  } catch (error) {
    console.error("[getTodayEvents] Error getting today's events:", error);
    throw error;
  }
}

// Function to get upcoming calendar events
async function getUpcomingEvents(days: number = 7, limit: number = 10): Promise<any[]> {
  console.error(`[getUpcomingEvents] Getting upcoming events for next ${days} days, limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set todayDate to current date
      set startOfToday to todayDate - (time of todayDate)
      set endDate to startOfToday + ${days} * days

      set eventList to every calendar event whose start time is greater than or equal to todayDate and start time is less than endDate

      set eventCount to count of eventList
      set limitCount to ${limit}
      if eventCount < limitCount then
        set limitCount to eventCount
      end if

      set resultStr to ""
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set evtSubject to subject of theEvent
        set evtStart to (start time of theEvent) as string
        set evtEnd to (end time of theEvent) as string
        -- For recurring masters whose first occurrence falls in the window (new series),
        -- read the actual occurrence time. A specific occurrence may have been rescheduled
        -- from its original slot. Only apply to non-declined/canceled events.
        set phaseOneInclude to true
        if is recurring of theEvent and not (evtSubject begins with "Declined:") and not (evtSubject begins with "Canceled:") then
          try
            set occActual to get occurrence of theEvent at (start time of theEvent)
            if occActual is not missing value then
              set occActualStart to start time of occActual
              if occActualStart >= todayDate and occActualStart < endDate then
                set evtStart to occActualStart as string
                set evtEnd to (end time of occActual) as string
                set evtSubject to subject of occActual
              else
                -- Rescheduled outside the window; skip from Phase 1
                set phaseOneInclude to false
              end if
            end if
          on error
            -- Occurrence was deleted by organizer; do not show it
            set phaseOneInclude to false
          end try
        end if
        if phaseOneInclude then
          set evtLoc to location of theEvent
          if evtLoc is missing value then set evtLoc to ""
          set evtId to (id of theEvent) as string

          set evtAttendees to ""
          try
            set attList to attendees of theEvent
            set attCount to count of attList
            repeat with j from 1 to attCount
              try
                set thisAtt to item j of attList
                set addrRec to email address of thisAtt
                set rawEmail to address of addrRec
                set displayName to name of addrRec
                set attEntry to rawEmail
                if displayName is not "" and displayName is not missing value then
                  if displayName contains rawEmail then
                    set attEntry to displayName
                  else
                    set attEntry to displayName & " <" & rawEmail & ">"
                  end if
                end if
                if evtAttendees is "" then
                  set evtAttendees to attEntry
                else
                  set evtAttendees to evtAttendees & ";" & attEntry
                end if
              on error
              end try
            end repeat
          on error
          end try

          set evtRec to evtSubject & "|||" & evtStart & "|||" & evtEnd & "|||" & evtLoc & "|||" & evtId & "|||" & evtAttendees
          if resultStr is "" then
            set resultStr to evtRec
          else
            set resultStr to resultStr & "###" & evtRec
          end if
        end if -- phaseOneInclude
      end repeat

      -- Phase 2: recurring master events that generate occurrences within the window
      -- Only synthesize for self-organized events; externally-organized events
      -- appear in Phase 1 when active, so their absence means they were canceled.
      -- Self-created events have organizer = missing value; externally-organized
      -- events have organizer set to the creator's record.
      set allRecurring to every calendar event whose is recurring is true
      set recurCount to count of allRecurring
      repeat with rk from 1 to recurCount
        set masterEvent to item rk of allRecurring
        set masterStart to start time of masterEvent
        -- Skip if the master itself falls within the window (already in Phase 1)
        if masterStart < todayDate or masterStart >= endDate then
          set recData to recurrence of masterEvent
          set recStartDate to start date of recData
          set recStartDay to recStartDate - (time of recStartDate)

          -- For each day in the window, check if this master generates an occurrence
          set checkDay to startOfToday
          repeat while checkDay < endDate
            if checkDay >= recStartDay then
              set inRange to true
              try
                set endInfo to end date of recData
                if (end type of endInfo) is end date type then
                  set recEndDay to (data of endInfo) - (time of (data of endInfo))
                  if checkDay > recEndDay then
                    set inRange to false
                  end if
                end if
              end try

              if inRange then
                -- No weekday/interval pre-filter: use get occurrence of to handle
                -- multi-day recurrence patterns (e.g. Mon+Thu) correctly.
                -- Check not already in Phase 1 (exact subject, Canceled:, or Declined: prefix)
                set alreadyCaptured to false
                set checkDayEnd to checkDay + 1 * days
                set masterSubject to subject of masterEvent
                repeat with k from 1 to eventCount
                  set kEvent to item k of eventList
                  set kSubject to subject of kEvent
                  if kSubject = masterSubject or kSubject = "Canceled: " & masterSubject or kSubject = "Declined: " & masterSubject then
                    set kStart to start time of kEvent
                    if kStart >= checkDay and kStart < checkDayEnd then
                      set alreadyCaptured to true
                      exit repeat
                    end if
                  end if
                end repeat

                if not alreadyCaptured then
                  -- Verify the occurrence actually exists and is not declined/deleted.
                  -- get occurrence of throws an error if the series was ended or the
                  -- specific occurrence was deleted by the organizer.
                  set occExists to false
                  try
                    set masterTimeOfDay to time of masterStart
                    set checkDateTime to checkDay + masterTimeOfDay
                    set testOcc to get occurrence of masterEvent at checkDateTime
                    if testOcc is not missing value then
                      if not (subject of testOcc begins with "Declined:") then
                        set occExists to true
                      end if
                    end if
                  on error
                    set occExists to false
                  end try

                  if occExists then
                  -- Use actual occurrence times (handles rescheduled occurrences)
                  set evtStartOnDay to start time of testOcc
                  set evtEndOnDay to end time of testOcc

                  -- Only include if the occurrence start is within window
                  if evtStartOnDay >= todayDate and evtStartOnDay < endDate then
                    set evtSubject to subject of masterEvent
                    set evtStart to evtStartOnDay as string
                    set evtEnd to evtEndOnDay as string
                    set evtLoc to location of masterEvent
                    if evtLoc is missing value then set evtLoc to ""
                    set evtId to (id of masterEvent) as string

                    set evtAttendees to ""
                    try
                      set attList to attendees of masterEvent
                      set attCount to count of attList
                      repeat with j from 1 to attCount
                        try
                          set thisAtt to item j of attList
                          set addrRec to email address of thisAtt
                          set rawEmail to address of addrRec
                          set displayName to name of addrRec
                          set attEntry to rawEmail
                          if displayName is not "" and displayName is not missing value then
                            if displayName contains rawEmail then
                              set attEntry to displayName
                            else
                              set attEntry to displayName & " <" & rawEmail & ">"
                            end if
                          end if
                          if evtAttendees is "" then
                            set evtAttendees to attEntry
                          else
                            set evtAttendees to evtAttendees & ";" & attEntry
                          end if
                        on error
                        end try
                      end repeat
                    on error
                    end try

                    set evtRec to evtSubject & "|||" & evtStart & "|||" & evtEnd & "|||" & evtLoc & "|||" & evtId & "|||" & evtAttendees
                    if resultStr is "" then
                      set resultStr to evtRec
                    else
                      set resultStr to resultStr & "###" & evtRec
                    end if
                  end if -- evtStartOnDay in window
                  end if -- occExists
                end if -- not alreadyCaptured
              end if -- inRange
            end if -- checkDay >= recStartDay
            set checkDay to checkDay + 1 * days
          end repeat
        end if -- masterStart outside window
      end repeat

      return resultStr
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getUpcomingEvents] Raw result length: ${result.length}`);

    if (!result || result.trim() === '') {
      console.error('[getUpcomingEvents] No upcoming events found');
      return [];
    }

    const events = [];
    const records = result.split('###');
    for (const record of records) {
      try {
        const parts = record.split('|||');
        if (parts.length >= 5 && parts[0].trim()) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim(),
            attendees: parts[5]
              ? parts[5].split(';').filter((e: string) => e && e.trim())
              : []
          });
        }
      } catch (parseError) {
        console.error('[getUpcomingEvents] Error parsing event record:', parseError);
      }
    }

    const deduped = deduplicateCalendarEvents(events);
    console.error(`[getUpcomingEvents] Found ${deduped.length} upcoming events (${events.length} before dedup)`);
    return deduped;
  } catch (error) {
    console.error("[getUpcomingEvents] Error getting upcoming events:", error);
    throw error;
  }
}

// Function to search calendar events
async function searchEvents(searchTerm: string, limit: number = 10): Promise<any[]> {
  console.error(`[searchEvents] Searching for events with term: "${searchTerm}", limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set allEvents to every calendar event
      set i to 0
      set searchString to "${escapeAppleScriptString(searchTerm)}"
      set resultStr to ""

      repeat with theEvent in allEvents
        if (subject of theEvent contains searchString) or (location of theEvent contains searchString) then
          set i to i + 1
          set evtSubject to subject of theEvent
          set evtStart to (start time of theEvent) as string
          set evtEnd to (end time of theEvent) as string
          set evtLoc to location of theEvent
          if evtLoc is missing value then set evtLoc to ""
          set evtId to (id of theEvent) as string

          set evtRec to evtSubject & "|||" & evtStart & "|||" & evtEnd & "|||" & evtLoc & "|||" & evtId
          if resultStr is "" then
            set resultStr to evtRec
          else
            set resultStr to resultStr & "###" & evtRec
          end if

          if i >= ${limit} then
            exit repeat
          end if
        end if
      end repeat

      return resultStr
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[searchEvents] Raw result length: ${result.length}`);

    if (!result || result.trim() === '') {
      console.error('[searchEvents] No matching events found');
      return [];
    }

    const events = [];
    const records = result.split('###');
    for (const record of records) {
      try {
        const parts = record.split('|||');
        if (parts.length >= 5 && parts[0].trim()) {
          events.push({
            subject: parts[0].trim(),
            start: parts[1].trim(),
            end: parts[2].trim(),
            location: parts[3].trim() || "No location",
            id: parts[4].trim()
          });
        }
      } catch (parseError) {
        console.error('[searchEvents] Error parsing event record:', parseError);
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
async function createEvent(subject: string, start: string, end: string, location?: string, body?: string, attendees?: string): Promise<string> {
  console.error(`[createEvent] Creating event: "${subject}", start: ${start}, end: ${end}`);
  await checkOutlookAccess();
  
  // Parse the ISO date strings to a format AppleScript can understand
  const startDate = new Date(start);
  const endDate = new Date(end);
  
  // Format for AppleScript (month/day/year hour:minute:second)
  const formattedStart = `date "${startDate.getMonth() + 1}/${startDate.getDate()}/${startDate.getFullYear()} ${startDate.getHours()}:${startDate.getMinutes()}:${startDate.getSeconds()}"`;
  const formattedEnd = `date "${endDate.getMonth() + 1}/${endDate.getDate()}/${endDate.getFullYear()} ${endDate.getHours()}:${endDate.getMinutes()}:${endDate.getSeconds()}"`;
  
  // Escape strings for AppleScript
  const escapedSubject = escapeAppleScriptString(subject);
  const escapedLocation = location ? escapeAppleScriptString(location) : "";
  const escapedBody = body ? escapeAppleScriptString(body) : "";
  
  let script = `
    tell application "Microsoft Outlook"
      set theCalendar to default calendar
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
    const attendeeList = attendees.split(',').map(email => email.trim());
    
    for (const attendee of attendeeList) {
      const escapedAttendee = escapeAppleScriptString(attendee);
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
            const props = match.substring(1, match.length - 1).split(',');
            const contact: any = {};
            
            props.forEach(prop => {
              const parts = prop.split(':');
              if (parts.length >= 2) {
                const key = parts[0].trim();
                const value = parts.slice(1).join(':').trim();
                contact[key] = value;
              }
            });
            
            if (contact.name) {
              contacts.push({
                name: contact.name,
                email: contact.email || "No email",
                phone: contact.phone || "No phone"
              });
            }
          } catch (parseError) {
            console.error('[listContacts] Error parsing contact match:', parseError);
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
        const simplifiedContacts = result.split(", ").map(name => ({
          name: name,
          email: "Not available with simplified method",
          phone: "Not available with simplified method"
        }));
        
        console.error(`[listContacts] Found ${simplifiedContacts.length} contacts using alternative method`);
        return simplifiedContacts;
      } catch (altError) {
        console.error("[listContacts] Alternative method also failed:", altError);
        throw new Error(`Error accessing contacts. The error might be related to Outlook permissions or configuration: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }

// Function to search contacts
// Function to search contacts with improved AppleScript syntax
async function searchContacts(searchTerm: string, limit: number = 10): Promise<any[]> {
    console.error(`[searchContacts] Searching for contacts with term: "${searchTerm}", limit: ${limit}`);
    await checkOutlookAccess();
    
    const script = `
      tell application "Microsoft Outlook"
        set searchResults to {}
        set allContacts to contacts
        set i to 0
        set searchString to "${escapeAppleScriptString(searchTerm)}"

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
            const props = match.substring(1, match.length - 1).split(',');
            const contact: any = {};
            
            props.forEach(prop => {
              const parts = prop.split(':');
              if (parts.length >= 2) {
                const key = parts[0].trim();
                const value = parts.slice(1).join(':').trim();
                contact[key] = value;
              }
            });
            
            if (contact.name) {
              contacts.push({
                name: contact.name,
                email: contact.email || "No email",
                phone: contact.phone || "No phone"
              });
            }
          } catch (parseError) {
            console.error('[searchContacts] Error parsing contact match:', parseError);
          }
        }
      }
      
      console.error(`[searchContacts] Found ${contacts.length} matching contacts`);
      return contacts;
    } catch (error) {
      console.error("[searchContacts] Error searching contacts:", error);
      
      // Try an alternative approach with a simpler script that just returns names
      try {
        const alternativeScript = `
          tell application "Microsoft Outlook"
            set matchingContacts to {}
            set searchString to "${escapeAppleScriptString(searchTerm)}"
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
        const simplifiedContacts = result.split(", ").map(name => ({
          name: name,
          email: "Not available with simplified method",
          phone: "Not available with simplified method"
        }));
        
        console.error(`[searchContacts] Found ${simplifiedContacts.length} contacts using alternative method`);
        return simplifiedContacts;
      } catch (altError) {
        console.error("[searchContacts] Alternative method also failed:", altError);
        throw new Error(`Error searching contacts. The error might be related to Outlook permissions or configuration: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  }

// ====================================================
// 7. TYPE GUARDS
// ====================================================

// Type guards for arguments
function isMailArgs(args: unknown): args is {
  operation: "unread" | "search" | "send" | "folders" | "read";
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
} {
  if (typeof args !== "object" || args === null) return false;
  
  const { operation } = args as any;
  
  if (!operation || !["unread", "search", "send", "folders", "read"].includes(operation)) {
    return false;
  }
  
  // Check required fields based on operation
  switch (operation) {
    case "search":
      if (!(args as any).searchTerm) return false;
      break;
    case "send":
      if (!(args as any).to || !(args as any).subject || !(args as any).body) return false;
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
  
  if (!operation || !["today", "upcoming", "search", "create"].includes(operation)) {
    return false;
  }
  
  // Check required fields based on operation
  switch (operation) {
    case "search":
      if (!(args as any).searchTerm) return false;
      break;
    case "create":
      if (!(args as any).subject || !(args as any).start || !(args as any).end) return false;
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
              content: [{ 
                type: "text", 
                text: emails.length > 0 ? 
                  `Found ${emails.length} unread email(s)${args.folder ? ` in folder "${args.folder}"` : ''}\n\n` +
                  emails.map(email => 
                    `[${email.dateSent}] From: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? '...' : ''}`
                  ).join("\n\n") :
                  `No unread emails found${args.folder ? ` in folder "${args.folder}"` : ''}`
              }],
              isError: false
            };
          }
          
          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const emails = await searchEmails(args.searchTerm, args.folder, args.limit);
            return {
              content: [{ 
                type: "text", 
                text: emails.length > 0 ? 
                  `Found ${emails.length} email(s) for "${args.searchTerm}"${args.folder ? ` in folder "${args.folder}"` : ''}\n\n` +
                  emails.map(email => 
                    `[${email.dateSent}] From: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? '...' : ''}`
                  ).join("\n\n") :
                  `No emails found for "${args.searchTerm}"${args.folder ? ` in folder "${args.folder}"` : ''}`
              }],
              isError: false
            };
          }
          
          // Update the handler in CallToolRequestSchema
          case "send": {
            if (!args.to || !args.subject || !args.body) {
              throw new Error("Recipient (to), subject, and body are required for send operation");
            }
            
            // Validate attachments if provided
            if (args.attachments && !Array.isArray(args.attachments)) {
              throw new Error("Attachments must be an array of file paths");
            }
            
            // Log attachment information for debugging
            console.error(`[CallTool] Send email with attachments: ${args.attachments ? JSON.stringify(args.attachments) : 'none'}`);
            
            const result = await sendEmail(
              args.to, 
              args.subject, 
              args.body, 
              args.cc, 
              args.bcc, 
              args.isHtml || false,
              args.attachments
            );
            
            return {
              content: [{ type: "text", text: result }],
              isError: false
            };
          }
          
          case "folders": {
            const folders = await getMailFolders();
            return {
              content: [{ 
                type: "text", 
                text: folders.length > 0 ? 
                  `Found ${folders.length} mail folders:\n\n${folders.join("\n")}` :
                  "No mail folders found. Make sure Outlook is running and properly configured."
              }],
              isError: false
            };
          }
          
          case "read": {
            const emails = await readEmails(args.folder, args.limit);
            return {
              content: [{ 
                type: "text", 
                text: emails.length > 0 ? 
                  `Found ${emails.length} email(s)${args.folder ? ` in folder "${args.folder}"` : ''}\n\n` +
                  emails.map(email => 
                    `[${email.dateSent}] From: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? '...' : ''}`
                  ).join("\n\n") :
                  `No emails found${args.folder ? ` in folder "${args.folder}"` : ''}`
              }],
              isError: false
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
              content: [{ 
                type: "text", 
                text: events.length > 0 ? 
                  `Found ${events.length} event(s) for today:\n\n` +
                  events.map(event => 
                    `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`
                  ).join("\n\n") :
                  "No events found for today"
              }],
              isError: false
            };
          }
          
          case "upcoming": {
            const days = args.days || 7;
            const events = await getUpcomingEvents(days, args.limit);
            return {
              content: [{ 
                type: "text", 
                text: events.length > 0 ? 
                  `Found ${events.length} upcoming event(s) for the next ${days} days:\n\n` +
                  events.map(event => 
                    `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`
                  ).join("\n\n") :
                  `No upcoming events found for the next ${days} days`
              }],
              isError: false
            };
          }
          
          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const events = await searchEvents(args.searchTerm, args.limit);
            return {
              content: [{ 
                type: "text", 
                text: events.length > 0 ? 
                  `Found ${events.length} event(s) matching "${args.searchTerm}":\n\n` +
                  events.map(event => 
                    `${event.subject}\nTime: ${event.start} - ${event.end}\nLocation: ${event.location}`
                  ).join("\n\n") :
                  `No events found matching "${args.searchTerm}"`
              }],
              isError: false
            };
          }
          
          case "create": {
            if (!args.subject || !args.start || !args.end) {
              throw new Error("Subject, start time, and end time are required for create operation");
            }
            const result = await createEvent(args.subject, args.start, args.end, args.location, args.body, args.attendees);
            return {
              content: [{ type: "text", text: result }],
              isError: false
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
              content: [{ 
                type: "text", 
                text: contacts.length > 0 ? 
                  `Found ${contacts.length} contact(s):\n\n` +
                  contacts.map(contact => 
                    `Name: ${contact.name}\nEmail: ${contact.email}\nPhone: ${contact.phone}`
                  ).join("\n\n") :
                  "No contacts found"
              }],
              isError: false
            };
          }
          
          case "search": {
            if (!args.searchTerm) {
              throw new Error("Search term is required for search operation");
            }
            const contacts = await searchContacts(args.searchTerm, args.limit);
            return {
              content: [{ 
                type: "text", 
                text: contacts.length > 0 ? 
                  `Found ${contacts.length} contact(s) matching "${args.searchTerm}":\n\n` +
                  contacts.map(contact => 
                    `Name: ${contact.name}\nEmail: ${contact.email}\nPhone: ${contact.phone}`
                  ).join("\n\n") :
                  `No contacts found matching "${args.searchTerm}"`
              }],
              isError: false
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