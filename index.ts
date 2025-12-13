#!/usr/bin/env bun
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { runAppleScript } from 'run-applescript';
import { parseEmailOutput, buildFolderRef, buildNestedFolderRef, escapeForAppleScript } from './helpers';

// Re-export helpers for testing
export { parseEmailOutput, buildFolderRef, buildNestedFolderRef, escapeForAppleScript } from './helpers';

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
        description: "Operation to perform: 'unread', 'search', 'send', 'draft', 'reply', 'forward', 'folders', 'read', 'create_folder', 'move_email', 'rename_folder', 'delete_folder', 'count', or 'save_attachments'",
        enum: ["unread", "search", "send", "draft", "reply", "forward", "folders", "read", "create_folder", "move_email", "rename_folder", "delete_folder", "count", "save_attachments"]
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
      startDate: {
        type: "string",
        description: "Filter emails sent on or after this date (ISO format, e.g., '2025-12-01')"
      },
      endDate: {
        type: "string",
        description: "Filter emails sent on or before this date (ISO format, e.g., '2025-12-05')"
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
        description: "Whether the body content is HTML. ALWAYS use HTML (true) for email creation - Outlook renders HTML natively but treats markdown as plain text. Default: false"
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
        description: "File paths to attach to the email (optional for send, reply, and forward operations)",
        items: {
          type: "string"
        }
      },
      name: {
        type: "string",
        description: "Folder name to create (required for create_folder operation)"
      },
      parent: {
        type: "string",
        description: "Parent folder path for nesting (optional for create_folder operation)"
      },
      messageId: {
        type: "string",
        description: "Email message ID (required for move_email, forward, and reply operations)"
      },
      forwardTo: {
        type: "string",
        description: "Email address to forward to (required for forward operation)"
      },
      forwardCc: {
        type: "string",
        description: "CC email address for forward (optional for forward operation)"
      },
      forwardComment: {
        type: "string",
        description: "Comment to add above forwarded message (optional for forward operation)"
      },
      includeOriginalAttachments: {
        type: "boolean",
        description: "Whether to include original email attachments when forwarding (optional for forward operation, default: true)"
      },
      replyBody: {
        type: "string",
        description: "Reply message content (required for reply operation)"
      },
      replyAll: {
        type: "boolean",
        description: "Whether to reply to all recipients (optional for reply operation, default: false)"
      },
      targetFolder: {
        type: "string",
        description: "Destination folder path (required for move_email operation)"
      },
      newName: {
        type: "string",
        description: "New folder name (required for rename_folder operation)"
      },
      destinationFolder: {
        type: "string",
        description: "Destination folder path to save attachments (required for save_attachments operation)"
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
        description: "Operation to perform: 'today', 'upcoming', 'search', 'create', 'delete', 'accept', 'decline', 'tentative', or 'propose_new_time'",
        enum: ["today", "upcoming", "search", "create", "delete", "accept", "decline", "tentative", "propose_new_time"]
      },
      deleteSubject: {
        type: "string",
        description: "Subject of event to delete (required for delete operation). Must match exactly."
      },
      deleteDate: {
        type: "string",
        description: "Date of event to delete in YYYY-MM-DD format (required for delete operation)"
      },
      responseSubject: {
        type: "string",
        description: "Subject of meeting invite to respond to (required for accept/decline/tentative/propose_new_time). Must match exactly."
      },
      responseDate: {
        type: "string",
        description: "Date of meeting invite in YYYY-MM-DD format (required for accept/decline/tentative/propose_new_time)"
      },
      proposedStart: {
        type: "string",
        description: "Proposed new start time in ISO format (required for propose_new_time operation)"
      },
      proposedEnd: {
        type: "string",
        description: "Proposed new end time in ISO format (required for propose_new_time operation)"
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
  
  const folderPath = folder === "Inbox" ? "inbox" : folder;
  const script = `
    tell application "Microsoft Outlook"
      try
        set theFolder to ${folderPath} -- Use the specified folder or default to inbox
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
async function searchEmails(searchTerm: string, folder: string = "Inbox", limit: number = 10, startDate?: string, endDate?: string): Promise<any[]> {
  console.error(`[searchEmails] Searching for "${searchTerm}" in folder: ${folder}, limit: ${limit}, startDate: ${startDate}, endDate: ${endDate}`);
  await checkOutlookAccess();

  const folderRef = buildFolderRef(folder);

  // Build date filter AppleScript code
  let dateFilterSetup = "";
  let dateFilterCheck = "";

  if (startDate || endDate) {
    if (startDate) {
      // Parse date string directly to avoid timezone issues
      // Input format: "YYYY-MM-DD"
      const [year, month, day] = startDate.split('-').map(Number);
      // AppleScript requires 12-hour format with AM/PM
      dateFilterSetup += `set filterStartDate to date "${month}/${day}/${year} 12:00:00 AM"\n`;
    }
    if (endDate) {
      // Parse date string directly to avoid timezone issues
      // Input format: "YYYY-MM-DD"
      const [year, month, day] = endDate.split('-').map(Number);
      // Set to end of day - AppleScript requires 12-hour format with AM/PM
      dateFilterSetup += `set filterEndDate to date "${month}/${day}/${year} 11:59:59 PM"\n`;
    }

    // Build the date check condition
    const startCheck = startDate ? "msgSentDate >= filterStartDate" : "";
    const endCheck = endDate ? "msgSentDate <= filterEndDate" : "";
    const dateChecks = [startCheck, endCheck].filter(c => c).join(" and ");
    dateFilterCheck = `
              set msgSentDate to time sent of theMsg
              if not (${dateChecks}) then
                -- Skip this message, outside date range
              else`;
  }

  const script = `
    tell application "Microsoft Outlook"
      try
        set searchString to "${searchTerm.replace(/"/g, '\\"')}"
        set messageOutput to ""
        set totalFound to 0
        ${dateFilterSetup}

        -- Search only the specified folder
        set theFolder to ${folderRef}
        set folderMsgs to messages of theFolder

        repeat with theMsg in folderMsgs
          if totalFound >= ${limit} then exit repeat

          try
            ${dateFilterCheck}
            set msgSubject to subject of theMsg
            set msgContent to ""

            -- Get plain text content
            try
              set msgContent to plain text content of theMsg
            on error
              try
                set msgContent to content of theMsg
              on error
                set msgContent to "[Content not available]"
              end try
            end try

            -- Check if search term is in subject or content (AppleScript 'contains' is case-insensitive)
            if (msgSubject contains searchString) or (msgContent contains searchString) then
              -- Get message ID
              set msgId to id of theMsg as string

              -- Get sender info
              set msgSender to "Unknown"
              try
                set senderObj to sender of theMsg
                if class of senderObj is text then
                  set msgSender to senderObj
                else
                  try
                    set msgSender to address of senderObj
                  on error
                    try
                      set msgSender to name of senderObj
                    on error
                      set msgSender to senderObj as text
                    end try
                  end try
                end if
              end try

              set msgDate to time sent of theMsg as string

              -- Truncate content for output
              if length of msgContent > 5000 then
                set msgContent to (text 1 thru 5000 of msgContent) & "..."
              end if

              -- Build output with delimiters
              set messageOutput to messageOutput & "<<<MSG>>>" & msgSubject & "<<<ID>>>" & msgId & "<<<FROM>>>" & msgSender & "<<<DATE>>>" & msgDate & "<<<CONTENT>>>" & msgContent & "<<<ENDMSG>>>"
              set totalFound to totalFound + 1
            end if
            ${dateFilterCheck ? "end if" : ""}
          end try
        end repeat

        return messageOutput
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

    // Parse messages using helper function
    const emails = parseEmailOutput(result);

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

  // Parse multiple recipients (comma-separated)
  const parseRecipients = (recipientStr: string): Array<{name: string, address: string}> => {
    return recipientStr.split(',').map(email => {
      const trimmed = email.trim();
      return { name: extractNameFromEmail(trimmed), address: trimmed };
    });
  };

  const ccRecipients = cc ? parseRecipients(cc) : [];
  const bccRecipients = bcc ? parseRecipients(bcc) : [];

  // Escape special characters for AppleScript
  const escapedSubject = subject.replace(/"/g, '\\"');
  const escapedBody = body.replace(/\\/g, '\\\\').replace(/"/g, '\\"');

  // Process attachments: Convert to absolute paths if they are relative
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    processedAttachments = attachments.map(path => {
      // Check if path is absolute (starts with /)
      if (path.startsWith('/')) {
        return path;
      }
      // Get current working directory and join with relative path
      const cwd = process.cwd();
      return `${cwd}/${path}`;
    });
    console.error(`[sendEmail] Processed attachments: ${JSON.stringify(processedAttachments)}`);
  }
  
  // Create attachment script part with better error handling
  const attachmentScript = processedAttachments.length > 0 
    ? processedAttachments.map(filePath => {
      const escapedPath = filePath.replace(/"/g, '\\"');
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
            `set content of msg to "${escapedBody}"`
          :
            `set plain text content of msg to "${escapedBody}"`
          }
          
          tell msg
            set recipTo to make new to recipient with properties {email address:{name:"${toName}", address:"${to}"}}
            ${ccRecipients.map(r => `make new cc recipient with properties {email address:{name:"${r.name}", address:"${r.address}"}}`).join('\n            ')}
            ${bccRecipients.map(r => `make new bcc recipient with properties {email address:{name:"${r.name}", address:"${r.address}"}}`).join('\n            ')}

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
              `set content of theMessage to "${escapedBody}"`
            :
              `set plain text content of theMessage to "${escapedBody}"`
            }
            
            set to recipients of theMessage to {"${to}"}
            ${cc ? `set cc recipients of theMessage to {"${cc}"}` : ''}
            ${bcc ? `set bcc recipients of theMessage to {"${bcc}"}` : ''}
            
            ${processedAttachments.map(filePath => {
              const escapedPath = filePath.replace(/"/g, '\\"');
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
                `set content of newMessage to "${escapedBody}"`
              :
                `set plain text content of newMessage to "${escapedBody}"`
              }
              
              set to recipients of newMessage to {"${to}"}
              ${cc ? `set cc recipients of newMessage to {"${cc}"}` : ''}
              ${bcc ? `set bcc recipients of newMessage to {"${bcc}"}` : ''}
              
              ${processedAttachments.map(filePath => {
                const escapedPath = filePath.replace(/"/g, '\\"');
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

// Function to create a draft email (opens in Outlook for editing)
async function createDraft(
  to: string,
  subject: string,
  body: string,
  cc?: string,
  bcc?: string,
  isHtml: boolean = false,
  attachments?: string[]
): Promise<string> {
  console.error(`[createDraft] Creating draft to: ${to}, subject: "${subject}"`);
  console.error(`[createDraft] Attachments: ${attachments ? JSON.stringify(attachments) : 'none'}`);

  await checkOutlookAccess();

  // Extract name from email if possible (for display name)
  const extractNameFromEmail = (email: string): string => {
    const namePart = email.split('@')[0];
    return namePart
      .split('.')
      .map(part => part.charAt(0).toUpperCase() + part.slice(1))
      .join(' ');
  };

  const toName = extractNameFromEmail(to);

  // Parse multiple recipients (comma-separated)
  const parseRecipients = (recipientStr: string): Array<{name: string, address: string}> => {
    return recipientStr.split(',').map(email => {
      const trimmed = email.trim();
      return { name: extractNameFromEmail(trimmed), address: trimmed };
    });
  };

  const ccRecipients = cc ? parseRecipients(cc) : [];
  const bccRecipients = bcc ? parseRecipients(bcc) : [];

  const escapedSubject = subject.replace(/"/g, '\\"');
  const escapedBody = body.replace(/\\/g, '\\\\').replace(/"/g, '\\"');

  // Process attachments
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    processedAttachments = attachments.map(path => {
      if (path.startsWith('/')) {
        return path;
      }
      const cwd = process.cwd();
      return `${cwd}/${path}`;
    });
  }

  const attachmentScript = processedAttachments.length > 0
    ? processedAttachments.map(filePath => {
      const escapedPath = filePath.replace(/"/g, '\\"');
      return `
        try
          set attachmentFile to POSIX file "${escapedPath}"
          make new attachment at newMessage with properties {file:attachmentFile}
        on error errMsg
          log "Failed to attach file: ${escapedPath} - Error: " & errMsg
        end try
      `;
    }).join('\n')
    : '';

  const script = `
    tell application "Microsoft Outlook"
      try
        set newMessage to make new outgoing message with properties {subject:"${escapedSubject}"}

        ${isHtml ?
          `set content of newMessage to "${escapedBody}"`
        :
          `set plain text content of newMessage to "${escapedBody}"`
        }

        tell newMessage
          make new to recipient with properties {email address:{name:"${toName}", address:"${to}"}}
          ${ccRecipients.map(r => `make new cc recipient with properties {email address:{name:"${r.name}", address:"${r.address}"}}`).join('\n          ')}
          ${bccRecipients.map(r => `make new bcc recipient with properties {email address:{name:"${r.name}", address:"${r.address}"}}`).join('\n          ')}
        end tell

        ${attachmentScript}

        -- Open the message window for editing
        open newMessage

        -- Bring Outlook to front
        activate

        return "Draft created and opened in Outlook"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[createDraft] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[createDraft] Error creating draft:", error);
    return `Error: ${error}`;
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

// Function to create a mail folder
async function createFolder(name: string, parent?: string): Promise<string> {
  console.error(`[createFolder] Creating folder: ${name}${parent ? ` under ${parent}` : ''}`);
  await checkOutlookAccess();

  let script: string;
  if (parent) {
    const parentRef = buildNestedFolderRef(parent);
    script = `
      tell application "Microsoft Outlook"
        try
          set parentFolder to ${parentRef}
          set newFolder to make new mail folder at parentFolder with properties {name:"${escapeForAppleScript(name)}"}
          return "Folder created: " & name of newFolder
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;
  } else {
    script = `
      tell application "Microsoft Outlook"
        try
          set newFolder to make new mail folder with properties {name:"${escapeForAppleScript(name)}"}
          return "Folder created: " & name of newFolder
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;
  }

  try {
    const result = await runAppleScript(script);
    console.error(`[createFolder] Result: ${result}`);
    if (result.startsWith("Error:")) {
      return result;
    }
    return result;
  } catch (error) {
    console.error("[createFolder] Error creating folder:", error);
    return `Error: ${error}`;
  }
}

// Function to move an email to a folder
async function moveEmail(messageId: string, targetFolder: string): Promise<string> {
  console.error(`[moveEmail] Moving message ${messageId} to folder: ${targetFolder}`);
  await checkOutlookAccess();

  const folderRef = buildNestedFolderRef(targetFolder);

  const script = `
    tell application "Microsoft Outlook"
      try
        set theMsg to message id ${messageId}
        set targetFolder to ${folderRef}
        move theMsg to targetFolder
        return "Email moved successfully to ${escapeForAppleScript(targetFolder)}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[moveEmail] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[moveEmail] Error moving email:", error);
    return `Error: ${error}`;
  }
}

// Function to forward an email
async function forwardEmail(messageId: string, forwardTo: string, forwardCc?: string, forwardComment?: string, attachments?: string[], includeOriginalAttachments: boolean = true): Promise<string> {
  console.error(`[forwardEmail] Forwarding message ${messageId} to: ${forwardTo}`);
  console.error(`[forwardEmail] New attachments: ${attachments ? JSON.stringify(attachments) : 'none'}`);
  console.error(`[forwardEmail] Include original attachments: ${includeOriginalAttachments}`);
  await checkOutlookAccess();

  const escapedComment = forwardComment ? escapeForAppleScript(forwardComment) : "";
  const escapedTo = escapeForAppleScript(forwardTo);
  const escapedCc = forwardCc ? escapeForAppleScript(forwardCc) : "";

  // Process new attachments: Convert to absolute paths if they are relative
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    processedAttachments = attachments.map(path => {
      if (path.startsWith('/')) {
        return path;
      }
      const cwd = process.cwd();
      return `${cwd}/${path}`;
    });
    console.error(`[forwardEmail] Processed attachments: ${JSON.stringify(processedAttachments)}`);
  }

  // Create script to add new attachments
  const attachmentScript = processedAttachments.length > 0
    ? processedAttachments.map(filePath => {
      const escapedPath = filePath.replace(/"/g, '\\"');
      return `
        try
          set attachmentFile to POSIX file "${escapedPath}"
          make new attachment at fwdMsg with properties {file:attachmentFile}
        on error errMsg
          log "Failed to attach file: ${escapedPath} - Error: " & errMsg
        end try
      `;
    }).join('\n')
    : '';

  // Script to remove original attachments if requested
  const removeOriginalAttachmentsScript = !includeOriginalAttachments ? `
        -- Remove original attachments from forwarded message
        try
          set attachmentCount to count of attachments of fwdMsg
          repeat while attachmentCount > 0
            delete attachment 1 of fwdMsg
            set attachmentCount to attachmentCount - 1
          end repeat
        on error errMsg
          log "Note: Could not remove original attachments - " & errMsg
        end try
  ` : '';

  const script = `
    tell application "Microsoft Outlook"
      try
        set theMsg to message id ${messageId}
        set fwdMsg to forward theMsg without opening window

        -- Add recipient
        tell fwdMsg
          make new to recipient with properties {email address:{address:"${escapedTo}"}}
          ${forwardCc ? `make new cc recipient with properties {email address:{address:"${escapedCc}"}}` : ''}
        end tell

        ${removeOriginalAttachmentsScript}

        -- Add new attachments if provided
        ${attachmentScript}

        -- Add comment above forwarded content if provided
        ${forwardComment ? `
        set currentContent to content of fwdMsg
        set content of fwdMsg to "${escapedComment}" & return & return & currentContent
        ` : ''}

        -- Send the forward
        send fwdMsg

        return "Email forwarded successfully to ${escapedTo}${processedAttachments.length > 0 ? ` with ${processedAttachments.length} new attachment(s)` : ''}${!includeOriginalAttachments ? ' (original attachments removed)' : ''}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[forwardEmail] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[forwardEmail] Error forwarding email:", error);
    return `Error: ${error}`;
  }
}

// Function to reply to an email (preserves thread)
async function replyEmail(messageId: string, replyBody: string, replyAll: boolean = false, isHtml: boolean = false, attachments?: string[]): Promise<string> {
  console.error(`[replyEmail] Replying to message ${messageId}, replyAll: ${replyAll}, isHtml: ${isHtml}`);
  console.error(`[replyEmail] Attachments: ${attachments ? JSON.stringify(attachments) : 'none'}`);
  await checkOutlookAccess();

  const escapedBody = escapeForAppleScript(replyBody);
  const replyCommand = replyAll ? "reply to theMsg with reply to all" : "reply to theMsg";

  // Process attachments: Convert to absolute paths if they are relative
  let processedAttachments: string[] = [];
  if (attachments && attachments.length > 0) {
    processedAttachments = attachments.map(path => {
      if (path.startsWith('/')) {
        return path;
      }
      const cwd = process.cwd();
      return `${cwd}/${path}`;
    });
    console.error(`[replyEmail] Processed attachments: ${JSON.stringify(processedAttachments)}`);
  }

  // Create attachment script part
  const attachmentScript = processedAttachments.length > 0
    ? processedAttachments.map(filePath => {
      const escapedPath = filePath.replace(/"/g, '\\"');
      return `
        try
          set attachmentFile to POSIX file "${escapedPath}"
          make new attachment at replyMsg with properties {file:attachmentFile}
        on error errMsg
          log "Failed to attach file: ${escapedPath} - Error: " & errMsg
        end try
      `;
    }).join('\n')
    : '';

  const script = `
    tell application "Microsoft Outlook"
      try
        set theMsg to message id ${messageId}
        set replyMsg to ${replyCommand} without opening window

        -- Set the reply content (prepend to existing quoted content)
        ${isHtml ? `
        set currentContent to content of replyMsg
        set content of replyMsg to "${escapedBody}" & "<br><br>" & currentContent
        ` : `
        set currentContent to plain text content of replyMsg
        set plain text content of replyMsg to "${escapedBody}" & return & return & currentContent
        `}

        -- Add attachments if provided
        ${attachmentScript}

        -- Send the reply
        send replyMsg

        return "Reply sent successfully${processedAttachments.length > 0 ? ` with ${processedAttachments.length} attachment(s)` : ''}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[replyEmail] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[replyEmail] Error replying to email:", error);
    return `Error: ${error}`;
  }
}

// Function to rename a folder
async function renameFolder(folder: string, newName: string): Promise<string> {
  console.error(`[renameFolder] Renaming folder: ${folder} to ${newName}`);
  await checkOutlookAccess();

  const folderRef = buildNestedFolderRef(folder);

  const script = `
    tell application "Microsoft Outlook"
      try
        set theFolder to ${folderRef}
        set name of theFolder to "${escapeForAppleScript(newName)}"
        return "Folder renamed to: ${escapeForAppleScript(newName)}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[renameFolder] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[renameFolder] Error renaming folder:", error);
    return `Error: ${error}`;
  }
}

// Function to delete a folder
async function deleteFolder(folder: string): Promise<string> {
  console.error(`[deleteFolder] Deleting folder: ${folder}`);
  await checkOutlookAccess();

  const folderRef = buildNestedFolderRef(folder);

  const script = `
    tell application "Microsoft Outlook"
      try
        set theFolder to ${folderRef}
        set msgCount to count of messages of theFolder
        if msgCount > 0 then
          return "Error: Folder contains " & msgCount & " email(s). Move or delete emails first."
        end if
        delete theFolder
        return "Folder deleted: ${escapeForAppleScript(folder)}"
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[deleteFolder] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[deleteFolder] Error deleting folder:", error);
    return `Error: ${error}`;
  }
}

// Function to count emails in a folder
async function countEmails(folder: string = "Inbox"): Promise<string> {
  console.error(`[countEmails] Counting emails in folder: ${folder}`);
  await checkOutlookAccess();

  const folderRef = buildFolderRef(folder);

  const script = `
    tell application "Microsoft Outlook"
      try
        set theFolder to ${folderRef}
        set totalCount to count of messages of theFolder
        set unreadCount to count of (messages of theFolder whose is read is false)
        return "Total: " & totalCount & ", Unread: " & unreadCount
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[countEmails] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[countEmails] Error counting emails:", error);
    return `Error: ${error}`;
  }
}

// Function to save attachments from an email to a destination folder
async function saveAttachments(messageId: string, destinationFolder: string): Promise<string> {
  console.error(`[saveAttachments] Saving attachments from message ${messageId} to ${destinationFolder}`);
  await checkOutlookAccess();

  const escapedDestination = destinationFolder.replace(/"/g, '\\"');

  const script = `
    tell application "Microsoft Outlook"
      try
        set theMsg to message id ${messageId}
        set theAttachments to attachments of theMsg
        set attachmentCount to count of theAttachments

        if attachmentCount is 0 then
          return "No attachments found in this email"
        end if

        set savedFiles to {}
        repeat with theAttachment in theAttachments
          set fileName to name of theAttachment
          set filePath to "${escapedDestination}/" & fileName
          save theAttachment in POSIX file filePath
          set end of savedFiles to fileName
        end repeat

        return "Saved " & attachmentCount & " attachment(s): " & (savedFiles as string)
      on error errMsg
        return "Error: " & errMsg
      end try
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[saveAttachments] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[saveAttachments] Error saving attachments:", error);
    return `Error: ${error}`;
  }
}

  // Function to read emails in a folder that uses simple AppleScript
async function readEmails(folder: string = "Inbox", limit: number = 10, startDate?: string, endDate?: string): Promise<any[]> {
    console.error(`[readEmails] Reading emails from folder: ${folder}, limit: ${limit}, startDate: ${startDate}, endDate: ${endDate}`);
    await checkOutlookAccess();

    const folderRef = buildFolderRef(folder);

    // Build date filter AppleScript code
    let dateFilterSetup = "";
    let dateFilterCheck = "";
    let dateFilterEndIf = "";

    if (startDate || endDate) {
      if (startDate) {
        // Parse date string directly to avoid timezone issues
        // Input format: "YYYY-MM-DD"
        const [year, month, day] = startDate.split('-').map(Number);
        // AppleScript requires 12-hour format with AM/PM
        const startDateStr = `${month}/${day}/${year} 12:00:00 AM`;
        console.error(`[readEmails] Start date filter: ${startDateStr}`);
        dateFilterSetup += `set filterStartDate to date "${startDateStr}"\n`;
      }
      if (endDate) {
        // Parse date string directly to avoid timezone issues
        // Input format: "YYYY-MM-DD"
        const [year, month, day] = endDate.split('-').map(Number);
        // Set to end of day - AppleScript requires 12-hour format with AM/PM
        const endDateStr = `${month}/${day}/${year} 11:59:59 PM`;
        console.error(`[readEmails] End date filter: ${endDateStr}`);
        dateFilterSetup += `set filterEndDate to date "${endDateStr}"\n`;
      }

      const startCheck = startDate ? "msgSentDate >= filterStartDate" : "";
      const endCheck = endDate ? "msgSentDate <= filterEndDate" : "";
      const dateChecks = [startCheck, endCheck].filter(c => c).join(" and ");
      dateFilterCheck = `
              set msgSentDate to time sent of theMsg
              if (${dateChecks}) then`;
      dateFilterEndIf = "end if";
    }

    const script = `
      tell application "Microsoft Outlook"
        try
          set theFolder to ${folderRef}
          set allMessages to messages of theFolder
          set msgCount to count of allMessages
          ${dateFilterSetup}

          set messageOutput to ""
          set foundCount to 0

          repeat with i from 1 to msgCount
            if foundCount >= ${limit} then exit repeat
            try
              set theMsg to item i of allMessages
              ${dateFilterCheck}
              set msgId to id of theMsg as string
              set msgSubject to subject of theMsg

              -- Get sender
              set msgSender to "Unknown"
              try
                set senderObj to sender of theMsg
                if class of senderObj is text then
                  set msgSender to senderObj
                else
                  try
                    set msgSender to address of senderObj
                  on error
                    try
                      set msgSender to name of senderObj
                    on error
                      set msgSender to senderObj as text
                    end try
                  end try
                end if
              end try

              set msgDate to time sent of theMsg as string

              -- Get content (truncated)
              set msgContent to ""
              try
                set msgContent to content of theMsg
                if length of msgContent > 5000 then
                  set msgContent to (text 1 thru 5000 of msgContent) & "..."
                end if
              on error
                set msgContent to "[Content not available]"
              end try

              set messageOutput to messageOutput & "<<<MSG>>>" & msgSubject & "<<<ID>>>" & msgId & "<<<FROM>>>" & msgSender & "<<<DATE>>>" & msgDate & "<<<CONTENT>>>" & msgContent & "<<<ENDMSG>>>"
              set foundCount to foundCount + 1
              ${dateFilterEndIf}
            end try
          end repeat

          return messageOutput
        on error errMsg
          return "Error: " & errMsg
        end try
      end tell
    `;

    try {
      const result = await runAppleScript(script);
      console.error(`[readEmails] AppleScript result length: ${result.length}`);

      if (result.startsWith("Error:")) {
        throw new Error(result);
      }

      // Parse messages using helper function
      const emails = parseEmailOutput(result);

      console.error(`[readEmails] Successfully parsed ${emails.length} emails from ${folder}`);
      return emails;
    } catch (error) {
      console.error("[readEmails] Error reading emails:", error);
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

      -- Find the calendar with the most events (main calendar)
      set allCals to every calendar
      set theCalendar to item 1 of allCals
      set maxEvents to 0
      repeat with cal in allCals
        set eventCount to count of (every calendar event of cal)
        if eventCount > maxEvents then
          set maxEvents to eventCount
          set theCalendar to cal
        end if
      end repeat

      set todayDate to current date
      set startOfDay to todayDate - (time of todayDate)
      set endOfDay to startOfDay + 1 * days
      set allEvents to every calendar event of theCalendar
      set limitCount to 0

      repeat with theEvent in allEvents
        set eventStart to start time of theEvent
        if eventStart >= startOfDay and eventStart < endOfDay then
          set limitCount to limitCount + 1
          set eventData to {subject:subject of theEvent, ¬
                       start:eventStart, ¬
                       |end|:end time of theEvent, ¬
                       location:location of theEvent, ¬
                       id:id of theEvent}
          set end of todayEvents to eventData
          if limitCount >= ${limit} then
            exit repeat
          end if
        end if
      end repeat

      return todayEvents
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[getTodayEvents] Raw result length: ${result.length}`);

    // Parse the results - split by ", subject:" to separate events
    const events = [];
    if (result && result.includes('subject:')) {
      const eventStrings = result.split(/, subject:/).map((s, i) => i === 0 ? s : 'subject:' + s);

      for (const eventStr of eventStrings) {
        try {
          const event: any = {};
          const parts = eventStr.split(/, (?=subject:|start:|end:|location:|id:)/i);
          for (const part of parts) {
            const colonIdx = part.indexOf(':');
            if (colonIdx > 0) {
              const key = part.substring(0, colonIdx).trim().toLowerCase();
              const value = part.substring(colonIdx + 1).trim();
              event[key] = value;
            }
          }

          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.start,
              end: event.end,
              location: event.location || "No location",
              id: event.id
            });
          }
        } catch (parseError) {
          console.error('[getTodayEvents] Error parsing event:', parseError);
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
async function getUpcomingEvents(days: number = 7, limit: number = 10): Promise<any[]> {
  console.error(`[getUpcomingEvents] Getting upcoming events for next ${days} days, limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set upcomingEvents to {}

      -- Find the calendar with the most events (main calendar)
      set allCals to every calendar
      set theCalendar to item 1 of allCals
      set maxEvents to 0
      repeat with cal in allCals
        set eventCount to count of (every calendar event of cal)
        if eventCount > maxEvents then
          set maxEvents to eventCount
          set theCalendar to cal
        end if
      end repeat

      set todayDate to current date
      set endDate to todayDate + ${days} * days
      set allEvents to every calendar event of theCalendar
      set limitCount to 0

      repeat with theEvent in allEvents
        set eventStart to start time of theEvent
        if eventStart >= todayDate and eventStart < endDate then
          set limitCount to limitCount + 1
          set eventData to {subject:subject of theEvent, ¬
                       start:eventStart, ¬
                       |end|:end time of theEvent, ¬
                       location:location of theEvent, ¬
                       id:id of theEvent}
          set end of upcomingEvents to eventData
          if limitCount >= ${limit} then
            exit repeat
          end if
        end if
      end repeat

      return upcomingEvents
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[getUpcomingEvents] Raw result length: ${result.length}`);

    // Parse the results - split by ", subject:" to separate events
    const events = [];
    if (result && result.includes('subject:')) {
      const eventStrings = result.split(/, subject:/).map((s, i) => i === 0 ? s : 'subject:' + s);

      for (const eventStr of eventStrings) {
        try {
          const event: any = {};
          // Parse key:value pairs, handling colons in values (like times)
          const keyValuePattern = /(subject|start|end|location|id):([^,]*(?:,[^a-z]*)?)/gi;
          let match;

          // Simple parsing: split by known keys
          const parts = eventStr.split(/, (?=subject:|start:|end:|location:|id:)/i);
          for (const part of parts) {
            const colonIdx = part.indexOf(':');
            if (colonIdx > 0) {
              const key = part.substring(0, colonIdx).trim().toLowerCase();
              const value = part.substring(colonIdx + 1).trim();
              if (key === 'end') {
                event['end'] = value;
              } else {
                event[key] = value;
              }
            }
          }

          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.start,
              end: event.end || event['|end|'],
              location: event.location || "No location",
              id: event.id
            });
          }
        } catch (parseError) {
          console.error('[getUpcomingEvents] Error parsing event:', parseError);
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
async function searchEvents(searchTerm: string, limit: number = 10): Promise<any[]> {
  console.error(`[searchEvents] Searching for events with term: "${searchTerm}", limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set searchResults to {}

      -- Find the calendar with the most events (main calendar)
      set allCals to every calendar
      set theCalendar to item 1 of allCals
      set maxEvents to 0
      repeat with cal in allCals
        set eventCount to count of (every calendar event of cal)
        if eventCount > maxEvents then
          set maxEvents to eventCount
          set theCalendar to cal
        end if
      end repeat

      set allEvents to every calendar event of theCalendar
      set i to 0
      set searchString to "${searchTerm.replace(/"/g, '\\"')}"

      repeat with theEvent in allEvents
        if (subject of theEvent contains searchString) or (location of theEvent contains searchString) then
          set i to i + 1
          set eventData to {subject:subject of theEvent, ¬
                       start:start time of theEvent, ¬
                       |end|:end time of theEvent, ¬
                       location:location of theEvent, ¬
                       id:id of theEvent}

          set end of searchResults to eventData

          -- Stop if we've reached the limit
          if i >= ${limit} then
            exit repeat
          end if
        end if
      end repeat

      return searchResults
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[searchEvents] Raw result length: ${result.length}`);

    // Parse the results - split by ", subject:" to separate events
    const events = [];
    if (result && result.includes('subject:')) {
      const eventStrings = result.split(/, subject:/).map((s, i) => i === 0 ? s : 'subject:' + s);

      for (const eventStr of eventStrings) {
        try {
          const event: any = {};
          const parts = eventStr.split(/, (?=subject:|start:|end:|location:|id:)/i);
          for (const part of parts) {
            const colonIdx = part.indexOf(':');
            if (colonIdx > 0) {
              const key = part.substring(0, colonIdx).trim().toLowerCase();
              const value = part.substring(colonIdx + 1).trim();
              event[key] = value;
            }
          }

          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.start,
              end: event.end,
              location: event.location || "No location",
              id: event.id
            });
          }
        } catch (parseError) {
          console.error('[searchEvents] Error parsing event:', parseError);
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
async function createEvent(subject: string, start: string, end: string, location?: string, body?: string, attendees?: string): Promise<string> {
  console.error(`[createEvent] Creating event: "${subject}", start: ${start}, end: ${end}`);
  await checkOutlookAccess();
  
  // Parse the ISO date strings to a format AppleScript can understand
  const startDate = new Date(start);
  const endDate = new Date(end);

  // Format for AppleScript: "M/D/YYYY H:MM:SS AM/PM" (12-hour format required)
  const formatForAppleScript = (d: Date): string => {
    const month = d.getMonth() + 1;
    const day = d.getDate();
    const year = d.getFullYear();
    let hours = d.getHours();
    const minutes = d.getMinutes().toString().padStart(2, '0');
    const seconds = d.getSeconds().toString().padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    if (hours === 0) hours = 12; // 0 becomes 12 for 12-hour format
    return `date "${month}/${day}/${year} ${hours}:${minutes}:${seconds} ${ampm}"`;
  };
  const formattedStart = formatForAppleScript(startDate);
  const formattedEnd = formatForAppleScript(endDate);
  
  // Escape strings for AppleScript
  const escapedSubject = subject.replace(/"/g, '\\"');
  const escapedLocation = location ? location.replace(/"/g, '\\"') : "";
  const escapedBody = body ? body.replace(/"/g, '\\"') : "";
  
  // Build properties object
  let properties = `subject:"${escapedSubject}", start time:${formattedStart}, end time:${formattedEnd}`;
  if (location) {
    properties += `, location:"${escapedLocation}"`;
  }
  if (body) {
    properties += `, content:"${escapedBody}"`;
  }

  let script = `
    tell application "Microsoft Outlook"
      -- Find the calendar with the most events (main calendar)
      set allCals to every calendar
      set theCalendar to item 1 of allCals
      set maxEvents to 0
      repeat with cal in allCals
        set eventCount to count of (every calendar event of cal)
        if eventCount > maxEvents then
          set maxEvents to eventCount
          set theCalendar to cal
        end if
      end repeat

      set newEvent to make new calendar event of theCalendar with properties {${properties}}
  `;
  
  // Add attendees if provided
  if (attendees) {
    const attendeeList = attendees.split(',').map(email => email.trim());
    
    for (const attendee of attendeeList) {
      const escapedAttendee = attendee.replace(/"/g, '\\"');
      script += `
        make new attendee at newEvent with properties {email address:"${escapedAttendee}"}
      `;
    }
  }
  
  script += `
      return "Event created: " & subject of newEvent
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

// Function to delete a calendar event by subject and date
async function deleteEvent(subject: string, dateStr: string): Promise<string> {
  console.error(`[deleteEvent] Deleting event: "${subject}" on ${dateStr}`);
  await checkOutlookAccess();

  // Parse the date string (YYYY-MM-DD) to get month, day, year
  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0]);
  const month = parseInt(dateParts[1]);
  const day = parseInt(dateParts[2]);

  // Escape the subject for AppleScript
  const escapedSubject = subject.replace(/"/g, '\\"');

  const script = `
    tell application "Microsoft Outlook"
      set targetDate to date "${month}/${day}/${year}"
      set startOfDay to targetDate
      set time of startOfDay to 0
      set endOfDay to targetDate + 1 * days

      set deletedCount to 0
      set allCals to every calendar

      repeat with cal in allCals
        set calEvents to (every calendar event of cal whose subject is "${escapedSubject}")
        repeat with evt in calEvents
          set evtStart to start time of evt
          if evtStart ≥ startOfDay and evtStart < endOfDay then
            delete evt
            set deletedCount to deletedCount + 1
          end if
        end repeat
      end repeat

      if deletedCount > 0 then
        return "Deleted " & deletedCount & " event(s) matching \\"${escapedSubject}\\" on ${month}/${day}/${year}"
      else
        return "No events found matching \\"${escapedSubject}\\" on ${month}/${day}/${year}"
      end if
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[deleteEvent] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[deleteEvent] Error deleting event:", error);
    throw error;
  }
}

// Function to respond to a meeting invite (accept, decline, tentative)
async function respondToMeeting(
  subject: string,
  dateStr: string,
  response: "accept" | "decline" | "tentative"
): Promise<string> {
  console.error(`[respondToMeeting] ${response} meeting: "${subject}" on ${dateStr}`);
  await checkOutlookAccess();

  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0]);
  const month = parseInt(dateParts[1]);
  const day = parseInt(dateParts[2]);

  const escapedSubject = subject.replace(/"/g, '\\"');

  // Map response to Outlook AppleScript command
  const responseCommand = response === "accept" ? "accept meeting"
    : response === "decline" ? "decline meeting"
    : "tentatively accept meeting";

  const script = `
    tell application "Microsoft Outlook"
      set targetDate to date "${month}/${day}/${year}"
      set startOfDay to targetDate
      set time of startOfDay to 0
      set endOfDay to targetDate + 1 * days

      set respondedCount to 0
      set allCals to every calendar

      repeat with cal in allCals
        set calEvents to (every calendar event of cal whose subject is "${escapedSubject}")
        repeat with evt in calEvents
          set evtStart to start time of evt
          if evtStart ≥ startOfDay and evtStart < endOfDay then
            try
              ${responseCommand} evt
              set respondedCount to respondedCount + 1
            on error errMsg
              -- Event might not be a meeting invite
              return "Error: " & errMsg
            end try
          end if
        end repeat
      end repeat

      if respondedCount > 0 then
        return "${response === "accept" ? "Accepted" : response === "decline" ? "Declined" : "Tentatively accepted"} " & respondedCount & " meeting(s) matching \\"${escapedSubject}\\" on ${month}/${day}/${year}"
      else
        return "No meetings found matching \\"${escapedSubject}\\" on ${month}/${day}/${year}"
      end if
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[respondToMeeting] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[respondToMeeting] Error responding to meeting:", error);
    throw error;
  }
}

// Function to propose a new time for a meeting
async function proposeNewTime(
  subject: string,
  dateStr: string,
  proposedStart: string,
  proposedEnd: string
): Promise<string> {
  console.error(`[proposeNewTime] Proposing new time for: "${subject}" on ${dateStr}`);
  await checkOutlookAccess();

  const dateParts = dateStr.split('-');
  const year = parseInt(dateParts[0]);
  const month = parseInt(dateParts[1]);
  const day = parseInt(dateParts[2]);

  // Parse the proposed times
  const startDate = new Date(proposedStart);
  const endDate = new Date(proposedEnd);

  // Format for AppleScript: "M/D/YYYY H:MM:SS AM/PM" (12-hour format required)
  const formatForAppleScript = (d: Date): string => {
    const m = d.getMonth() + 1;
    const dy = d.getDate();
    const yr = d.getFullYear();
    let hours = d.getHours();
    const minutes = d.getMinutes().toString().padStart(2, '0');
    const seconds = d.getSeconds().toString().padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    if (hours === 0) hours = 12;
    return `date "${m}/${dy}/${yr} ${hours}:${minutes}:${seconds} ${ampm}"`;
  };

  const formattedProposedStart = formatForAppleScript(startDate);
  const formattedProposedEnd = formatForAppleScript(endDate);

  const escapedSubject = subject.replace(/"/g, '\\"');

  const script = `
    tell application "Microsoft Outlook"
      set targetDate to date "${month}/${day}/${year}"
      set startOfDay to targetDate
      set time of startOfDay to 0
      set endOfDay to targetDate + 1 * days

      set proposedCount to 0
      set allCals to every calendar

      repeat with cal in allCals
        set calEvents to (every calendar event of cal whose subject is "${escapedSubject}")
        repeat with evt in calEvents
          set evtStart to start time of evt
          if evtStart ≥ startOfDay and evtStart < endOfDay then
            try
              propose new time evt proposed start time ${formattedProposedStart} proposed end time ${formattedProposedEnd}
              set proposedCount to proposedCount + 1
            on error errMsg
              return "Error: " & errMsg
            end try
          end if
        end repeat
      end repeat

      if proposedCount > 0 then
        return "Proposed new time for " & proposedCount & " meeting(s) matching \\"${escapedSubject}\\""
      else
        return "No meetings found matching \\"${escapedSubject}\\" on ${month}/${day}/${year}"
      end if
    end tell
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[proposeNewTime] Result: ${result}`);
    return result;
  } catch (error) {
    console.error("[proposeNewTime] Error proposing new time:", error);
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
        set searchString to "${searchTerm.replace(/"/g, '\\"')}"
        
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
            set searchString to "${searchTerm.replace(/"/g, '\\"')}"
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
  operation: "unread" | "search" | "send" | "draft" | "reply" | "forward" | "folders" | "read" | "create_folder" | "move_email" | "rename_folder" | "delete_folder" | "count" | "save_attachments";
  folder?: string;
  limit?: number;
  searchTerm?: string;
  startDate?: string;
  endDate?: string;
  to?: string;
  subject?: string;
  body?: string;
  isHtml?: boolean;
  cc?: string;
  bcc?: string;
  attachments?: string[];
  name?: string;
  parent?: string;
  messageId?: string;
  targetFolder?: string;
  destinationFolder?: string;
  newName?: string;
  forwardTo?: string;
  forwardCc?: string;
  forwardComment?: string;
  includeOriginalAttachments?: boolean;
  replyBody?: string;
  replyAll?: boolean;
} {
  if (typeof args !== "object" || args === null) return false;

  const { operation } = args as any;

  if (!operation || !["unread", "search", "send", "draft", "reply", "forward", "folders", "read", "create_folder", "move_email", "rename_folder", "delete_folder", "count", "save_attachments"].includes(operation)) {
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
    case "draft":
      if (!(args as any).to || !(args as any).subject || !(args as any).body) return false;
      break;
    case "reply":
      if (!(args as any).messageId || !(args as any).replyBody) return false;
      break;
    case "forward":
      if (!(args as any).messageId || !(args as any).forwardTo) return false;
      break;
    case "create_folder":
      if (!(args as any).name) return false;
      break;
    case "move_email":
      if (!(args as any).messageId || !(args as any).targetFolder) return false;
      break;
    case "rename_folder":
      if (!(args as any).folder || !(args as any).newName) return false;
      break;
    case "delete_folder":
      if (!(args as any).folder) return false;
      break;
    case "save_attachments":
      if (!(args as any).messageId || !(args as any).destinationFolder) return false;
      break;
  }

  return true;
}

function isCalendarArgs(args: unknown): args is {
  operation: "today" | "upcoming" | "search" | "create" | "delete" | "accept" | "decline" | "tentative" | "propose_new_time";
  searchTerm?: string;
  limit?: number;
  days?: number;
  subject?: string;
  start?: string;
  end?: string;
  location?: string;
  body?: string;
  attendees?: string;
  deleteSubject?: string;
  deleteDate?: string;
  responseSubject?: string;
  responseDate?: string;
  proposedStart?: string;
  proposedEnd?: string;
} {
  if (typeof args !== "object" || args === null) return false;

  const { operation } = args as any;

  if (!operation || !["today", "upcoming", "search", "create", "delete", "accept", "decline", "tentative", "propose_new_time"].includes(operation)) {
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
    case "delete":
      if (!(args as any).deleteSubject || !(args as any).deleteDate) return false;
      break;
    case "accept":
    case "decline":
    case "tentative":
      if (!(args as any).responseSubject || !(args as any).responseDate) return false;
      break;
    case "propose_new_time":
      if (!(args as any).responseSubject || !(args as any).responseDate ||
          !(args as any).proposedStart || !(args as any).proposedEnd) return false;
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
            const emails = await searchEmails(args.searchTerm, args.folder, args.limit, args.startDate, args.endDate);
            const dateRange = args.startDate || args.endDate
              ? ` (${args.startDate || 'any'} to ${args.endDate || 'any'})`
              : '';
            return {
              content: [{
                type: "text",
                text: emails.length > 0 ?
                  `Found ${emails.length} email(s) for "${args.searchTerm}"${dateRange}\n\n` +
                  emails.map((email, i) =>
                    `--- Email ${i + 1} ---\nID: ${email.messageId || 'Unknown'}\nFrom: ${email.sender}\nDate: ${email.dateSent}\nSubject: ${email.subject}\n\n${email.content}`
                  ).join("\n\n") :
                  `No emails found for "${args.searchTerm}"${dateRange}`
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

          case "draft": {
            if (!args.to || !args.subject || !args.body) {
              throw new Error("Recipient (to), subject, and body are required for draft operation");
            }

            // Validate attachments if provided
            if (args.attachments && !Array.isArray(args.attachments)) {
              throw new Error("Attachments must be an array of file paths");
            }

            console.error(`[CallTool] Create draft with attachments: ${args.attachments ? JSON.stringify(args.attachments) : 'none'}`);

            const draftResult = await createDraft(
              args.to,
              args.subject,
              args.body,
              args.cc,
              args.bcc,
              args.isHtml || false,
              args.attachments
            );

            return {
              content: [{ type: "text", text: draftResult }],
              isError: draftResult.startsWith("Error:")
            };
          }

          case "reply": {
            if (!args.messageId || !args.replyBody) {
              throw new Error("Message ID and replyBody are required for reply operation");
            }
            const result = await replyEmail(args.messageId, args.replyBody, args.replyAll || false, args.isHtml || false, args.attachments);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
            };
          }

          case "forward": {
            if (!args.messageId || !args.forwardTo) {
              throw new Error("Message ID and forwardTo are required for forward operation");
            }
            const result = await forwardEmail(
              args.messageId,
              args.forwardTo,
              args.forwardCc,
              args.forwardComment,
              args.attachments,
              args.includeOriginalAttachments !== undefined ? args.includeOriginalAttachments : true
            );
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
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
            console.error(`[MCP read] folder=${args.folder}, limit=${args.limit}, startDate=${args.startDate}, endDate=${args.endDate}`);
            const emails = await readEmails(args.folder, args.limit, args.startDate, args.endDate);
            const dateRange = args.startDate || args.endDate
              ? ` (${args.startDate || 'any'} to ${args.endDate || 'any'})`
              : '';
            return {
              content: [{
                type: "text",
                text: emails.length > 0 ?
                  `Found ${emails.length} email(s)${args.folder ? ` in folder "${args.folder}"` : ''}${dateRange}\n\n` +
                  emails.map(email =>
                    `[${email.dateSent}] ID: ${email.messageId || 'Unknown'}\nFrom: ${email.sender}\nSubject: ${email.subject}\n${email.content.substring(0, 200)}${email.content.length > 200 ? '...' : ''}`
                  ).join("\n\n") :
                  `No emails found${args.folder ? ` in folder "${args.folder}"` : ''}${dateRange}`
              }],
              isError: false
            };
          }

          case "create_folder": {
            if (!args.name) {
              throw new Error("Folder name is required for create_folder operation");
            }
            const result = await createFolder(args.name, args.parent);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
            };
          }

          case "move_email": {
            if (!args.messageId || !args.targetFolder) {
              throw new Error("Message ID and target folder are required for move_email operation");
            }
            const result = await moveEmail(args.messageId, args.targetFolder);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
            };
          }

          case "rename_folder": {
            if (!args.folder || !args.newName) {
              throw new Error("Folder and new name are required for rename_folder operation");
            }
            const result = await renameFolder(args.folder, args.newName);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
            };
          }

          case "delete_folder": {
            if (!args.folder) {
              throw new Error("Folder is required for delete_folder operation");
            }
            const result = await deleteFolder(args.folder);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
            };
          }

          case "count": {
            const folder = args.folder || "Inbox";
            const result = await countEmails(folder);
            return {
              content: [{ type: "text", text: `${folder}: ${result}` }],
              isError: result.startsWith("Error:")
            };
          }

          case "save_attachments": {
            if (!args.messageId) {
              throw new Error("messageId is required for save_attachments operation");
            }
            if (!args.destinationFolder) {
              throw new Error("destinationFolder is required for save_attachments operation");
            }
            const result = await saveAttachments(args.messageId, args.destinationFolder);
            return {
              content: [{ type: "text", text: result }],
              isError: result.startsWith("Error:")
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

          case "delete": {
            if (!args.deleteSubject || !args.deleteDate) {
              throw new Error("deleteSubject and deleteDate are required for delete operation");
            }
            const deleteResult = await deleteEvent(args.deleteSubject, args.deleteDate);
            return {
              content: [{ type: "text", text: deleteResult }],
              isError: false
            };
          }

          case "accept": {
            if (!args.responseSubject || !args.responseDate) {
              throw new Error("responseSubject and responseDate are required for accept operation");
            }
            const acceptResult = await respondToMeeting(args.responseSubject, args.responseDate, "accept");
            return {
              content: [{ type: "text", text: acceptResult }],
              isError: false
            };
          }

          case "decline": {
            if (!args.responseSubject || !args.responseDate) {
              throw new Error("responseSubject and responseDate are required for decline operation");
            }
            const declineResult = await respondToMeeting(args.responseSubject, args.responseDate, "decline");
            return {
              content: [{ type: "text", text: declineResult }],
              isError: false
            };
          }

          case "tentative": {
            if (!args.responseSubject || !args.responseDate) {
              throw new Error("responseSubject and responseDate are required for tentative operation");
            }
            const tentativeResult = await respondToMeeting(args.responseSubject, args.responseDate, "tentative");
            return {
              content: [{ type: "text", text: tentativeResult }],
              isError: false
            };
          }

          case "propose_new_time": {
            if (!args.responseSubject || !args.responseDate || !args.proposedStart || !args.proposedEnd) {
              throw new Error("responseSubject, responseDate, proposedStart, and proposedEnd are required for propose_new_time operation");
            }
            const proposeResult = await proposeNewTime(args.responseSubject, args.responseDate, args.proposedStart, args.proposedEnd);
            return {
              content: [{ type: "text", text: proposeResult }],
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
// 9. EXPORTS FOR TESTING
// ====================================================

// Export functions for integration testing
export {
  readEmails,
  searchEmails,
  getUnreadEmails,
  sendEmail,
  createDraft,
  replyEmail,
  forwardEmail,
  getMailFolders,
  createFolder,
  moveEmail,
  renameFolder,
  deleteFolder,
  getTodayEvents,
  getUpcomingEvents,
  searchEvents,
  createEvent,
  listContacts,
  searchContacts
};

// ====================================================
// 10. START SERVER
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