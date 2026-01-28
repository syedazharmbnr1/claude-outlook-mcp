#!/usr/bin/env bun
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type Tool,
} from "@modelcontextprotocol/sdk/types.js";
import { runAppleScript } from 'run-applescript';

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
  
  const script = `
    tell application "Microsoft Outlook"
      try
        -- Get the folder by name safely (handle localization)
        set targetFolder to null
        set allFolders to mail folders
        
        -- First try exact match
        repeat with mailFolder in allFolders
          if name of mailFolder is "${folder}" then
            set targetFolder to mailFolder
            exit repeat
          end if
        end repeat
        
        -- If not found and looking for Inbox, try common localizations
        if targetFolder is null and ("${folder}" is "Inbox" or "${folder}" is "inbox") then
          repeat with mailFolder in allFolders
            set folderName to name of mailFolder
            if folderName is "Posteingang" or folderName is "Inbox" or folderName is "Boîte de réception" or folderName is "Bandeja de entrada" or folderName is "Posta in arrivo" or folderName is "Postvak IN" or folderName is "Caixa de Entrada" or folderName is "Indbakke" or folderName is "Innboks" or folderName is "Skrzynka odbiorcza" then
              set targetFolder to mailFolder
              exit repeat
            end if
          end repeat
        end if
        
        -- Last resort: use AppleScript inbox keyword
        if targetFolder is null then set targetFolder to inbox
        
        set outputText to ""
        set allMessages to messages of targetFolder
        set i to 0
        
        repeat with theMessage in allMessages
          if i >= ${limit} then exit repeat
          
          if is read of theMessage is false then
            try
              set msgSubject to subject of theMessage
              
              -- Get sender info properly
              set senderObj to sender of theMessage
              try
                set senderName to name of senderObj
                set senderAddr to address of senderObj
                set msgSender to senderName & " <" & senderAddr & ">"
              on error
                set msgSender to senderObj as string
              end try
              
              set msgDate to time sent of theMessage as string
              
              -- Get content preview
              set msgContent to ""
              try
                set msgContent to plain text content of theMessage
                if length of msgContent > 300 then
                  set msgContent to text 1 thru 300 of msgContent
                end if
              on error
                try
                  set msgContent to content of theMessage
                  if length of msgContent > 300 then
                    set msgContent to text 1 thru 300 of msgContent
                  end if
                on error
                  set msgContent to "[Content not available]"
                end try
              end try
              
              set outputText to outputText & "<<<EMAIL>>>" & msgSubject & "<<<FIELD>>>" & msgSender & "<<<FIELD>>>" & msgDate & "<<<FIELD>>>" & msgContent
              set i to i + 1
            on error
              -- Skip problematic messages
            end try
          end if
        end repeat
        
        return outputText
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
    
    if (!result || result.trim() === "") {
      return [];
    }
    
    // Parse using our custom delimiters
    const emailStrings = result.split("<<<EMAIL>>>").filter(s => s.trim() !== "");
    const emails = emailStrings.map(emailStr => {
      const parts = emailStr.split("<<<FIELD>>>");
      return {
        subject: parts[0] || "No subject",
        sender: parts[1] || "Unknown sender",
        dateSent: parts[2] || new Date().toString(),
        content: parts[3] || "[Content not available]"
      };
    });
    
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
  
  const script = `
    tell application "Microsoft Outlook"
      try
        -- Get the folder by name safely (handle localization)
        set targetFolder to null
        set allFolders to mail folders
        
        -- First try exact match
        repeat with mailFolder in allFolders
          if name of mailFolder is "${folder}" then
            set targetFolder to mailFolder
            exit repeat
          end if
        end repeat
        
        -- If not found and looking for Inbox, try common localizations
        if targetFolder is null and ("${folder}" is "Inbox" or "${folder}" is "inbox") then
          repeat with mailFolder in allFolders
            set folderName to name of mailFolder
            if folderName is "Posteingang" or folderName is "Inbox" or folderName is "Boîte de réception" or folderName is "Bandeja de entrada" or folderName is "Posta in arrivo" or folderName is "Postvak IN" or folderName is "Caixa de Entrada" or folderName is "Indbakke" or folderName is "Innboks" or folderName is "Skrzynka odbiorcza" then
              set targetFolder to mailFolder
              exit repeat
            end if
          end repeat
        end if
        
        -- Last resort: use AppleScript inbox keyword
        if targetFolder is null then set targetFolder to inbox
        
        set outputText to ""
        set allMessages to messages of targetFolder
        set i to 0
        set searchString to "${searchTerm.replace(/"/g, '\\"')}"
        
        repeat with theMessage in allMessages
          if i >= ${limit} then exit repeat
          
          try
            set msgSubject to subject of theMessage
            set msgContentRaw to ""
            try
              set msgContentRaw to plain text content of theMessage
            on error
              try
                set msgContentRaw to content of theMessage
              end try
            end try
            
            if (msgSubject contains searchString) or (msgContentRaw contains searchString) then
              -- Get sender info properly
              set senderObj to sender of theMessage
              try
                set senderName to name of senderObj
                set senderAddr to address of senderObj
                set msgSender to senderName & " <" & senderAddr & ">"
              on error
                set msgSender to senderObj as string
              end try
              
              set msgDate to time sent of theMessage as string
              
              -- Get content preview
              set msgContent to msgContentRaw
              if length of msgContent > 300 then
                set msgContent to text 1 thru 300 of msgContent
              end if
              if msgContent is "" then set msgContent to "[Content not available]"
              
              set outputText to outputText & "<<<EMAIL>>>" & msgSubject & "<<<FIELD>>>" & msgSender & "<<<FIELD>>>" & msgDate & "<<<FIELD>>>" & msgContent
              set i to i + 1
            end if
          on error
            -- Skip problematic messages
          end try
        end repeat
        
        return outputText
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
    
    if (!result || result.trim() === "") {
      return [];
    }
    
    // Parse using our custom delimiters
    const emailStrings = result.split("<<<EMAIL>>>").filter(s => s.trim() !== "");
    const emails = emailStrings.map(emailStr => {
      const parts = emailStr.split("<<<FIELD>>>");
      return {
        subject: parts[0] || "No subject",
        sender: parts[1] || "Unknown sender",
        dateSent: parts[2] || new Date().toString(),
        content: parts[3] || "[Content not available]"
      };
    });
    
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
    if (!filePath.startsWith('/')) {
      const cwd = process.cwd();
      fullPath = `${cwd}/${filePath}`;
    }
    
    // Check if the file exists and is readable
    const fs = require('fs');
    const { promisify } = require('util');
    const access = promisify(fs.access);
    const stat = promisify(fs.stat);
    
    try {
      await access(fullPath, fs.constants.R_OK);
      const stats = await stat(fullPath);
      
      return `File exists and is readable: ${fullPath}\nSize: ${stats.size} bytes\nPermissions: ${stats.mode.toString(8)}\nLast modified: ${stats.mtime}`;
    } catch (err) {
      return `ERROR: Cannot access file: ${fullPath}\nError details: ${err.message}`;
    }
  } catch (error) {
    return `Failed to check attachment path: ${error.message}`;
  }
}

// Add a debug version of sending email with attachment to test if files are accessible
async function debugSendEmailWithAttachment(
  to: string,
  subject: string,
  body: string,
  attachmentPath: string
): Promise<string> {
  // First check if the file exists and is readable
  const fileStatus = await checkAttachmentPath(attachmentPath);
  console.error(`[debugSendEmail] Attachment status: ${fileStatus}`);
  
  // Create a simple AppleScript that just attempts to open the file
  const script = `
    set theFile to POSIX file "${attachmentPath.replace(/"/g, '\\"')}"
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
          set newMessage to make new outgoing message with properties {subject:"DEBUG: ${subject.replace(/"/g, '\\"')}", visible:true}
          set content of newMessage to "${body.replace(/"/g, '\\"')}"
          set to recipients of newMessage to {"${to}"}
          
          try
            set attachmentFile to POSIX file "${attachmentPath.replace(/"/g, '\\"')}"
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
    return `Debugging error: ${error.message}\n\nFile check: ${fileStatus}`;
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
  const escapedSubject = subject.replace(/"/g, '\\"');
  const escapedBody = body.replace(/"/g, '\\"').replace(/\n/g, '\\n');
  
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
            `set content type of msg to HTML
             set content of msg to "${escapedBody}"` 
          : 
            `set content of msg to "${escapedBody}"`
          }
          
          tell msg
            set recipTo to make new to recipient with properties {email address:{name:"${toName}", address:"${to}"}}
            ${cc ? `set recipCc to make new cc recipient with properties {email address:{name:"${ccName}", address:"${cc}"}}` : ''}
            ${bcc ? `set recipBcc to make new bcc recipient with properties {email address:{name:"${bccName}", address:"${bcc}"}}` : ''}
            
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
                `set content type of newMessage to HTML
                 set content of newMessage to "${escapedBody}"` 
              : 
                `set content of newMessage to "${escapedBody}"`
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
    
    const script = `
      tell application "Microsoft Outlook"
        try
          -- Get the folder by name safely (handle localization)
          set targetFolder to null
          set allFolders to mail folders
          
          -- First try exact match
          repeat with mailFolder in allFolders
            if name of mailFolder is "${folder}" then
              set targetFolder to mailFolder
              exit repeat
            end if
          end repeat
          
          -- If not found and looking for Inbox, try common localizations
          if targetFolder is null and ("${folder}" is "Inbox" or "${folder}" is "inbox") then
            repeat with mailFolder in allFolders
              set folderName to name of mailFolder
              if folderName is "Posteingang" or folderName is "Inbox" or folderName is "Boîte de réception" or folderName is "Bandeja de entrada" or folderName is "Posta in arrivo" or folderName is "Postvak IN" or folderName is "Caixa de Entrada" or folderName is "Indbakke" or folderName is "Innboks" or folderName is "Skrzynka odbiorcza" then
                set targetFolder to mailFolder
                exit repeat
              end if
            end repeat
          end if
          
          -- Last resort: use AppleScript inbox keyword
          if targetFolder is null then set targetFolder to inbox
          
          -- Get messages
          set outputText to ""
          set msgCount to 0
          set allMsgs to messages of targetFolder
          
          repeat with i from 1 to (count of allMsgs)
            if msgCount >= ${limit} then exit repeat
            
            try
              set theMsg to item i of allMsgs
              set msgSubject to subject of theMsg
              
              -- Get sender info properly
              set senderObj to sender of theMsg
              try
                set senderName to name of senderObj
                set senderAddr to address of senderObj
                set msgSender to senderName & " <" & senderAddr & ">"
              on error
                set msgSender to senderObj as string
              end try
              
              set msgDate to time sent of theMsg as string
              
              -- Get content preview
              set msgContent to ""
              try
                set msgContent to plain text content of theMsg
                if length of msgContent > 300 then
                  set msgContent to text 1 thru 300 of msgContent
                end if
              on error
                try
                  set msgContent to content of theMsg
                  if length of msgContent > 300 then
                    set msgContent to text 1 thru 300 of msgContent
                  end if
                on error
                  set msgContent to "[Content not available]"
                end try
              end try
              
              -- Use a unique delimiter that won't appear in email content
              set outputText to outputText & "<<<EMAIL>>>" & msgSubject & "<<<FIELD>>>" & msgSender & "<<<FIELD>>>" & msgDate & "<<<FIELD>>>" & msgContent
              set msgCount to msgCount + 1
            on error errMsg
              -- Skip problematic messages
            end try
          end repeat
          
          return outputText
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
      
      if (!result || result.trim() === "") {
        return [];
      }
      
      // Parse using our custom delimiters
      const emailStrings = result.split("<<<EMAIL>>>").filter(s => s.trim() !== "");
      const emails = emailStrings.map(emailStr => {
        const parts = emailStr.split("<<<FIELD>>>");
        return {
          subject: parts[0] || "No subject",
          sender: parts[1] || "Unknown sender",
          dateSent: parts[2] || new Date().toString(),
          content: parts[3] || "[Content not available]"
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
// 5. CALENDAR FUNCTIONS
// ====================================================

// Function to get today's calendar events
async function getTodayEvents(limit: number = 10): Promise<any[]> {
  console.error(`[getTodayEvents] Getting today's events, limit: ${limit}`);
  await checkOutlookAccess();
  
  const script = `
    tell application "Microsoft Outlook"
      set todayEvents to {}
      set theCalendar to default calendar
      set todayDate to current date
      set startOfDay to todayDate - (time of todayDate)
      set endOfDay to startOfDay + 1 * days
      
      set eventList to events of theCalendar whose start time is greater than or equal to startOfDay and start time is less than endOfDay
      
      set eventCount to count of eventList
      set limitCount to ${limit}
      
      if eventCount < limitCount then
        set limitCount to eventCount
      end if
      
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set eventSubject to subject of theEvent
        set eventStart to start time of theEvent
        set eventEnd to end time of theEvent
        set eventLocation to location of theEvent
        set eventId to id of theEvent
        set eventData to {subject:eventSubject, startTime:eventStart, endTime:eventEnd, location:eventLocation, eventId:eventId}
        
        set end of todayEvents to eventData
      end repeat
      
      return todayEvents
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[getTodayEvents] Raw result length: ${result.length}`);
    
    // Parse the results
    const events = [];
    const matches = result.match(/\{([^}]+)\}/g);
    
    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(',');
          const event: any = {};
          
          props.forEach(prop => {
            const parts = prop.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(':').trim();
              event[key] = value;
            }
          });
          
          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.startTime || event.start,
              end: event.endTime || event.end,
              location: event.location || "No location",
              id: event.eventId || event.id
            });
          }
        } catch (parseError) {
          console.error('[getTodayEvents] Error parsing event match:', parseError);
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
      set theCalendar to default calendar
      set todayDate to current date
      set startOfToday to todayDate - (time of todayDate)
      set endDate to startOfToday + ${days} * days
      
      set eventList to events of theCalendar whose start time is greater than or equal to todayDate and start time is less than endDate
      
      set eventCount to count of eventList
      set limitCount to ${limit}
      
      if eventCount < limitCount then
        set limitCount to eventCount
      end if
      
      repeat with i from 1 to limitCount
        set theEvent to item i of eventList
        set eventSubject to subject of theEvent
        set eventStart to start time of theEvent
        set eventEnd to end time of theEvent
        set eventLocation to location of theEvent
        set eventId to id of theEvent
        set eventData to {subject:eventSubject, startTime:eventStart, endTime:eventEnd, location:eventLocation, eventId:eventId}
        
        set end of upcomingEvents to eventData
      end repeat
      
      return upcomingEvents
    end tell
  `;
  
  try {
    const result = await runAppleScript(script);
    console.error(`[getUpcomingEvents] Raw result length: ${result.length}`);
    
    // Parse the results
    const events = [];
    const matches = result.match(/\{([^}]+)\}/g);
    
    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(',');
          const event: any = {};
          
          props.forEach(prop => {
            const parts = prop.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(':').trim();
              event[key] = value;
            }
          });
          
          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.startTime || event.start,
              end: event.endTime || event.end,
              location: event.location || "No location",
              id: event.eventId || event.id
            });
          }
        } catch (parseError) {
          console.error('[getUpcomingEvents] Error parsing event match:', parseError);
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
      set theCalendar to default calendar
      set allEvents to events of theCalendar
      set i to 0
      set searchString to "${searchTerm.replace(/"/g, '\\"')}"
      
      repeat with theEvent in allEvents
        if (subject of theEvent contains searchString) or (location of theEvent contains searchString) then
          set i to i + 1
          set eventSubject to subject of theEvent
          set eventStart to start time of theEvent
          set eventEnd to end time of theEvent
          set eventLocation to location of theEvent
          set eventId to id of theEvent
          set eventData to {subject:eventSubject, startTime:eventStart, endTime:eventEnd, location:eventLocation, eventId:eventId}
          
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
    
    // Parse the results
    const events = [];
    const matches = result.match(/\{([^}]+)\}/g);
    
    if (matches && matches.length > 0) {
      for (const match of matches) {
        try {
          const props = match.substring(1, match.length - 1).split(',');
          const event: any = {};
          
          props.forEach(prop => {
            const parts = prop.split(':');
            if (parts.length >= 2) {
              const key = parts[0].trim();
              const value = parts.slice(1).join(':').trim();
              event[key] = value;
            }
          });
          
          if (event.subject) {
            events.push({
              subject: event.subject,
              start: event.startTime || event.start,
              end: event.endTime || event.end,
              location: event.location || "No location",
              id: event.eventId || event.id
            });
          }
        } catch (parseError) {
          console.error('[searchEvents] Error parsing event match:', parseError);
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
  
  // Format for AppleScript (month/day/year hour:minute:second)
  const formattedStart = `date "${startDate.getMonth() + 1}/${startDate.getDate()}/${startDate.getFullYear()} ${startDate.getHours()}:${startDate.getMinutes()}:${startDate.getSeconds()}"`;
  const formattedEnd = `date "${endDate.getMonth() + 1}/${endDate.getDate()}/${endDate.getFullYear()} ${endDate.getHours()}:${endDate.getMinutes()}:${endDate.getSeconds()}"`;
  
  // Escape strings for AppleScript
  const escapedSubject = subject.replace(/"/g, '\\"');
  const escapedLocation = location ? location.replace(/"/g, '\\"') : "";
  const escapedBody = body ? body.replace(/"/g, '\\"') : "";
  
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
      const escapedAttendee = attendee.replace(/"/g, '\\"');
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