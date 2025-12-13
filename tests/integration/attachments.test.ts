import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_RECIPIENT, TEST_TIMEOUT } from "../setup";
import { sendEmail, replyEmail, forwardEmail, readEmails } from "../../index";
import { writeFileSync, unlinkSync, existsSync, mkdirSync } from "fs";
import { join } from "path";

// Test fixtures directory
const TEST_FIXTURES_DIR = join(__dirname, "../fixtures");
const TEST_ATTACHMENT_PATH = join(TEST_FIXTURES_DIR, "test-attachment.txt");
const TEST_PDF_PATH = join(TEST_FIXTURES_DIR, "test-document.pdf");

describe("Attachments", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();

    // Create fixtures directory if it doesn't exist
    if (!existsSync(TEST_FIXTURES_DIR)) {
      mkdirSync(TEST_FIXTURES_DIR, { recursive: true });
    }

    // Create test attachment files
    writeFileSync(TEST_ATTACHMENT_PATH, "This is a test attachment file for Outlook MCP tests.\nGenerated at: " + new Date().toISOString());

    // Create a minimal PDF-like file (not a real PDF, but good enough for attachment testing)
    writeFileSync(TEST_PDF_PATH, "%PDF-1.4\nTest PDF content for attachment testing\n%%EOF");
  });

  describe("sendEmail with attachments", () => {
    it("sends email with single attachment", async () => {
      const subject = `[TEST] Send with Attachment - ${Date.now()}`;
      const result = await sendEmail(
        TEST_RECIPIENT,
        subject,
        "This test email includes a single attachment.",
        undefined, // cc
        undefined, // bcc
        false,     // isHtml
        [TEST_ATTACHMENT_PATH]  // attachments
      );

      expect(result).toContain("success");
      expect(result.toLowerCase()).toContain("attachment");
    }, TEST_TIMEOUT);

    it("sends email with multiple attachments", async () => {
      const subject = `[TEST] Send with Multiple Attachments - ${Date.now()}`;
      const result = await sendEmail(
        TEST_RECIPIENT,
        subject,
        "This test email includes multiple attachments.",
        undefined,
        undefined,
        false,
        [TEST_ATTACHMENT_PATH, TEST_PDF_PATH]
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);

    it("sends email without attachments (backward compatibility)", async () => {
      const subject = `[TEST] Send without Attachment - ${Date.now()}`;
      const result = await sendEmail(
        TEST_RECIPIENT,
        subject,
        "This test email has no attachments.",
        undefined,
        undefined,
        false
        // No attachments parameter
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);

    it("handles non-existent attachment gracefully", async () => {
      const subject = `[TEST] Send with Missing Attachment - ${Date.now()}`;
      const result = await sendEmail(
        TEST_RECIPIENT,
        subject,
        "This test email references a non-existent file.",
        undefined,
        undefined,
        false,
        ["/non/existent/file.txt"]
      );

      // Should still send (with warning logged), attachment just won't be included
      expect(result).toContain("success");
    }, TEST_TIMEOUT);
  });

  describe("replyEmail with attachments", () => {
    it("replies to email with attachment", async () => {
      // First, find an email to reply to
      const emails = await readEmails("Inbox", 1);

      // readEmails returns an array of email objects
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test reply with attachment");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await replyEmail(
        messageId,
        "[TEST] Reply with attachment - please ignore",
        false, // replyAll
        [TEST_ATTACHMENT_PATH]
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);

    it("replies without attachment (backward compatibility)", async () => {
      const emails = await readEmails("Inbox", 1);
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test reply");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await replyEmail(
        messageId,
        "[TEST] Reply without attachment - please ignore",
        false
        // No attachments
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);
  });

  describe("forwardEmail with attachments", () => {
    it("forwards email with additional attachment (keeping originals)", async () => {
      // Find an email to forward
      const emails = await readEmails("Inbox", 1);
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test forward with attachment");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await forwardEmail(
        messageId,
        TEST_RECIPIENT,
        undefined, // cc
        "[TEST] Forward with attachment (keeping originals) - please ignore",
        [TEST_ATTACHMENT_PATH],
        true // includeOriginalAttachments
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);

    it("forwards email with new attachment, removing originals", async () => {
      const emails = await readEmails("Inbox", 1);
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test forward");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await forwardEmail(
        messageId,
        TEST_RECIPIENT,
        undefined,
        "[TEST] Forward with new attachment only - please ignore",
        [TEST_ATTACHMENT_PATH],
        false // Remove original attachments
      );

      expect(result).toContain("success");
      expect(result).toContain("original attachments removed");
    }, TEST_TIMEOUT);

    it("forwards without additional attachment (backward compatibility)", async () => {
      const emails = await readEmails("Inbox", 1);
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test forward");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await forwardEmail(
        messageId,
        TEST_RECIPIENT,
        undefined,
        "[TEST] Forward without attachment - please ignore"
        // No attachments, default includeOriginalAttachments=true
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);

    it("forwards removing original attachments without adding new ones", async () => {
      const emails = await readEmails("Inbox", 1);
      if (!emails || emails.length === 0 || !emails[0].id) {
        console.log("No emails found in Inbox to test forward");
        return;
      }

      const messageId = String(emails[0].id);
      const result = await forwardEmail(
        messageId,
        TEST_RECIPIENT,
        undefined,
        "[TEST] Forward with no attachments at all - please ignore",
        undefined, // No new attachments
        false // Remove original attachments
      );

      expect(result).toContain("success");
    }, TEST_TIMEOUT);
  });
});
