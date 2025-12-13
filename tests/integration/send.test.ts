import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_RECIPIENT, TEST_TIMEOUT } from "../setup";
import { sendEmail } from "../../index";

describe("sendEmail", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("sends email successfully", async () => {
    const subject = `[TEST] Outlook MCP Test - ${Date.now()}`;
    const result = await sendEmail(
      TEST_RECIPIENT,
      subject,
      "This is an automated test email from Outlook MCP plugin tests. Safe to delete.",
      undefined, // cc
      undefined, // bcc
      false      // isHtml
    );

    expect(result).toContain("success");
  }, TEST_TIMEOUT);

  it("sends email with CC", async () => {
    const subject = `[TEST] Outlook MCP CC Test - ${Date.now()}`;
    const result = await sendEmail(
      TEST_RECIPIENT,
      subject,
      "This is a test email with CC. Safe to delete.",
      TEST_RECIPIENT, // CC to same recipient for testing
      undefined,
      false
    );

    expect(result).toContain("success");
  }, TEST_TIMEOUT);

  it("sends HTML email", async () => {
    const subject = `[TEST] Outlook MCP HTML Test - ${Date.now()}`;
    const result = await sendEmail(
      TEST_RECIPIENT,
      subject,
      "<html><body><h1>Test</h1><p>This is an <b>HTML</b> test email.</p></body></html>",
      undefined,
      undefined,
      true // isHtml
    );

    expect(result).toContain("success");
  }, TEST_TIMEOUT);
});
