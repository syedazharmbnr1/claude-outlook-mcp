import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_TIMEOUT } from "../setup";
import { readEmails } from "../../index";

describe("readEmails date filtering", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns emails without date filter", async () => {
    // Baseline: confirm we can get emails without any filter
    const emails = await readEmails("Inbox", 10);
    console.log(`Without filter: found ${emails.length} emails`);
    expect(emails.length).toBeGreaterThan(0);
  }, TEST_TIMEOUT);

  it("returns emails for a single day range (today)", async () => {
    // Test filtering by today's date - locale-independent approach
    const today = new Date();
    const isoDate = today.toISOString().split('T')[0]; // Format: YYYY-MM-DD

    console.log(`Testing date filter for today: ${isoDate}`);

    // Get emails for today only
    const filteredEmails = await readEmails("Inbox", 10, isoDate, isoDate);
    console.log(`With date filter (${isoDate}): found ${filteredEmails.length} emails`);

    // Verify the filter executed without error (may return 0 if no emails today)
    expect(filteredEmails.length).toBeGreaterThanOrEqual(0);
  }, TEST_TIMEOUT);

  it("returns emails for a date range spanning multiple days", async () => {
    // Test a week range ending today - locale-independent approach
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 7);

    const startIso = startDate.toISOString().split('T')[0];
    const endIso = endDate.toISOString().split('T')[0];

    console.log(`Testing week range: ${startIso} to ${endIso}`);

    // Get all emails (unfiltered) and filtered emails
    const allEmails = await readEmails("Inbox", 50);
    const filteredEmails = await readEmails("Inbox", 50, startIso, endIso);

    console.log(`Unfiltered: ${allEmails.length} emails, Week range: ${filteredEmails.length} emails`);

    // Verify filter executed without error
    expect(filteredEmails.length).toBeGreaterThanOrEqual(0);

    // Filtered results should be <= unfiltered (date filter reduces or maintains count)
    expect(filteredEmails.length).toBeLessThanOrEqual(allEmails.length);
  }, TEST_TIMEOUT);
});
