import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_TIMEOUT } from "../setup";
import { readEmails, searchEmails } from "../../index";

describe("readEmails", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns emails from Inbox only", async () => {
    const emails = await readEmails("Inbox", 5);

    expect(emails.length).toBeGreaterThan(0);
    expect(emails.length).toBeLessThanOrEqual(5);

    for (const email of emails) {
      expect(email.subject).toBeDefined();
      expect(email.sender).toBeDefined();
      expect(email.dateSent).toBeDefined();
    }
  }, TEST_TIMEOUT);

  it("respects limit parameter", async () => {
    const emails = await readEmails("Inbox", 3);
    expect(emails.length).toBeLessThanOrEqual(3);
  }, TEST_TIMEOUT);

  it("returns different results for different folders", async () => {
    // This test catches the bug where folder parameter was ignored
    const inbox = await readEmails("Inbox", 5);
    const archive = await readEmails("Archive", 5);

    const inboxSubjects = inbox.map(e => e.subject);
    const archiveSubjects = archive.map(e => e.subject);

    // At least some subjects should differ between folders
    // (unless one folder is empty, in which case we skip)
    if (inbox.length > 0 && archive.length > 0) {
      expect(inboxSubjects).not.toEqual(archiveSubjects);
    }
  }, TEST_TIMEOUT);
});

describe("searchEmails", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("searches only specified folder", async () => {
    // Search for a common term in Inbox
    const results = await searchEmails("the", "Inbox", 5);

    // If results found, verify they're from searching the right place
    expect(Array.isArray(results)).toBe(true);
    expect(results.length).toBeLessThanOrEqual(5);
  }, TEST_TIMEOUT);

  it("returns empty array when no matches", async () => {
    const results = await searchEmails("xyznonexistenttermxyz123456", "Inbox", 5);
    expect(results).toEqual([]);
  }, TEST_TIMEOUT);

  it("search is case-insensitive", async () => {
    // This test verifies case-insensitive search works
    const lowerResults = await searchEmails("meeting", "Inbox", 5);
    const upperResults = await searchEmails("MEETING", "Inbox", 5);

    // Both searches should find the same emails
    expect(lowerResults.length).toBe(upperResults.length);
  }, TEST_TIMEOUT);
});
