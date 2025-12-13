import { describe, expect, it } from "bun:test";
import { buildFolderRef, buildNestedFolderRef, escapeForAppleScript } from "../../helpers";

describe("buildFolderRef", () => {
  it("returns 'inbox' for Inbox folder", () => {
    expect(buildFolderRef("Inbox")).toBe("inbox");
  });

  it("returns mail folder reference for other folders", () => {
    expect(buildFolderRef("Archive")).toBe('mail folder "Archive"');
    expect(buildFolderRef("Sent Items")).toBe('mail folder "Sent Items"');
    expect(buildFolderRef("Drafts")).toBe('mail folder "Drafts"');
  });

  it("handles folders with special characters", () => {
    expect(buildFolderRef("My Folder")).toBe('mail folder "My Folder"');
    expect(buildFolderRef("Work/Projects")).toBe('mail folder "Work/Projects"');
  });

  it("is case-sensitive for Inbox", () => {
    // Only exact "Inbox" should use inbox reference
    expect(buildFolderRef("inbox")).toBe('mail folder "inbox"');
    expect(buildFolderRef("INBOX")).toBe('mail folder "INBOX"');
  });

  it("handles empty string", () => {
    expect(buildFolderRef("")).toBe('mail folder ""');
  });
});

describe("buildNestedFolderRef", () => {
  it("returns 'inbox' for Inbox folder", () => {
    expect(buildNestedFolderRef("Inbox")).toBe("inbox");
  });

  it("returns mail folder reference for flat folders", () => {
    expect(buildNestedFolderRef("Archive")).toBe('mail folder "Archive"');
    expect(buildNestedFolderRef("Reports")).toBe('mail folder "Reports"');
  });

  it("builds nested reference for paths with slashes", () => {
    expect(buildNestedFolderRef("Work/Reports")).toBe('mail folder "Reports" of mail folder "Work"');
  });

  it("handles deeply nested paths", () => {
    expect(buildNestedFolderRef("A/B/C")).toBe('mail folder "C" of mail folder "B" of mail folder "A"');
  });

  it("handles folders with spaces in nested paths", () => {
    expect(buildNestedFolderRef("My Programs/My Project")).toBe('mail folder "My Project" of mail folder "My Programs"');
  });
});

describe("escapeForAppleScript", () => {
  it("escapes double quotes", () => {
    expect(escapeForAppleScript('He said "hello"')).toBe('He said \\"hello\\"');
  });

  it("escapes backslashes", () => {
    expect(escapeForAppleScript("path\\to\\file")).toBe("path\\\\to\\\\file");
  });

  it("escapes both quotes and backslashes", () => {
    expect(escapeForAppleScript('path\\with"quotes')).toBe('path\\\\with\\"quotes');
  });

  it("returns unchanged string when no escaping needed", () => {
    expect(escapeForAppleScript("Hello world")).toBe("Hello world");
    expect(escapeForAppleScript("test@email.com")).toBe("test@email.com");
  });

  it("handles empty string", () => {
    expect(escapeForAppleScript("")).toBe("");
  });

  it("handles multiple consecutive special characters", () => {
    // Input: """\\  (three quotes and two backslashes)
    // Output: \"\"\"\\\\  (escaped quotes and backslashes)
    expect(escapeForAppleScript('"""' + "\\\\" + "")).toBe('\\"\\"\\"\\\\\\\\');
  });
});
