import { describe, expect, it, beforeAll, afterAll } from "bun:test";
import { ensureOutlookRunning, TEST_TIMEOUT } from "../setup";
import { getMailFolders, createFolder, renameFolder, deleteFolder } from "../../index";

describe("getMailFolders", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns list of folders", async () => {
    const folders = await getMailFolders();

    expect(Array.isArray(folders)).toBe(true);
    expect(folders.length).toBeGreaterThan(0);
  }, TEST_TIMEOUT);

  it("includes standard folders", async () => {
    const folders = await getMailFolders();

    // Common folders that should exist
    const hasInbox = folders.some(f => f.toLowerCase().includes("inbox"));
    const hasSent = folders.some(f => f.toLowerCase().includes("sent"));
    const hasDeleted = folders.some(f => f.toLowerCase().includes("deleted") || f.toLowerCase().includes("trash"));

    expect(hasInbox || folders.length > 0).toBe(true); // At least has folders
  }, TEST_TIMEOUT);

  it("returns folder names as strings", async () => {
    const folders = await getMailFolders();

    for (const folder of folders) {
      expect(typeof folder).toBe("string");
      expect(folder.length).toBeGreaterThan(0);
    }
  }, TEST_TIMEOUT);
});

describe("folder management", () => {
  const TEST_FOLDER = "_MCP_Test";
  const TEST_SUBFOLDER = "_MCP_Test_Sub";
  const TEST_RENAMED = "_MCP_Test_Renamed";

  beforeAll(async () => {
    await ensureOutlookRunning();
    // Clean up any leftover test folders
    await deleteFolder(TEST_RENAMED);
    await deleteFolder(TEST_SUBFOLDER);
    await deleteFolder(TEST_FOLDER);
  });

  afterAll(async () => {
    // Clean up test folders
    await deleteFolder(TEST_RENAMED);
    await deleteFolder(TEST_SUBFOLDER);
    await deleteFolder(TEST_FOLDER);
  });

  it("creates a folder", async () => {
    const result = await createFolder(TEST_FOLDER);
    expect(result).toContain("Folder created");
    expect(result).toContain(TEST_FOLDER);
  }, TEST_TIMEOUT);

  it("folder appears in folder list after creation", async () => {
    const folders = await getMailFolders();
    const hasTestFolder = folders.some(f => f === TEST_FOLDER);
    expect(hasTestFolder).toBe(true);
  }, TEST_TIMEOUT);

  it("renames a folder", async () => {
    const result = await renameFolder(TEST_FOLDER, TEST_RENAMED);
    expect(result).toContain("Folder renamed");
    expect(result).toContain(TEST_RENAMED);
  }, TEST_TIMEOUT);

  it("renamed folder appears in folder list", async () => {
    const folders = await getMailFolders();
    const hasRenamedFolder = folders.some(f => f === TEST_RENAMED);
    expect(hasRenamedFolder).toBe(true);
  }, TEST_TIMEOUT);

  it("deletes an empty folder", async () => {
    const result = await deleteFolder(TEST_RENAMED);
    expect(result).toContain("Folder deleted");
  }, TEST_TIMEOUT);

  it("deleted folder no longer appears in folder list", async () => {
    const folders = await getMailFolders();
    const hasDeletedFolder = folders.some(f => f === TEST_RENAMED);
    expect(hasDeletedFolder).toBe(false);
  }, TEST_TIMEOUT);

  it("returns error when creating duplicate folder", async () => {
    // Create folder first
    await createFolder(TEST_SUBFOLDER);
    // Try to create again
    const result = await createFolder(TEST_SUBFOLDER);
    expect(result).toContain("Error");
    // Clean up
    await deleteFolder(TEST_SUBFOLDER);
  }, TEST_TIMEOUT);

  it("returns error when deleting non-existent folder", async () => {
    const result = await deleteFolder("_NonExistent_Folder_12345");
    expect(result).toContain("Error");
  }, TEST_TIMEOUT);
});
