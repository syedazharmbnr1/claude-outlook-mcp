import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_TIMEOUT } from "../setup";
import { getTodayEvents, getUpcomingEvents, searchEvents, createEvent } from "../../index";

describe("getTodayEvents", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns array of events", async () => {
    const events = await getTodayEvents(10);

    expect(Array.isArray(events)).toBe(true);
    // May be empty if no events today
  }, TEST_TIMEOUT);

  it("respects limit parameter", async () => {
    const events = await getTodayEvents(3);
    expect(events.length).toBeLessThanOrEqual(3);
  }, TEST_TIMEOUT);
});

describe("getUpcomingEvents", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns array of upcoming events", async () => {
    const events = await getUpcomingEvents(7, 10);

    expect(Array.isArray(events)).toBe(true);
    // May be empty if no upcoming events
  }, TEST_TIMEOUT);

  it("respects limit parameter", async () => {
    const events = await getUpcomingEvents(7, 3);
    expect(events.length).toBeLessThanOrEqual(3);
  }, TEST_TIMEOUT);
});

describe("searchEvents", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns array for search", async () => {
    const events = await searchEvents("meeting", 5);

    expect(Array.isArray(events)).toBe(true);
    expect(events.length).toBeLessThanOrEqual(5);
  }, TEST_TIMEOUT);

  it("returns empty array when no matches", async () => {
    const events = await searchEvents("xyznonexistenteventxyz123456", 5);
    expect(events).toEqual([]);
  }, TEST_TIMEOUT);
});

describe("createEvent", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("creates event and returns confirmation", async () => {
    // Create event 1 day in the future
    const tomorrow = new Date(Date.now() + 86400000);
    const endTime = new Date(tomorrow.getTime() + 3600000); // 1 hour later

    const result = await createEvent(
      `[TEST] MCP Test Event - ${Date.now()}`,
      tomorrow.toISOString(),
      endTime.toISOString(),
      "Test Location",
      "This is an automated test event from Outlook MCP plugin tests. Safe to delete."
    );

    expect(result).toContain("success");
  }, TEST_TIMEOUT);
});
