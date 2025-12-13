// Shared test configuration and helpers
import { runAppleScript } from "run-applescript";

// Set TEST_RECIPIENT environment variable to your email for integration tests
// Example: TEST_RECIPIENT=you@example.com bun test
export const TEST_RECIPIENT = process.env.TEST_RECIPIENT || "test@example.com";
export const TEST_TIMEOUT = 30000; // AppleScript can be slow

/**
 * Check if Outlook is running before integration tests
 */
export async function ensureOutlookRunning(): Promise<void> {
  const script = `
    tell application "System Events"
      return (name of processes) contains "Microsoft Outlook"
    end tell
  `;

  const result = await runAppleScript(script);
  if (result !== "true") {
    throw new Error("Microsoft Outlook must be running for integration tests");
  }
}

/**
 * Helper to wait for Outlook to process commands
 */
export async function waitForOutlook(ms: number = 1000): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}
