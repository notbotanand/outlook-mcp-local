import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  isCalendarWriteEnabled,
  parseOutlookMcpMode,
  parseOutlookMcpModeFromEnv
} from "../../src/policy/mode.js";

describe("mode policy", () => {
  it("defaults to read mode", () => {
    assert.equal(parseOutlookMcpMode(undefined), "read");
    assert.equal(parseOutlookMcpMode(""), "read");
    assert.equal(parseOutlookMcpModeFromEnv({}), "read");
  });

  it("accepts only explicit supported modes", () => {
    assert.equal(parseOutlookMcpMode("read"), "read");
    assert.equal(parseOutlookMcpMode("calendar-write"), "calendar-write");
    assert.equal(isCalendarWriteEnabled("read"), false);
    assert.equal(isCalendarWriteEnabled("calendar-write"), true);
  });

  it("fails closed for unsupported modes", () => {
    assert.throws(() => parseOutlookMcpMode("write"), /Unsupported OUTLOOK_MCP_MODE/);
  });
});
