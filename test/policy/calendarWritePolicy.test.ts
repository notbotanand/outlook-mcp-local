import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { evaluateCalendarBlockPolicy } from "../../src/policy/calendarWritePolicy.js";

describe("calendar write policy", () => {
  it("blocks calendar creation in read mode", () => {
    const result = evaluateCalendarBlockPolicy("read", {});

    assert.deepEqual(result, {
      allowed: false,
      reason: "Calendar write mode is required to create a personal calendar block."
    });
  });

  it("allows personal blocks in calendar-write mode", () => {
    assert.deepEqual(evaluateCalendarBlockPolicy("calendar-write", {}), { allowed: true });
  });

  it("rejects attendees, online meetings, recurrence, attachments, and other calendars", () => {
    assert.equal(evaluateCalendarBlockPolicy("calendar-write", { attendees: ["person@example.com"] }).allowed, false);
    assert.equal(evaluateCalendarBlockPolicy("calendar-write", { isOnlineMeeting: true }).allowed, false);
    assert.equal(evaluateCalendarBlockPolicy("calendar-write", { recurrence: {} }).allowed, false);
    assert.equal(evaluateCalendarBlockPolicy("calendar-write", { attachments: [{}] }).allowed, false);
    assert.equal(evaluateCalendarBlockPolicy("calendar-write", { calendarId: "shared" }).allowed, false);
  });
});
