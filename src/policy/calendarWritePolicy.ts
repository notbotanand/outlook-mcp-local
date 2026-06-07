import { isCalendarWriteEnabled, type OutlookMcpMode } from "./mode.js";

export type CalendarBlockPolicyInput = {
  attendees?: unknown[];
  isOnlineMeeting?: boolean;
  recurrence?: unknown;
  attachments?: unknown[];
  calendarId?: string;
};

export type CalendarBlockPolicyResult =
  | { allowed: true }
  | { allowed: false; reason: string };

export function evaluateCalendarBlockPolicy(
  mode: OutlookMcpMode,
  input: CalendarBlockPolicyInput
): CalendarBlockPolicyResult {
  if (!isCalendarWriteEnabled(mode)) {
    return {
      allowed: false,
      reason: "Calendar write mode is required to create a personal calendar block."
    };
  }

  if (Array.isArray(input.attendees) && input.attendees.length > 0) {
    return { allowed: false, reason: "Attendees are not allowed for calendar blocks." };
  }

  if (input.isOnlineMeeting === true) {
    return { allowed: false, reason: "Online meetings are not allowed for calendar blocks." };
  }

  if (input.recurrence !== undefined && input.recurrence !== null) {
    return { allowed: false, reason: "Recurring calendar blocks are not supported in v1." };
  }

  if (Array.isArray(input.attachments) && input.attachments.length > 0) {
    return { allowed: false, reason: "Attachments are not allowed for calendar blocks." };
  }

  if (input.calendarId !== undefined && input.calendarId !== "me") {
    return {
      allowed: false,
      reason: "Calendar blocks may only target the signed-in user's own calendar."
    };
  }

  return { allowed: true };
}

export function assertCalendarBlockAllowed(
  mode: OutlookMcpMode,
  input: CalendarBlockPolicyInput
): void {
  const result = evaluateCalendarBlockPolicy(mode, input);
  if (!result.allowed) {
    throw new Error(result.reason);
  }
}
