import { type OutlookMcpMode } from "./mode.js";

export const disallowedEmailScopes = [
  "Mail.Send",
  "Mail.ReadWrite",
  "MailboxSettings.ReadWrite"
] as const;

const baseScopes = ["openid", "profile", "email", "offline_access", "Mail.Read"] as const;

export function scopesForMode(mode: OutlookMcpMode): string[] {
  const calendarScope = mode === "calendar-write" ? "Calendars.ReadWrite" : "Calendars.Read";
  const scopes = [...baseScopes, calendarScope];
  assertAllowedScopes(scopes);
  return scopes;
}

export function assertAllowedScopes(scopes: readonly string[]): void {
  const blocked = scopes.filter((scope) =>
    disallowedEmailScopes.some((disallowed) => disallowed.toLowerCase() === scope.toLowerCase())
  );

  if (blocked.length > 0) {
    throw new Error(`Disallowed email write scopes requested: ${blocked.join(", ")}`);
  }
}
