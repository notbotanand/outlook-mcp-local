import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { assertAllowedScopes, scopesForMode } from "../../src/policy/scopePolicy.js";

describe("scope policy", () => {
  it("uses read-only mail and calendar scopes by default", () => {
    assert.deepEqual(scopesForMode("read"), [
      "openid",
      "profile",
      "email",
      "offline_access",
      "Mail.Read",
      "Calendars.Read"
    ]);
  });

  it("uses calendar write scope only in explicit calendar-write mode", () => {
    assert.deepEqual(scopesForMode("calendar-write"), [
      "openid",
      "profile",
      "email",
      "offline_access",
      "Mail.Read",
      "Calendars.ReadWrite"
    ]);
  });

  it("rejects email write scopes", () => {
    assert.throws(() => assertAllowedScopes(["Mail.Send"]), /Disallowed email write scopes/);
    assert.throws(() => assertAllowedScopes(["Mail.ReadWrite"]), /Disallowed email write scopes/);
    assert.throws(
      () => assertAllowedScopes(["MailboxSettings.ReadWrite"]),
      /Disallowed email write scopes/
    );
  });
});
