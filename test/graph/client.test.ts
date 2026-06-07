import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { buildGraphUrl, mapGraphError } from "../../src/graph/client.js";

describe("Graph client helpers", () => {
  it("builds Microsoft Graph v1 URLs", () => {
    const url = new URL(
      buildGraphUrl("/me/messages", {
        "$select": "id,subject",
        "$top": 10
      })
    );

    assert.equal(url.origin, "https://graph.microsoft.com");
    assert.equal(url.pathname, "/v1.0/me/messages");
    assert.equal(url.searchParams.get("$select"), "id,subject");
    assert.equal(url.searchParams.get("$top"), "10");
  });

  it("maps Graph errors without raw response content", () => {
    assert.equal(mapGraphError(401), "Microsoft Graph authentication failed. Run auth_login again.");
    assert.equal(
      mapGraphError(403),
      "Microsoft Graph denied access. Re-run setup/auth with the required read or calendar-write mode."
    );
    assert.equal(mapGraphError(429), "Microsoft Graph throttled the request. Try again later.");
  });
});
