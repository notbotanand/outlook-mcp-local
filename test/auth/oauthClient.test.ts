import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { buildAuthorizationUrl, toTokenSet } from "../../src/auth/oauthClient.js";

describe("OAuth client helpers", () => {
  it("builds a Microsoft authorization URL without client secrets", () => {
    const url = new URL(
      buildAuthorizationUrl(
        {
          clientId: "synthetic-client-id",
          tenant: "consumers",
          redirectUri: "http://127.0.0.1:53682/callback",
          mode: "read"
        },
        ["openid", "Mail.Read", "Calendars.Read"],
        {
          verifier: "verifier",
          challenge: "challenge",
          method: "S256"
        },
        "state"
      )
    );

    assert.equal(url.origin, "https://login.microsoftonline.com");
    assert.equal(url.searchParams.get("client_id"), "synthetic-client-id");
    assert.equal(url.searchParams.get("client_secret"), null);
    assert.equal(url.searchParams.get("code_challenge_method"), "S256");
  });

  it("maps token responses to internal token sets", () => {
    const tokens = toTokenSet(
      {
        access_token: "synthetic-access-token",
        refresh_token: "synthetic-refresh-token",
        expires_in: 60,
        scope: "Mail.Read Calendars.Read",
        token_type: "Bearer"
      },
      1000
    );

    assert.equal(tokens.accessToken, "synthetic-access-token");
    assert.equal(tokens.refreshToken, "synthetic-refresh-token");
    assert.equal(tokens.expiresAt, 61000);
  });
});
