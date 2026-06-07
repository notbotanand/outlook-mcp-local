import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { AuthManager } from "../../src/auth/authManager.js";
import { InMemoryTokenStore } from "../../src/auth/tokenCache.js";

describe("AuthManager", () => {
  it("reports unconfigured status without tokens", async () => {
    const manager = new AuthManager(
      {
        tenant: "consumers",
        redirectUri: "http://127.0.0.1:53682/callback",
        mode: "read"
      },
      new InMemoryTokenStore()
    );

    const status = await manager.status();

    assert.equal(status.configured, false);
    assert.equal(status.authenticated, false);
    assert.equal(status.tokenStatus.hasAccessToken, false);
    assert.equal(status.tokenStatus.hasRefreshToken, false);
  });

  it("rejects login without BYO client ID", async () => {
    const manager = new AuthManager(
      {
        tenant: "consumers",
        redirectUri: "http://127.0.0.1:53682/callback",
        mode: "read"
      },
      new InMemoryTokenStore()
    );

    await assert.rejects(() => manager.login(), /OUTLOOK_MCP_CLIENT_ID is required/);
  });

  it("starts and completes device login without returning token values", async () => {
    let requestCount = 0;
    const fetchMock = async () => {
      requestCount += 1;

      if (requestCount === 1) {
        return new Response(
          JSON.stringify({
            device_code: "synthetic-device-code",
            user_code: "ABCD-EFGH",
            verification_uri: "https://microsoft.com/devicelogin",
            expires_in: 900,
            interval: 5
          }),
          { status: 200, headers: { "content-type": "application/json" } }
        );
      }

      return new Response(
        JSON.stringify({
          access_token: "synthetic-access-token",
          refresh_token: "synthetic-refresh-token",
          expires_in: 60,
          scope: "Mail.Read Calendars.Read",
          token_type: "Bearer"
        }),
        { status: 200, headers: { "content-type": "application/json" } }
      );
    };

    const manager = new AuthManager(
      {
        clientId: "synthetic-client-id",
        tenant: "consumers",
        redirectUri: "http://127.0.0.1:53682/callback",
        mode: "read"
      },
      new InMemoryTokenStore(),
      fetchMock
    );

    const login = await manager.login({ flow: "device" });
    assert.equal(login.flow, "device");
    assert.equal("deviceCode" in login, false);

    const complete = await manager.complete(login.authSessionId);
    assert.equal(complete.authenticated, true);
    assert.equal(complete.status.hasAccessToken, true);
    assert.equal(complete.status.hasRefreshToken, true);
    assert.equal(JSON.stringify(complete).includes("synthetic-access-token"), false);
    assert.equal(JSON.stringify(complete).includes("synthetic-refresh-token"), false);
  });
});
