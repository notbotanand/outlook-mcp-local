import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  assertClientIdConfigured,
  parseTenant,
  readAuthConfigFromEnv
} from "../../src/config/authConfig.js";

describe("auth config", () => {
  it("defaults tenant and redirect URI safely", () => {
    const config = readAuthConfigFromEnv({}, "read");

    assert.equal(config.clientId, undefined);
    assert.equal(config.tenant, "consumers");
    assert.equal(config.redirectUri, "http://127.0.0.1:53682/callback");
  });

  it("accepts configured client ID and tenant", () => {
    const config = readAuthConfigFromEnv(
      {
        OUTLOOK_MCP_CLIENT_ID: "synthetic-client-id",
        OUTLOOK_MCP_TENANT: "common",
        OUTLOOK_MCP_REDIRECT_URI: "http://127.0.0.1:60000/callback"
      },
      "calendar-write"
    );

    assert.equal(assertClientIdConfigured(config), "synthetic-client-id");
    assert.equal(config.tenant, "common");
    assert.equal(config.redirectUri, "http://127.0.0.1:60000/callback");
  });

  it("rejects unsupported tenants", () => {
    assert.throws(() => parseTenant("organizations"), /Unsupported OUTLOOK_MCP_TENANT/);
  });
});
