import assert from "node:assert/strict";
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { describe, it } from "node:test";
import { InMemoryTokenStore, MacOsKeychainRefreshTokenStore } from "../../src/auth/tokenCache.js";
import { getStorageStatus } from "../../src/config/storageStatus.js";

describe("storage status", () => {
  it("reports config permissions and never includes client ID values", () => {
    const dir = join(tmpdir(), `outlook-mcp-local-storage-${Date.now()}`);
    mkdirSync(dir);
    const envPath = join(dir, ".env");
    writeFileSync(envPath, 'OUTLOOK_MCP_CLIENT_ID="synthetic-client-id"\n', { mode: 0o600 });

    try {
      const previous = process.env.OUTLOOK_MCP_CLIENT_ID;
      process.env.OUTLOOK_MCP_CLIENT_ID = "synthetic-client-id";
      const status = getStorageStatus(new InMemoryTokenStore(), envPath);
      process.env.OUTLOOK_MCP_CLIENT_ID = previous;

      assert.equal(status.configFile.exists, true);
      assert.equal(status.configFile.mode, "600");
      assert.equal(status.configFile.ownerOnlyReadWrite, true);
      assert.equal(status.configFile.containsClientId, true);
      assert.equal(status.configFile.containsClientIdValue, false);
      assert.equal(JSON.stringify(status).includes("synthetic-client-id"), false);
      assert.equal(status.serverPersistsMailBodies, false);
      assert.equal(status.serverPersistsRawGraphResponses, false);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("describes macOS keychain token storage without token values", () => {
    assert.deepEqual(new MacOsKeychainRefreshTokenStore("service", "account").describe(), {
      refreshTokenStore: "macOS Keychain",
      keychainService: "service",
      keychainAccount: "account",
      accessTokenStore: "process memory"
    });
  });
});
