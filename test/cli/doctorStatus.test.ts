import assert from "node:assert/strict";
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { describe, it } from "node:test";
import { AuthManager } from "../../src/auth/authManager.js";
import { InMemoryTokenStore } from "../../src/auth/tokenCache.js";
import { createDoctorReport } from "../../src/cli/doctorStatus.js";

describe("doctor status", () => {
  it("reports readiness without exposing client IDs or tokens", async () => {
    const dir = join(tmpdir(), `outlook-mcp-doctor-${Date.now()}`);
    mkdirSync(join(dir, "dist/src"), { recursive: true });
    writeFileSync(join(dir, "dist/src/index.js"), "");
    writeFileSync(join(dir, ".env"), 'OUTLOOK_MCP_CLIENT_ID="synthetic-client-id"\n', { mode: 0o600 });

    const previousCwd = process.cwd();
    const previousClientId = process.env.OUTLOOK_MCP_CLIENT_ID;
    process.chdir(dir);
    process.env.OUTLOOK_MCP_CLIENT_ID = "synthetic-client-id";

    try {
      const tokenStore = new InMemoryTokenStore();
      await tokenStore.write({
        accessToken: "synthetic-access-token",
        refreshToken: "synthetic-refresh-token",
        expiresAt: Date.now() + 60_000
      });
      const authConfig = {
        clientId: "synthetic-client-id",
        tenant: "consumers" as const,
        redirectUri: "http://127.0.0.1:53682/callback",
        mode: "read" as const
      };
      const report = await createDoctorReport({
        mode: "read",
        authConfig,
        tokenStore,
        authManager: new AuthManager(authConfig, tokenStore),
        projectDir: dir,
        nodeVersion: "v20.0.0"
      });
      const serialized = JSON.stringify(report);

      assert.equal(report.ok, true);
      assert.equal(report.authConfigured, true);
      assert.equal(report.authenticated, true);
      assert.match(report.mcp.claudeCodeAddCommand, /claude mcp add/);
      assert.equal(serialized.includes("synthetic-client-id"), false);
      assert.equal(serialized.includes("synthetic-access-token"), false);
      assert.equal(serialized.includes("synthetic-refresh-token"), false);
    } finally {
      process.chdir(previousCwd);
      if (previousClientId === undefined) {
        delete process.env.OUTLOOK_MCP_CLIENT_ID;
      } else {
        process.env.OUTLOOK_MCP_CLIENT_ID = previousClientId;
      }
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
