import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { loadEnvFile, parseEnvFile } from "../../src/config/envFile.js";

describe("env file parsing", () => {
  it("parses quoted local config values", () => {
    assert.deepEqual(
      parseEnvFile(`
# comment
OUTLOOK_MCP_CLIENT_ID="synthetic-client-id"
OUTLOOK_MCP_TENANT=consumers
OUTLOOK_MCP_MODE='read'
`),
      {
        OUTLOOK_MCP_CLIENT_ID: "synthetic-client-id",
        OUTLOOK_MCP_TENANT: "consumers",
        OUTLOOK_MCP_MODE: "read"
      }
    );
  });

  it("records the loaded env file path", () => {
    const env: NodeJS.ProcessEnv = {};
    loadEnvFile("/tmp/nonexistent-outlook-mcp-local-env", env);

    assert.equal(env.OUTLOOK_MCP_ENV_FILE, "/tmp/nonexistent-outlook-mcp-local-env");
  });
});
