#!/usr/bin/env node
import { stdout } from "node:process";
import { AuthManager } from "../auth/authManager.js";
import { createDefaultTokenStore } from "../auth/tokenCache.js";
import { readAuthConfigFromEnv } from "../config/authConfig.js";
import { loadEnvFile } from "../config/envFile.js";
import { parseOutlookMcpModeFromEnv } from "../policy/mode.js";
import { createDoctorReport } from "./doctorStatus.js";

async function main(): Promise<void> {
  loadEnvFile();
  const mode = parseOutlookMcpModeFromEnv(process.env);
  const authConfig = readAuthConfigFromEnv(process.env, mode);
  const tokenStore = createDefaultTokenStore();
  const authManager = new AuthManager(authConfig, tokenStore);
  const report = await createDoctorReport({ mode, authConfig, tokenStore, authManager });

  if (process.argv.includes("--json")) {
    stdout.write(`${JSON.stringify(report, null, 2)}\n`);
    process.exitCode = report.ok ? 0 : 1;
    return;
  }

  stdout.write("outlook-mcp-local doctor\n\n");
  for (const check of report.checks) {
    stdout.write(`${check.ok ? "ok" : "fail"}  ${check.name}: ${check.detail}\n`);
  }

  stdout.write("\nMCP server command:\n");
  stdout.write(`${report.mcp.serverCommand}\n\n`);
  stdout.write("Claude Code add command:\n");
  stdout.write(`${report.mcp.claudeCodeAddCommand}\n\n`);
  stdout.write("Enabled tools:\n");
  for (const toolName of report.mcp.toolNames) {
    stdout.write(`- ${toolName}\n`);
  }

  stdout.write("\nNo client IDs, tokens, mail bodies, or calendar bodies are printed by doctor.\n");
  process.exitCode = report.ok ? 0 : 1;
}

main().catch((error: unknown) => {
  stdout.write(`Doctor failed: ${error instanceof Error ? error.message : "Unknown error"}\n`);
  process.exitCode = 1;
});
