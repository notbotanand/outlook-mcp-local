import { existsSync } from "node:fs";
import { resolve } from "node:path";
import { type AuthManager } from "../auth/authManager.js";
import { type TokenStore } from "../auth/tokenCache.js";
import { type AuthConfig } from "../config/authConfig.js";
import { getStorageStatus } from "../config/storageStatus.js";
import { getEnabledTools } from "../mcp/tools.js";
import { type OutlookMcpMode } from "../policy/mode.js";

export type DoctorCheck = {
  name: string;
  ok: boolean;
  detail: string;
};

export type DoctorReport = {
  ok: boolean;
  nodeVersion: string;
  mode: OutlookMcpMode;
  tenant: string;
  authConfigured: boolean;
  authenticated: boolean;
  storage: ReturnType<typeof getStorageStatus>;
  mcp: {
    protocolVersion: "2025-11-25";
    transport: "stdio";
    serverCommand: string;
    claudeCodeAddCommand: string;
    toolNames: string[];
  };
  checks: DoctorCheck[];
};

export async function createDoctorReport(input: {
  mode: OutlookMcpMode;
  authConfig: AuthConfig;
  authManager: AuthManager;
  tokenStore: TokenStore;
  projectDir?: string;
  nodeVersion?: string;
}): Promise<DoctorReport> {
  const projectDir = input.projectDir ?? process.cwd();
  const nodeVersion = input.nodeVersion ?? process.version;
  const storage = getStorageStatus(input.tokenStore);
  const authStatus = await input.authManager.status();
  const serverPath = resolve(projectDir, "dist/src/index.js");
  const toolNames = getEnabledTools(input.mode).map((tool) => tool.name);

  const checks: DoctorCheck[] = [
    {
      name: "node_version",
      ok: nodeMajorVersion(nodeVersion) >= 20,
      detail: `Node ${nodeVersion}; required >=20.`
    },
    {
      name: "build_output",
      ok: existsSync(serverPath),
      detail: `MCP server entry: ${serverPath}`
    },
    {
      name: "env_file",
      ok: storage.configFile.exists,
      detail: storage.configFile.exists ? `.env found at ${storage.configFile.path}` : ".env not found; run npm run setup."
    },
    {
      name: "env_permissions",
      ok: storage.configFile.ownerOnlyReadWrite === true,
      detail: storage.configFile.mode
        ? `.env mode is ${storage.configFile.mode}; expected 600.`
        : ".env mode unavailable."
    },
    {
      name: "client_id_configured",
      ok: input.authConfig.clientId !== undefined,
      detail: input.authConfig.clientId !== undefined
        ? "Microsoft client ID is configured; value is hidden."
        : "Microsoft client ID is missing; run npm run setup."
    },
    {
      name: "authenticated",
      ok: authStatus.authenticated,
      detail: authStatus.authenticated
        ? "Local token store has authentication state."
        : "Not authenticated; run npm run setup or auth_login."
    },
    {
      name: "calendar_write_mode",
      ok: input.mode === "read" || toolNames.includes("outlook_create_calendar_block"),
      detail: input.mode === "calendar-write"
        ? "Calendar block creation tool is enabled."
        : "Read mode active; calendar block creation tool is disabled."
    }
  ];

  return {
    ok: checks.every((check) => check.ok),
    nodeVersion,
    mode: input.mode,
    tenant: input.authConfig.tenant,
    authConfigured: input.authConfig.clientId !== undefined,
    authenticated: authStatus.authenticated,
    storage,
    mcp: {
      protocolVersion: "2025-11-25",
      transport: "stdio",
      serverCommand: `node ${serverPath}`,
      claudeCodeAddCommand: `claude mcp add --transport stdio outlook-mcp-local -- node ${serverPath}`,
      toolNames
    },
    checks
  };
}

function nodeMajorVersion(value: string): number {
  const match = /^v?(\d+)/.exec(value);
  return match ? Number(match[1]) : 0;
}
