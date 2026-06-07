import { type OutlookMcpMode } from "../policy/mode.js";

export type OutlookTenant = "consumers" | "common";

export type AuthConfig = {
  clientId?: string;
  tenant: OutlookTenant;
  redirectUri: string;
  mode: OutlookMcpMode;
};

const defaultTenant: OutlookTenant = "consumers";
const defaultRedirectUri = "http://127.0.0.1:53682/callback";

export function readAuthConfigFromEnv(
  env: NodeJS.ProcessEnv,
  mode: OutlookMcpMode
): AuthConfig {
  return {
    clientId: normalizeOptional(env.OUTLOOK_MCP_CLIENT_ID),
    tenant: parseTenant(env.OUTLOOK_MCP_TENANT),
    redirectUri: normalizeOptional(env.OUTLOOK_MCP_REDIRECT_URI) ?? defaultRedirectUri,
    mode
  };
}

export function assertClientIdConfigured(config: AuthConfig): string {
  if (config.clientId === undefined) {
    throw new Error("OUTLOOK_MCP_CLIENT_ID is required for Microsoft authentication.");
  }

  return config.clientId;
}

export function parseTenant(value: string | undefined): OutlookTenant {
  if (value === undefined || value.trim() === "") {
    return defaultTenant;
  }

  if (value === "consumers" || value === "common") {
    return value;
  }

  throw new Error(
    `Unsupported OUTLOOK_MCP_TENANT "${value}". Supported tenants are "consumers" and "common".`
  );
}

function normalizeOptional(value: string | undefined): string | undefined {
  if (value === undefined || value.trim() === "") {
    return undefined;
  }

  return value.trim();
}
