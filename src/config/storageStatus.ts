import { existsSync, statSync } from "node:fs";
import { resolve } from "node:path";
import { type TokenStore } from "../auth/tokenCache.js";

export type StorageStatus = {
  configFile: {
    path: string;
    exists: boolean;
    mode?: string;
    ownerOnlyReadWrite?: boolean;
    containsClientId: boolean;
    containsClientIdValue: false;
  };
  tokenStore: ReturnType<TokenStore["describe"]>;
  accessTokenStore: "process memory";
  sessionReferenceStore: "process memory";
  serverPersistsMailBodies: false;
  serverPersistsCalendarBodies: false;
  serverPersistsRawGraphResponses: false;
  mcpClientCaveat: string;
};

export function getStorageStatus(
  tokenStore: TokenStore,
  envPath = process.env.OUTLOOK_MCP_ENV_FILE ?? ".env"
): StorageStatus {
  const absolutePath = resolve(envPath);
  const exists = existsSync(envPath);
  const mode = exists ? statSync(envPath).mode & 0o777 : undefined;

  return {
    configFile: {
      path: absolutePath,
      exists,
      mode: mode === undefined ? undefined : mode.toString(8).padStart(3, "0"),
      ownerOnlyReadWrite: mode === undefined ? undefined : mode === 0o600,
      containsClientId: process.env.OUTLOOK_MCP_CLIENT_ID !== undefined,
      containsClientIdValue: false
    },
    tokenStore: tokenStore.describe(),
    accessTokenStore: "process memory",
    sessionReferenceStore: "process memory",
    serverPersistsMailBodies: false,
    serverPersistsCalendarBodies: false,
    serverPersistsRawGraphResponses: false,
    mcpClientCaveat:
      "When a tool returns mail or calendar content, the MCP client may store it in its own transcript, logs, or history outside this server."
  };
}
