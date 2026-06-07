#!/usr/bin/env node
import { createInterface } from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { AuthManager } from "../auth/authManager.js";
import { createDefaultTokenStore } from "../auth/tokenCache.js";
import { readAuthConfigFromEnv } from "../config/authConfig.js";
import { loadEnvFile, writeEnvFile } from "../config/envFile.js";
import { getStorageStatus } from "../config/storageStatus.js";
import { parseOutlookMcpMode } from "../policy/mode.js";

const appRegistrationUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade";
const defaultRedirectUri = "http://127.0.0.1:53682/callback";

async function main(): Promise<void> {
  loadEnvFile();

  const rl = createInterface({ input, output });
  try {
    output.write("outlook-mcp-local setup\n\n");
    output.write("This connector runs locally, uses your Microsoft app/client ID, and does not ship a shared OAuth app.\n");
    output.write("Register a Microsoft app with a localhost redirect URI, then paste the client ID here.\n");
    output.write("The recommended auth flow is OAuth 2.0 authorization code with PKCE.\n\n");
    output.write(`Microsoft app registrations: ${appRegistrationUrl}\n`);
    const redirectUri = process.env.OUTLOOK_MCP_REDIRECT_URI ?? defaultRedirectUri;
    output.write(`Redirect URI to register: ${redirectUri}\n\n`);

    const currentClientId = process.env.OUTLOOK_MCP_CLIENT_ID;
    const clientId = await askSensitive(rl, "Microsoft client ID", currentClientId);
    validateClientId(clientId);

    const tenant = await ask(rl, "Tenant [consumers/common]", process.env.OUTLOOK_MCP_TENANT ?? "consumers");
    if (tenant !== "consumers" && tenant !== "common") {
      throw new Error('Tenant must be "consumers" or "common".');
    }

    const mode = await ask(rl, "Mode [read/calendar-write]", process.env.OUTLOOK_MCP_MODE ?? "read");
    parseOutlookMcpMode(mode);

    writeEnvFile(".env", {
      OUTLOOK_MCP_CLIENT_ID: clientId,
      OUTLOOK_MCP_TENANT: tenant,
      OUTLOOK_MCP_REDIRECT_URI: redirectUri,
      OUTLOOK_MCP_MODE: mode
    });

    process.env.OUTLOOK_MCP_CLIENT_ID = clientId;
    process.env.OUTLOOK_MCP_TENANT = tenant;
    process.env.OUTLOOK_MCP_REDIRECT_URI = redirectUri;
    process.env.OUTLOOK_MCP_MODE = mode;

    output.write("\nWrote .env with mode 0600.\n");
    const runAuth = await ask(rl, "Start browser OAuth with PKCE now? [Y/n]", "Y");
    if (runAuth.toLowerCase() !== "n" && runAuth.toLowerCase() !== "no") {
      try {
        await runPkceAuth(rl);
      } catch (error) {
        throw new Error(formatAuthSetupError(error));
      }
    }

    output.write("\nSetup complete.\n");
    printStorageSummary(createDefaultTokenStore());
    output.write("Start the MCP server with: node dist/src/index.js\n");
  } finally {
    rl.close();
  }
}

function printStorageSummary(tokenStore: ReturnType<typeof createDefaultTokenStore>): void {
  const status = getStorageStatus(tokenStore);
    output.write("\nLocal storage summary:\n");
    output.write(`- Config file: ${status.configFile.path} (${status.configFile.mode ?? "missing"})\n`);
    output.write(`- Refresh token store: ${status.tokenStore.refreshTokenStore}\n`);
    output.write("- Access token store: process memory\n");
    output.write("- Mail/calendar bodies persisted by this server: no\n");
    output.write("- MCP clients may keep their own transcripts outside this server.\n\n");
}

async function runPkceAuth(rl: ReturnType<typeof createInterface>): Promise<void> {
  const mode = parseOutlookMcpMode(process.env.OUTLOOK_MCP_MODE);
  const authManager = new AuthManager(
    readAuthConfigFromEnv(process.env, mode),
    createDefaultTokenStore()
  );
  const login = await authManager.login({ flow: "pkce" });
  if (login.flow !== "pkce") {
    throw new Error("Expected PKCE login flow.");
  }

  output.write("\nOpen this URL in your browser:\n");
  output.write(`${login.authorizationUrl}\n\n`);
  output.write("After Microsoft redirects to the local callback and the browser says authentication is complete, press Enter here.\n");

  for (let attempt = 0; attempt < 120; attempt += 1) {
    await rl.question("");
    const status = await authManager.status();
    if (status.authenticated) {
      output.write("Authenticated: yes\n");
      return;
    }

    output.write("Not authenticated yet. Complete the browser flow, then press Enter again.\n");
  }

  throw new Error("Timed out waiting for browser authentication.");
}

async function ask(
  rl: ReturnType<typeof createInterface>,
  question: string,
  defaultValue?: string
): Promise<string> {
  const suffix = defaultValue ? ` (${defaultValue})` : "";
  const answer = (await rl.question(`${question}${suffix}: `)).trim();
  return answer.length > 0 ? answer : defaultValue ?? "";
}

async function askSensitive(
  rl: ReturnType<typeof createInterface>,
  question: string,
  currentValue?: string
): Promise<string> {
  const suffix = currentValue ? " (configured; press Enter to keep)" : "";
  const answer = (await rl.question(`${question}${suffix}: `)).trim();
  return answer.length > 0 ? answer : currentValue ?? "";
}

function validateClientId(clientId: string): void {
  if (!/^[0-9a-fA-F-]{8,}$/.test(clientId)) {
    throw new Error("Client ID does not look valid. Paste the Application (client) ID from Microsoft.");
  }
}

function formatAuthSetupError(error: unknown): string {
  const message = error instanceof Error ? error.message : "Unknown authentication error.";
  if (message.includes("AADSTS70002")) {
    return [
      message,
      "",
      "Microsoft rejected this OAuth flow for the app registration.",
      "Confirm the app has a localhost redirect URI matching:",
      `${process.env.OUTLOOK_MCP_REDIRECT_URI ?? defaultRedirectUri}`,
      "Then rerun npm run setup."
    ].join("\n");
  }

  return message;
}

main().catch((error: unknown) => {
  output.write(`Setup failed: ${error instanceof Error ? error.message : "Unknown error"}\n`);
  process.exitCode = 1;
});
