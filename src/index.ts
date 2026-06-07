#!/usr/bin/env node
import { createMcpServer } from "./mcp/server.js";
import { AuthManager } from "./auth/authManager.js";
import { createDefaultTokenStore } from "./auth/tokenCache.js";
import { GraphCalendarService } from "./calendar/graphCalendar.js";
import { readAuthConfigFromEnv } from "./config/authConfig.js";
import { loadEnvFile } from "./config/envFile.js";
import { createGraphClient } from "./graph/client.js";
import { GraphMailService } from "./mail/graphMail.js";
import { parseOutlookMcpModeFromEnv } from "./policy/mode.js";
import { createLogger } from "./util/logger.js";

loadEnvFile();

const logger = createLogger();
const mode = parseOutlookMcpModeFromEnv(process.env);
const authConfig = readAuthConfigFromEnv(process.env, mode);
const tokenStore = createDefaultTokenStore();
const authManager = new AuthManager(authConfig, tokenStore);
const graphClient = createGraphClient(authManager);
const mailService = new GraphMailService(graphClient);
const calendarService = new GraphCalendarService(graphClient);
const server = createMcpServer({
  mode,
  logger,
  authManager,
  tokenStore,
  mailService,
  calendarService
});

logger.info("Starting outlook-mcp-local", {
  mode,
  authConfigured: authConfig.clientId !== undefined,
  tenant: authConfig.tenant,
  tools: server.getToolNames()
});

server.start(process.stdin, process.stdout);
