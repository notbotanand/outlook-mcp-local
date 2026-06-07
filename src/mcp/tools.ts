import {
  authCompleteInputSchema,
  authLoginInputSchema,
  authLogoutInputSchema,
  authStatusInputSchema,
  outlookCreateCalendarBlockInputSchema,
  outlookGetCalendarViewInputSchema,
  outlookListMailInputSchema,
  outlookReadMailInputSchema,
  outlookSearchCalendarInputSchema,
  outlookSearchMailInputSchema,
  storageStatusInputSchema,
  type JsonSchema
} from "./schemas.js";
import { type AuthManager } from "../auth/authManager.js";
import { type TokenStore } from "../auth/tokenCache.js";
import { GraphCalendarService } from "../calendar/graphCalendar.js";
import { getStorageStatus } from "../config/storageStatus.js";
import { createGraphClient } from "../graph/client.js";
import { emailContentWarning } from "../mail/constants.js";
import { GraphMailService } from "../mail/graphMail.js";
import { evaluateCalendarBlockPolicy } from "../policy/calendarWritePolicy.js";
import { type OutlookMcpMode } from "../policy/mode.js";

type MailService = Pick<GraphMailService, "searchMail" | "listMail" | "readMail">;
type CalendarService = Pick<
  GraphCalendarService,
  "getCalendarView" | "searchCalendar" | "createCalendarBlock"
>;

export type McpTool = {
  name: string;
  description: string;
  inputSchema: JsonSchema;
};

export type ToolResult = {
  content: Array<{
    type: "text";
    text: string;
  }>;
  isError?: boolean;
};

export type ToolRuntime = {
  mode: OutlookMcpMode;
  authManager?: AuthManager;
  tokenStore?: TokenStore;
  mailService?: MailService;
  calendarService?: CalendarService;
};

export const toolDefinitions = [
  {
    name: "outlook_search_mail",
    description: "Search Outlook mail and return metadata only.",
    inputSchema: outlookSearchMailInputSchema
  },
  {
    name: "outlook_list_mail",
    description: "List Outlook mail metadata for a date range.",
    inputSchema: outlookListMailInputSchema
  },
  {
    name: "outlook_read_mail",
    description: "Read a selected email by local message reference.",
    inputSchema: outlookReadMailInputSchema
  },
  {
    name: "outlook_get_calendar_view",
    description: "List Outlook calendar event metadata for a date range.",
    inputSchema: outlookGetCalendarViewInputSchema
  },
  {
    name: "outlook_search_calendar",
    description: "Search Outlook calendar event metadata by query and date range.",
    inputSchema: outlookSearchCalendarInputSchema
  },
  {
    name: "outlook_create_calendar_block",
    description:
      "Create a personal calendar block only when calendar-write mode is enabled. Stubbed in Sprint 0.",
    inputSchema: outlookCreateCalendarBlockInputSchema
  },
  {
    name: "auth_status",
    description: "Return safe Microsoft authentication status without exposing tokens.",
    inputSchema: authStatusInputSchema
  },
  {
    name: "auth_login",
    description:
      "Start Microsoft OAuth login using BYO client ID. Supports PKCE localhost callback or device code.",
    inputSchema: authLoginInputSchema
  },
  {
    name: "auth_complete",
    description:
      "Complete a pending device-code login poll. PKCE login completes through the local callback.",
    inputSchema: authCompleteInputSchema
  },
  {
    name: "auth_logout",
    description: "Clear locally stored Microsoft authentication tokens without returning token values.",
    inputSchema: authLogoutInputSchema
  },
  {
    name: "storage_status",
    description: "Show where local config, tokens, and transient data are stored without exposing values.",
    inputSchema: storageStatusInputSchema
  }
] satisfies McpTool[];

export function getEnabledTools(mode: OutlookMcpMode): McpTool[] {
  if (mode === "calendar-write") {
    return toolDefinitions;
  }

  return toolDefinitions.filter((tool) => tool.name !== "outlook_create_calendar_block");
}

export function callTool(
  runtime: ToolRuntime,
  name: string,
  args: Record<string, unknown>
): Promise<ToolResult> {
  switch (name) {
    case "outlook_search_mail":
      return callMailService(runtime, (mailService) =>
        mailService.searchMail({
          query: parseRequiredString(args.query, "query"),
          startDate: parseOptionalString(args.startDate),
          endDate: parseOptionalString(args.endDate),
          maxResults: parseOptionalNumber(args.maxResults)
        })
      );
    case "outlook_list_mail":
      return callMailService(runtime, (mailService) =>
        mailService.listMail({
          startDate: parseRequiredString(args.startDate, "startDate"),
          endDate: parseRequiredString(args.endDate, "endDate"),
          folder: parseOptionalString(args.folder),
          maxResults: parseOptionalNumber(args.maxResults)
        })
      );
    case "outlook_read_mail":
      return callMailService(runtime, (mailService) =>
        mailService.readMail({ messageRef: parseRequiredString(args.messageRef, "messageRef") })
      );
    case "outlook_get_calendar_view":
      return callCalendarService(runtime, (calendarService) =>
        calendarService.getCalendarView({
          startDateTime: parseRequiredString(args.startDateTime, "startDateTime"),
          endDateTime: parseRequiredString(args.endDateTime, "endDateTime"),
          maxResults: parseOptionalNumber(args.maxResults)
        })
      );
    case "outlook_search_calendar":
      return callCalendarService(runtime, (calendarService) =>
        calendarService.searchCalendar({
          query: parseRequiredString(args.query, "query"),
          startDateTime: parseRequiredString(args.startDateTime, "startDateTime"),
          endDateTime: parseRequiredString(args.endDateTime, "endDateTime"),
          maxResults: parseOptionalNumber(args.maxResults)
        })
      );
    case "outlook_create_calendar_block": {
      const policy = evaluateCalendarBlockPolicy(runtime.mode, args);
      if (!policy.allowed) {
        return Promise.resolve(jsonResult({ error: policy.reason }, true));
      }

      return callCalendarService(runtime, (calendarService) =>
        calendarService.createCalendarBlock({
          subject: parseRequiredString(args.subject, "subject"),
          body: parseRequiredString(args.body, "body"),
          startDateTime: parseRequiredString(args.startDateTime, "startDateTime"),
          endDateTime: parseRequiredString(args.endDateTime, "endDateTime"),
          timeZone: parseOptionalString(args.timeZone),
          location: parseOptionalString(args.location)
        })
      );
    }
    case "auth_status":
      return callAuthManager(runtime, (authManager) => authManager.status());
    case "auth_login":
      return callAuthManager(runtime, (authManager) =>
        authManager.login({ flow: parseAuthFlow(args.flow) })
      );
    case "auth_complete":
      return callAuthManager(runtime, (authManager) =>
        authManager.complete(parseRequiredString(args.authSessionId, "authSessionId"))
      );
    case "auth_logout":
      return callAuthManager(runtime, (authManager) => authManager.logout());
    case "storage_status":
      if (runtime.tokenStore === undefined) {
        return Promise.resolve(jsonResult({ error: "Token store is not configured." }, true));
      }

      return Promise.resolve(jsonResult(getStorageStatus(runtime.tokenStore)));
    default:
      return Promise.resolve(jsonResult({ error: `Unknown tool: ${name}` }, true));
  }
}

async function callCalendarService(
  runtime: ToolRuntime,
  callback: (calendarService: CalendarService) => Promise<unknown>
): Promise<ToolResult> {
  const calendarService = runtime.calendarService ?? createCalendarService(runtime.authManager);
  if (calendarService === undefined) {
    return jsonResult({ error: "Authentication is required before using calendar tools." }, true);
  }

  try {
    return jsonResult(await callback(calendarService));
  } catch (error) {
    return jsonResult({ error: error instanceof Error ? error.message : "Calendar tool failed." }, true);
  }
}

function createCalendarService(authManager: AuthManager | undefined): CalendarService | undefined {
  if (authManager === undefined) {
    return undefined;
  }

  return new GraphCalendarService(createGraphClient(authManager));
}

async function callMailService(
  runtime: ToolRuntime,
  callback: (mailService: MailService) => Promise<unknown>
): Promise<ToolResult> {
  const mailService = runtime.mailService ?? createMailService(runtime.authManager);
  if (mailService === undefined) {
    return jsonResult({ error: "Authentication is required before using mail tools." }, true);
  }

  try {
    return jsonResult(await callback(mailService));
  } catch (error) {
    return jsonResult({ error: error instanceof Error ? error.message : "Mail tool failed." }, true);
  }
}

function createMailService(authManager: AuthManager | undefined): MailService | undefined {
  if (authManager === undefined) {
    return undefined;
  }

  return new GraphMailService(createGraphClient(authManager));
}

async function callAuthManager(
  runtime: ToolRuntime,
  callback: (authManager: AuthManager) => Promise<unknown>
): Promise<ToolResult> {
  if (runtime.authManager === undefined) {
    return jsonResult({ error: "Authentication manager is not configured." }, true);
  }

  try {
    return jsonResult(await callback(runtime.authManager));
  } catch (error) {
    return jsonResult({ error: error instanceof Error ? error.message : "Authentication failed." }, true);
  }
}

function parseAuthFlow(value: unknown): "pkce" | "device" | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (value === "pkce" || value === "device") {
    return value;
  }

  throw new Error('flow must be "pkce" or "device".');
}

function parseRequiredString(value: unknown, name: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`${name} is required.`);
  }

  return value;
}

function parseOptionalString(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function parseOptionalNumber(value: unknown): number | undefined {
  return typeof value === "number" ? value : undefined;
}

function jsonResult(payload: unknown, isError = false): ToolResult {
  const result: ToolResult = {
    content: [
      {
        type: "text",
        text: JSON.stringify(payload, null, 2)
      }
    ]
  };

  if (isError) {
    result.isError = true;
  }

  return result;
}
