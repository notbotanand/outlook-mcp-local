export type OutlookMcpMode = "read" | "calendar-write";

const defaultMode: OutlookMcpMode = "read";

export function parseOutlookMcpMode(value: string | undefined): OutlookMcpMode {
  if (value === undefined || value.trim() === "") {
    return defaultMode;
  }

  if (value === "read" || value === "calendar-write") {
    return value;
  }

  throw new Error(
    `Unsupported OUTLOOK_MCP_MODE "${value}". Supported modes are "read" and "calendar-write".`
  );
}

export function parseOutlookMcpModeFromEnv(env: NodeJS.ProcessEnv): OutlookMcpMode {
  return parseOutlookMcpMode(env.OUTLOOK_MCP_MODE);
}

export function isCalendarWriteEnabled(mode: OutlookMcpMode): boolean {
  return mode === "calendar-write";
}
