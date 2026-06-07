export type Logger = {
  info(message: string, context?: Record<string, unknown>): void;
  warn(message: string, context?: Record<string, unknown>): void;
  error(message: string, context?: Record<string, unknown>): void;
};

type LogLevel = "info" | "warn" | "error";

const redactedKeys = new Set([
  "accessToken",
  "refreshToken",
  "authorizationCode",
  "clientSecret",
  "body",
  "bodyText",
  "rawGraphResponse"
]);

export function createLogger(stream: NodeJS.WritableStream = process.stderr): Logger {
  return {
    info: (message, context) => writeLog(stream, "info", message, context),
    warn: (message, context) => writeLog(stream, "warn", message, context),
    error: (message, context) => writeLog(stream, "error", message, context)
  };
}

function writeLog(
  stream: NodeJS.WritableStream,
  level: LogLevel,
  message: string,
  context?: Record<string, unknown>
): void {
  const entry = {
    level,
    message,
    context: context ? sanitizeContext(context) : undefined
  };

  stream.write(`${JSON.stringify(entry)}\n`);
}

function sanitizeContext(context: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(context).map(([key, value]) => [
      key,
      redactedKeys.has(key) ? "[redacted]" : value
    ])
  );
}
