import { chmodSync, existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export type EnvMap = Record<string, string>;

export function parseEnvFile(source: string): EnvMap {
  const values: EnvMap = {};

  for (const line of source.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (trimmed.length === 0 || trimmed.startsWith("#")) {
      continue;
    }

    const equalsIndex = trimmed.indexOf("=");
    if (equalsIndex <= 0) {
      continue;
    }

    const key = trimmed.slice(0, equalsIndex).trim();
    const rawValue = trimmed.slice(equalsIndex + 1).trim();
    values[key] = unquoteEnvValue(rawValue);
  }

  return values;
}

export function defaultEnvFilePath(): string {
  const moduleDir = dirname(fileURLToPath(import.meta.url));
  const distRelativePath = resolve(moduleDir, "../../..", ".env");
  if (existsSync(distRelativePath)) {
    return distRelativePath;
  }

  return resolve(".env");
}

export function loadEnvFile(path = defaultEnvFilePath(), env: NodeJS.ProcessEnv = process.env): EnvMap {
  env.OUTLOOK_MCP_ENV_FILE = path;
  if (!existsSync(path)) {
    return {};
  }

  const values = parseEnvFile(readFileSync(path, "utf8"));
  for (const [key, value] of Object.entries(values)) {
    if (env[key] === undefined) {
      env[key] = value;
    }
  }

  return values;
}

export function writeEnvFile(path: string, values: EnvMap): void {
  const lines = [
    "# Local outlook-mcp-local configuration.",
    "# Do not commit this file.",
    "",
    `OUTLOOK_MCP_CLIENT_ID=${quoteEnvValue(values.OUTLOOK_MCP_CLIENT_ID)}`,
    `OUTLOOK_MCP_TENANT=${quoteEnvValue(values.OUTLOOK_MCP_TENANT)}`,
    `OUTLOOK_MCP_REDIRECT_URI=${quoteEnvValue(values.OUTLOOK_MCP_REDIRECT_URI)}`,
    `OUTLOOK_MCP_MODE=${quoteEnvValue(values.OUTLOOK_MCP_MODE)}`,
    ""
  ];

  writeFileSync(path, lines.join("\n"), { mode: 0o600 });
  chmodSync(path, 0o600);
}

function unquoteEnvValue(value: string): string {
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    return value.slice(1, -1);
  }

  return value;
}

function quoteEnvValue(value: string | undefined): string {
  return JSON.stringify(value ?? "");
}
