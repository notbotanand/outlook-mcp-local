#!/usr/bin/env node
import { existsSync, rmSync } from "node:fs";
import { createInterface } from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { createDefaultTokenStore } from "../auth/tokenCache.js";
import { loadEnvFile } from "../config/envFile.js";

async function main(): Promise<void> {
  loadEnvFile();
  const rl = createInterface({ input, output });
  try {
    output.write("outlook-mcp-local local reset\n\n");
    output.write("This clears the local refresh token store. It does not touch Outlook mailbox or calendar data.\n");

    const clearConfig = await ask(rl, "Also remove local .env config? [y/N]", "N");
    await createDefaultTokenStore().clear();
    output.write("Cleared local token store.\n");

    if ((clearConfig.toLowerCase() === "y" || clearConfig.toLowerCase() === "yes") && existsSync(".env")) {
      rmSync(".env");
      output.write("Removed .env.\n");
    }
  } finally {
    rl.close();
  }
}

async function ask(
  rl: ReturnType<typeof createInterface>,
  question: string,
  defaultValue: string
): Promise<string> {
  const answer = (await rl.question(`${question} (${defaultValue}): `)).trim();
  return answer.length > 0 ? answer : defaultValue;
}

main().catch((error: unknown) => {
  output.write(`Reset failed: ${error instanceof Error ? error.message : "Unknown error"}\n`);
  process.exitCode = 1;
});
