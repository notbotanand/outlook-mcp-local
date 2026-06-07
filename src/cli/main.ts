#!/usr/bin/env node

const command = process.argv[2] ?? "server";

switch (command) {
  case "server":
    await import("../index.js");
    break;
  case "setup":
    await import("./setup.js");
    break;
  case "doctor":
    process.argv.splice(2, 1);
    await import("./doctor.js");
    break;
  case "reset-local":
    await import("./resetLocal.js");
    break;
  default:
    process.stdout.write(
      [
        "Usage: outlook-mcp-local <command>",
        "",
        "Commands:",
        "  server       Run the stdio MCP server",
        "  setup        Run local onboarding and PKCE auth",
        "  doctor       Check local readiness without printing secrets",
        "  reset-local  Clear local token state and optionally remove .env",
        ""
      ].join("\n")
    );
    process.exitCode = 1;
}
