# Claude Code Setup

This repository provides a local stdio MCP server for personal Outlook.com mail and calendar.

## Agent Checklist

An agent setting this up for a user should follow these steps:

1. Clone the repository.
2. Run:

```bash
npm install
npm run setup
```

3. Let the human complete Microsoft app registration and PKCE browser auth when prompted.
   The setup wizard displays the redirect URI to register. Do not choose a different redirect URI unless intentionally using `OUTLOOK_MCP_REDIRECT_URI` as an advanced override.
4. Run:

```bash
npm run doctor
```

5. If doctor passes, add the MCP server to Claude Code.

## Claude Code Command

Use the command printed by `npm run doctor`.

It will look like:

```bash
claude mcp add --transport stdio outlook-mcp-local -- node /absolute/path/to/outlook-mcp-local/dist/src/index.js
```

The absolute path matters because Claude Code may launch the server from a different working directory.

## Manual `.mcp.json` Example

If using a project-scoped MCP config, use an absolute path:

```json
{
  "mcpServers": {
    "outlook-mcp-local": {
      "type": "stdio",
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp-local/dist/src/index.js"],
      "env": {}
    }
  }
}
```

The server loads its local `.env` file on startup. Do not put client IDs or tokens in the MCP config.

## Useful Commands

```bash
npm run setup
npm run doctor
npm run server
npm run reset-local
```

For a packaged install, the equivalent commands are:

```bash
outlook-mcp-local setup
outlook-mcp-local doctor
outlook-mcp-local server
outlook-mcp-local reset-local
```

If doctor reports `.env mode is 644; expected 600`, rerun setup or run:

```bash
chmod 600 .env
```

## Safety Notes

- The setup wizard does not print the configured Microsoft client ID.
- `doctor` does not print client IDs, access tokens, refresh tokens, mail bodies, or calendar bodies.
- Search/list tools return metadata first.
- Full email body content is returned only by `outlook_read_mail`.
- Calendar write requires `OUTLOOK_MCP_MODE=calendar-write` and only creates personal blocks.
