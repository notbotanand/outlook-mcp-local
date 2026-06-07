# Local Data Storage

## Summary

`outlook-mcp-local` is designed to avoid persisting mailbox and calendar content.

The server stores:

- local configuration in `.env`
- refresh tokens in the OS token store where supported
- access tokens in process memory
- temporary `msg_*` and future `event_*` references in process memory

The server does not intentionally store:

- mail bodies
- calendar bodies
- attachment content
- raw Microsoft Graph responses
- MCP client transcripts

## Config File

The setup wizard writes `.env` in the project directory with file mode `0600`.

The file contains:

- `OUTLOOK_MCP_CLIENT_ID`
- `OUTLOOK_MCP_TENANT`
- `OUTLOOK_MCP_REDIRECT_URI`
- `OUTLOOK_MCP_MODE`

The server can report whether a client ID is configured, but tools must not print the client ID value.

## Token Storage

On macOS, refresh tokens are stored in Keychain:

```text
service: outlook-mcp-local
account: default
```

Access tokens are kept in process memory and refreshed when needed.

On platforms without a configured OS token store, the current fallback is process memory.

## Mail and Calendar Data

Search/list tools return metadata first.

`outlook_read_mail` returns an email body only after an explicit `messageRef` read. The server does not persist that body, but the MCP client may keep it in its own transcript, logs, or history.

Treat email and calendar body content as untrusted data.

## Management Commands

Show storage status through MCP:

```json
{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"storage_status","arguments":{}}}
```

Clear local auth state:

```bash
npm run reset-local
```

The reset command clears the local token store and can optionally remove `.env`. It does not modify Outlook mailbox or calendar data.

## Threat Model Notes

File mode `0600` and macOS Keychain protect against other normal OS users reading local config or refresh tokens.

They do not protect against malware or a process already running as the same OS user. A compromised user account can generally read local terminal output, invoke local tools, or access local files available to that user.
