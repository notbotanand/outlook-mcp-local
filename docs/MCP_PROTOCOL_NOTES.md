# MCP Protocol Notes

This project targets the official MCP specification version `2025-11-25`.

## stdio Transport

The current official stdio transport uses UTF-8 JSON-RPC messages over standard input and output.

For stdio:

- each message is one JSON-RPC request, notification, or response
- messages are delimited by newlines
- messages must not contain embedded newlines
- server logs must go to `stderr`
- `stdout` must contain only valid MCP messages

This server intentionally uses newline-delimited JSON-RPC for stdio. It does not use `Content-Length` framing for stdio.

## Implemented Hardening

- validates `jsonrpc: "2.0"`
- validates `method`, `id`, and `params` shape
- ignores JSON-RPC notifications without sending a response
- accepts inbound JSON-RPC responses and ignores them because this server does not issue client requests yet
- validates tool arguments against declared input schemas
- returns tool input validation failures as tool execution errors with `isError: true`
- enforces a 1 MB stdio line/message size limit
- maps Microsoft Graph errors to safe actionable messages

## Current Scope

The server exposes the `tools` capability only. It does not currently implement resources, prompts, logging notifications, completions, sampling, roots, or elicitation.
