import assert from "node:assert/strict";
import { PassThrough } from "node:stream";
import { describe, it } from "node:test";
import { createMcpServer } from "../../src/mcp/server.js";
import { createLogger } from "../../src/util/logger.js";

describe("MCP server protocol handling", () => {
  it("uses newline-delimited JSON-RPC stdio messages per MCP 2025-11-25", async () => {
    const { input, output, nextMessage } = startTestServer();

    input.write(
      `${JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2025-11-25",
          capabilities: {},
          clientInfo: { name: "SyntheticClient", version: "0.0.0" }
        }
      })}\n`
    );

    const response = await nextMessage();
    assert.equal(response.jsonrpc, "2.0");
    assert.equal(response.id, 1);
    assert.equal(response.result.protocolVersion, "2025-11-25");
    assert.deepEqual(response.result.capabilities, { tools: {} });
    output.destroy();
  });

  it("rejects malformed JSON-RPC requests", async () => {
    const { input, output, nextMessage } = startTestServer();

    input.write(`${JSON.stringify({ id: 1, method: "tools/list", params: {} })}\n`);

    const response = await nextMessage();
    assert.equal(response.error.code, -32600);
    assert.match(response.error.message, /JSON-RPC version/);
    output.destroy();
  });

  it("validates tool arguments against declared schemas", async () => {
    const { input, output, nextMessage } = startTestServer();

    input.write(
      `${JSON.stringify({
        jsonrpc: "2.0",
        id: 2,
        method: "tools/call",
        params: {
          name: "outlook_search_mail",
          arguments: { query: "Contoso", maxResults: 500 }
        }
      })}\n`
    );

    const response = await nextMessage();
    assert.equal(response.result.isError, true);
    assert.match(response.result.content[0].text, /above the maximum/);
    output.destroy();
  });
});

function startTestServer(): {
  input: PassThrough;
  output: PassThrough;
  nextMessage: () => Promise<Record<string, any>>;
} {
  const input = new PassThrough();
  const output = new PassThrough();
  const logger = createLogger(new PassThrough());
  createMcpServer({ mode: "read", logger }).start(input, output);

  let buffer = "";
  const pending: Array<(message: Record<string, any>) => void> = [];

  output.on("data", (chunk: Buffer) => {
    buffer += chunk.toString("utf8");
    let newlineIndex = buffer.indexOf("\n");
    while (newlineIndex >= 0) {
      const line = buffer.slice(0, newlineIndex);
      buffer = buffer.slice(newlineIndex + 1);
      pending.shift()?.(JSON.parse(line) as Record<string, any>);
      newlineIndex = buffer.indexOf("\n");
    }
  });

  return {
    input,
    output,
    nextMessage: () =>
      new Promise((resolve) => {
        pending.push(resolve);
      })
  };
}
