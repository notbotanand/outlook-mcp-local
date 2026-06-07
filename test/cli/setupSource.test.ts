import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";

describe("setup source", () => {
  it("does not prompt users to choose a redirect URI", () => {
    const source = readFileSync("src/cli/setup.ts", "utf8");

    assert.match(source, /Redirect URI to register/);
    assert.equal(source.includes('ask(\n      rl,\n      "Redirect URI"'), false);
  });
});
