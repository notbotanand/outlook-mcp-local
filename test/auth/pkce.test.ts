import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { createPkcePair, createState } from "../../src/auth/pkce.js";

describe("PKCE helpers", () => {
  it("creates URL-safe verifier, challenge, and state", () => {
    const pkce = createPkcePair();
    const state = createState();

    assert.equal(pkce.method, "S256");
    assert.match(pkce.verifier, /^[A-Za-z0-9_-]+$/);
    assert.match(pkce.challenge, /^[A-Za-z0-9_-]+$/);
    assert.match(state, /^[A-Za-z0-9_-]+$/);
  });
});
