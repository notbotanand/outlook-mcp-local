import { createHash, randomBytes } from "node:crypto";

export type PkcePair = {
  verifier: string;
  challenge: string;
  method: "S256";
};

export function createPkcePair(): PkcePair {
  const verifier = base64UrlEncode(randomBytes(32));
  const challenge = base64UrlEncode(createHash("sha256").update(verifier).digest());

  return {
    verifier,
    challenge,
    method: "S256"
  };
}

export function createState(): string {
  return base64UrlEncode(randomBytes(24));
}

function base64UrlEncode(value: Buffer): string {
  return value
    .toString("base64")
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replaceAll("=", "");
}
