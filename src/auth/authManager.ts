import { createServer, type Server } from "node:http";
import { randomUUID } from "node:crypto";
import { type AddressInfo } from "node:net";
import {
  buildAuthorizationUrl,
  exchangeAuthorizationCode,
  pollDeviceCodeToken,
  refreshAccessToken,
  requestDeviceCode,
  toTokenSet,
  type DeviceCodeResponse
} from "./oauthClient.js";
import { createPkcePair, createState, type PkcePair } from "./pkce.js";
import { type TokenStatus, type TokenStore, tokenStatus } from "./tokenCache.js";
import { type AuthConfig } from "../config/authConfig.js";
import { scopesForMode } from "../policy/scopePolicy.js";

export type AuthLoginInput = {
  flow?: "pkce" | "device";
};

export type AuthLoginResult =
  | {
      configured: true;
      authenticated: boolean;
      flow: "pkce";
      authSessionId: string;
      authorizationUrl: string;
      redirectUri: string;
      expiresInSeconds: number;
      instructions: string;
    }
  | {
      configured: true;
      authenticated: boolean;
      flow: "device";
      authSessionId: string;
      userCode: string;
      verificationUri: string;
      verificationUriComplete?: string;
      expiresInSeconds: number;
      intervalSeconds: number;
      instructions: string;
    };

export type AuthStatusResult = {
  configured: boolean;
  tenant: string;
  mode: string;
  authenticated: boolean;
  tokenStatus: TokenStatus;
};

type PendingPkce = {
  kind: "pkce";
  pkce: PkcePair;
  state: string;
  server: Server;
  expiresAt: number;
};

type PendingDevice = {
  kind: "device";
  deviceCode: string;
  intervalSeconds: number;
  expiresAt: number;
};

type PendingAuth = PendingPkce | PendingDevice;

const authContentWarning =
  "Authentication output never includes access tokens, refresh tokens, authorization codes, or raw Microsoft responses.";

export class AuthManager {
  private readonly pendingAuth = new Map<string, PendingAuth>();

  constructor(
    private readonly config: AuthConfig,
    private readonly tokenStore: TokenStore,
    private readonly fetchImpl: typeof fetch = fetch
  ) {}

  async status(): Promise<AuthStatusResult> {
    const tokens = await this.tokenStore.read();
    const status = tokenStatus(tokens);

    return {
      configured: this.config.clientId !== undefined,
      tenant: this.config.tenant,
      mode: this.config.mode,
      authenticated: status.hasRefreshToken || status.hasAccessToken,
      tokenStatus: status
    };
  }

  async login(input: AuthLoginInput = {}): Promise<AuthLoginResult> {
    this.assertConfigured();
    const flow = input.flow ?? "pkce";

    if (flow === "device") {
      return this.startDeviceLogin();
    }

    return this.startPkceLogin();
  }

  async complete(authSessionId: string): Promise<{
    authenticated: boolean;
    status: TokenStatus;
    contentWarning: string;
  }> {
    const pending = this.pendingAuth.get(authSessionId);
    if (pending === undefined) {
      throw new Error("Unknown or expired auth session.");
    }

    if (Date.now() > pending.expiresAt) {
      this.closePending(authSessionId, pending);
      throw new Error("Auth session expired. Start login again.");
    }

    if (pending.kind === "pkce") {
      throw new Error("PKCE login completes automatically through the local callback.");
    }

    const response = await pollDeviceCodeToken(this.config, pending.deviceCode, this.fetchImpl);
    if (response.error === "authorization_pending") {
      return {
        authenticated: false,
        status: tokenStatus(await this.tokenStore.read()),
        contentWarning: authContentWarning
      };
    }

    const tokens = toTokenSet(response);
    await this.tokenStore.write(tokens);
    this.pendingAuth.delete(authSessionId);

    return {
      authenticated: true,
      status: tokenStatus(tokens),
      contentWarning: authContentWarning
    };
  }

  async logout(): Promise<{ authenticated: false; contentWarning: string }> {
    for (const [authSessionId, pending] of this.pendingAuth.entries()) {
      this.closePending(authSessionId, pending);
    }

    await this.tokenStore.clear();
    return {
      authenticated: false,
      contentWarning: authContentWarning
    };
  }

  async getAccessToken(): Promise<string> {
    const tokens = await this.tokenStore.read();
    if (
      tokens?.accessToken !== undefined &&
      tokens.expiresAt !== undefined &&
      tokens.expiresAt > Date.now() + 60_000
    ) {
      return tokens.accessToken;
    }

    if (tokens?.refreshToken === undefined) {
      throw new Error("Not authenticated. Run auth_login first.");
    }

    const refreshed = await refreshAccessToken(
      this.config,
      scopesForMode(this.config.mode),
      tokens.refreshToken,
      this.fetchImpl
    );
    await this.tokenStore.write({
      ...refreshed,
      refreshToken: refreshed.refreshToken ?? tokens.refreshToken
    });

    if (refreshed.accessToken === undefined) {
      throw new Error("Microsoft did not return an access token.");
    }

    return refreshed.accessToken;
  }

  private async startPkceLogin(): Promise<AuthLoginResult> {
    const pkce = createPkcePair();
    const state = createState();
    const authSessionId = randomUUID();
    const expiresAt = Date.now() + 10 * 60 * 1000;
    const redirectUri = await this.createPkceCallbackServer(authSessionId, pkce, state, expiresAt);
    const authorizationUrl = buildAuthorizationUrl(
      { ...this.config, redirectUri },
      scopesForMode(this.config.mode),
      pkce,
      state
    );

    return {
      configured: true,
      authenticated: false,
      flow: "pkce",
      authSessionId,
      authorizationUrl,
      redirectUri,
      expiresInSeconds: 600,
      instructions:
        "Open authorizationUrl in a browser. After Microsoft redirects to the local callback, auth_status will show whether login succeeded."
    };
  }

  private async createPkceCallbackServer(
    authSessionId: string,
    pkce: PkcePair,
    state: string,
    expiresAt: number
  ): Promise<string> {
    const configuredUrl = new URL(this.config.redirectUri);
    if (configuredUrl.protocol !== "http:") {
      throw new Error("OUTLOOK_MCP_REDIRECT_URI must be an http:// localhost URL.");
    }

    const host = configuredUrl.hostname === "localhost" ? "127.0.0.1" : configuredUrl.hostname;
    if (host !== "127.0.0.1") {
      throw new Error("OUTLOOK_MCP_REDIRECT_URI must bind to localhost.");
    }

    const server = createServer(async (request, response) => {
      try {
        const requestUrl = new URL(request.url ?? "/", `${configuredUrl.protocol}//${configuredUrl.host}`);
        if (requestUrl.pathname !== configuredUrl.pathname) {
          response.writeHead(404).end("Not found");
          return;
        }

        const returnedState = requestUrl.searchParams.get("state");
        const code = requestUrl.searchParams.get("code");
        const error = requestUrl.searchParams.get("error");

        if (error !== null) {
          response.writeHead(400).end("Microsoft authentication failed. You can close this tab.");
          this.closePending(authSessionId, { kind: "pkce", pkce, state, server, expiresAt });
          return;
        }

        if (returnedState !== state || code === null) {
          response.writeHead(400).end("Invalid authentication callback. You can close this tab.");
          return;
        }

        const tokens = await exchangeAuthorizationCode(
          { ...this.config, redirectUri: this.callbackRedirectUri(server, configuredUrl) },
          scopesForMode(this.config.mode),
          code,
          pkce.verifier,
          this.fetchImpl
        );
        await this.tokenStore.write(tokens);
        response.writeHead(200).end("Authentication complete. You can close this tab.");
        this.closePending(authSessionId, { kind: "pkce", pkce, state, server, expiresAt });
      } catch {
        response.writeHead(500).end("Authentication failed. You can close this tab.");
        this.closePending(authSessionId, this.pendingAuth.get(authSessionId));
      }
    });

    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.listen(Number(configuredUrl.port), host, () => {
        server.off("error", reject);
        resolve();
      });
    });

    this.pendingAuth.set(authSessionId, { kind: "pkce", pkce, state, server, expiresAt });
    server.setTimeout(10 * 60 * 1000, () => this.closePending(authSessionId, this.pendingAuth.get(authSessionId)));

    return this.callbackRedirectUri(server, configuredUrl);
  }

  private async startDeviceLogin(): Promise<AuthLoginResult> {
    const response: DeviceCodeResponse = await requestDeviceCode(
      this.config,
      scopesForMode(this.config.mode),
      this.fetchImpl
    );
    const authSessionId = randomUUID();
    const intervalSeconds = response.interval ?? 5;

    this.pendingAuth.set(authSessionId, {
      kind: "device",
      deviceCode: response.device_code,
      intervalSeconds,
      expiresAt: Date.now() + response.expires_in * 1000
    });

    return {
      configured: true,
      authenticated: false,
      flow: "device",
      authSessionId,
      userCode: response.user_code,
      verificationUri: response.verification_uri,
      verificationUriComplete: response.verification_uri_complete,
      expiresInSeconds: response.expires_in,
      intervalSeconds,
      instructions:
        response.message ??
        "Open verificationUri and enter userCode, then call auth_complete with authSessionId."
    };
  }

  private callbackRedirectUri(server: Server, configuredUrl: URL): string {
    const address = server.address() as AddressInfo;
    configuredUrl.port = String(address.port);
    configuredUrl.hostname = "127.0.0.1";
    return configuredUrl.toString();
  }

  private closePending(authSessionId: string, pending: PendingAuth | undefined): void {
    if (pending?.kind === "pkce") {
      pending.server.close();
    }

    this.pendingAuth.delete(authSessionId);
  }

  private assertConfigured(): void {
    if (this.config.clientId === undefined) {
      throw new Error("OUTLOOK_MCP_CLIENT_ID is required for Microsoft authentication.");
    }
  }
}
