import { type AuthConfig, assertClientIdConfigured } from "../config/authConfig.js";
import { type PkcePair } from "./pkce.js";
import { type TokenSet } from "./tokenCache.js";

export type OAuthTokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  token_type?: string;
  error?: string;
  error_description?: string;
};

export type DeviceCodeResponse = {
  device_code: string;
  user_code: string;
  verification_uri: string;
  verification_uri_complete?: string;
  expires_in: number;
  interval?: number;
  message?: string;
};

export function buildAuthorizationUrl(
  config: AuthConfig,
  scopes: readonly string[],
  pkce: PkcePair,
  state: string
): string {
  const clientId = assertClientIdConfigured(config);
  const url = new URL(`https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/authorize`);
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", config.redirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", scopes.join(" "));
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", pkce.challenge);
  url.searchParams.set("code_challenge_method", pkce.method);
  return url.toString();
}

export async function requestDeviceCode(
  config: AuthConfig,
  scopes: readonly string[],
  fetchImpl: typeof fetch = fetch
): Promise<DeviceCodeResponse> {
  const clientId = assertClientIdConfigured(config);
  const response = await postForm(
    `https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/devicecode`,
    {
      client_id: clientId,
      scope: scopes.join(" ")
    },
    fetchImpl
  );

  return response as DeviceCodeResponse;
}

export async function exchangeAuthorizationCode(
  config: AuthConfig,
  scopes: readonly string[],
  code: string,
  verifier: string,
  fetchImpl: typeof fetch = fetch
): Promise<TokenSet> {
  const clientId = assertClientIdConfigured(config);
  const response = (await postForm(
    `https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/token`,
    {
      client_id: clientId,
      scope: scopes.join(" "),
      code,
      redirect_uri: config.redirectUri,
      grant_type: "authorization_code",
      code_verifier: verifier
    },
    fetchImpl
  )) as OAuthTokenResponse;

  return toTokenSet(response);
}

export async function pollDeviceCodeToken(
  config: AuthConfig,
  deviceCode: string,
  fetchImpl: typeof fetch = fetch
): Promise<OAuthTokenResponse> {
  const clientId = assertClientIdConfigured(config);
  return (await postForm(
    `https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/token`,
    {
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: clientId,
      device_code: deviceCode
    },
    fetchImpl,
    true
  )) as OAuthTokenResponse;
}

export async function refreshAccessToken(
  config: AuthConfig,
  scopes: readonly string[],
  refreshToken: string,
  fetchImpl: typeof fetch = fetch
): Promise<TokenSet> {
  const clientId = assertClientIdConfigured(config);
  const response = (await postForm(
    `https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/token`,
    {
      client_id: clientId,
      scope: scopes.join(" "),
      refresh_token: refreshToken,
      grant_type: "refresh_token"
    },
    fetchImpl
  )) as OAuthTokenResponse;

  return toTokenSet(response);
}

export function toTokenSet(response: OAuthTokenResponse, now = Date.now()): TokenSet {
  if (response.error !== undefined) {
    throw new Error(response.error_description ?? response.error);
  }

  return {
    accessToken: response.access_token,
    refreshToken: response.refresh_token,
    expiresAt:
      response.expires_in !== undefined ? now + response.expires_in * 1000 : undefined,
    scope: response.scope,
    tokenType: response.token_type
  };
}

async function postForm(
  url: string,
  fields: Record<string, string>,
  fetchImpl: typeof fetch,
  allowOAuthError = false
): Promise<unknown> {
  const body = new URLSearchParams(fields);
  const response = await fetchImpl(url, {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded"
    },
    body
  });

  const payload = (await response.json()) as OAuthTokenResponse | DeviceCodeResponse;
  if (!response.ok && !allowOAuthError) {
    const message =
      "error_description" in payload && payload.error_description
        ? payload.error_description
        : `Microsoft authentication request failed with status ${response.status}.`;
    throw new Error(message);
  }

  return payload;
}
