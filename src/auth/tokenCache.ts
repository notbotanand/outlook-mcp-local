import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

export type TokenSet = {
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: number;
  scope?: string;
  tokenType?: string;
};

export type TokenStatus = {
  hasAccessToken: boolean;
  hasRefreshToken: boolean;
  expiresAt?: number;
  scope?: string;
};

export interface TokenStore {
  read(): Promise<TokenSet | undefined>;
  write(tokens: TokenSet): Promise<void>;
  clear(): Promise<void>;
  describe(): TokenStoreDescription;
}

export type TokenStoreDescription = {
  refreshTokenStore: string;
  keychainService?: string;
  keychainAccount?: string;
  accessTokenStore: "process memory";
};

export class InMemoryTokenStore implements TokenStore {
  private tokens: TokenSet | undefined;

  async read(): Promise<TokenSet | undefined> {
    return this.tokens;
  }

  async write(tokens: TokenSet): Promise<void> {
    this.tokens = { ...tokens };
  }

  async clear(): Promise<void> {
    this.tokens = undefined;
  }

  describe(): TokenStoreDescription {
    return {
      refreshTokenStore: "process memory",
      accessTokenStore: "process memory"
    };
  }
}

export class MacOsKeychainRefreshTokenStore implements TokenStore {
  private accessTokenCache: Omit<TokenSet, "refreshToken"> = {};

  constructor(
    private readonly service = "outlook-mcp-local",
    private readonly account = "default"
  ) {}

  async read(): Promise<TokenSet | undefined> {
    const refreshToken = await this.readRefreshToken();
    if (refreshToken === undefined && this.accessTokenCache.accessToken === undefined) {
      return undefined;
    }

    return {
      ...this.accessTokenCache,
      refreshToken
    };
  }

  async write(tokens: TokenSet): Promise<void> {
    const { refreshToken, ...accessTokenCache } = tokens;
    this.accessTokenCache = accessTokenCache;

    if (refreshToken !== undefined) {
      await this.writeRefreshToken(refreshToken);
    }
  }

  async clear(): Promise<void> {
    this.accessTokenCache = {};
    await execFileAsync("security", [
      "delete-generic-password",
      "-s",
      this.service,
      "-a",
      this.account
    ]).catch(() => undefined);
  }

  describe(): TokenStoreDescription {
    return {
      refreshTokenStore: "macOS Keychain",
      keychainService: this.service,
      keychainAccount: this.account,
      accessTokenStore: "process memory"
    };
  }

  private async readRefreshToken(): Promise<string | undefined> {
    try {
      const { stdout } = await execFileAsync("security", [
        "find-generic-password",
        "-w",
        "-s",
        this.service,
        "-a",
        this.account
      ]);
      const token = stdout.trim();
      return token.length > 0 ? token : undefined;
    } catch {
      return undefined;
    }
  }

  private async writeRefreshToken(refreshToken: string): Promise<void> {
    await execFileAsync("security", [
      "add-generic-password",
      "-U",
      "-s",
      this.service,
      "-a",
      this.account,
      "-w",
      refreshToken
    ]);
  }
}

export function tokenStatus(tokens: TokenSet | undefined): TokenStatus {
  return {
    hasAccessToken: tokens?.accessToken !== undefined,
    hasRefreshToken: tokens?.refreshToken !== undefined,
    expiresAt: tokens?.expiresAt,
    scope: tokens?.scope
  };
}

export function createDefaultTokenStore(platform = process.platform): TokenStore {
  if (platform === "darwin") {
    return new MacOsKeychainRefreshTokenStore();
  }

  return new InMemoryTokenStore();
}
