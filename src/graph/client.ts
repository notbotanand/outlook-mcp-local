import { type AuthManager } from "../auth/authManager.js";

export type GraphClient = {
  get<T>(path: string, query?: Record<string, string | number | undefined>): Promise<T>;
  post<T>(path: string, body: unknown): Promise<T>;
};

export function createGraphClient(
  authManager: AuthManager,
  fetchImpl: typeof fetch = fetch
): GraphClient {
  return {
    async get<T>(
      path: string,
      query?: Record<string, string | number | undefined>
    ): Promise<T> {
      const accessToken = await authManager.getAccessToken();
      const url = buildGraphUrl(path, query);
      const response = await fetchImpl(url, {
        method: "GET",
        headers: {
          authorization: `Bearer ${accessToken}`,
          accept: "application/json"
        }
      });

      if (!response.ok) {
        throw new Error(mapGraphError(response.status));
      }

      return (await response.json()) as T;
    },

    async post<T>(path: string, body: unknown): Promise<T> {
      const accessToken = await authManager.getAccessToken();
      const response = await fetchImpl(buildGraphUrl(path), {
        method: "POST",
        headers: {
          authorization: `Bearer ${accessToken}`,
          accept: "application/json",
          "content-type": "application/json"
        },
        body: JSON.stringify(body)
      });

      if (!response.ok) {
        throw new Error(mapGraphError(response.status));
      }

      return (await response.json()) as T;
    }
  };
}

export function mapGraphError(status: number): string {
  switch (status) {
    case 400:
      return "Microsoft Graph rejected the request. Check date ranges, query text, and tool arguments.";
    case 401:
      return "Microsoft Graph authentication failed. Run auth_login again.";
    case 403:
      return "Microsoft Graph denied access. Re-run setup/auth with the required read or calendar-write mode.";
    case 404:
      return "Microsoft Graph item was not found. The local reference may be stale.";
    case 429:
      return "Microsoft Graph throttled the request. Try again later.";
    default:
      if (status >= 500) {
        return "Microsoft Graph is temporarily unavailable. Try again later.";
      }

      return `Microsoft Graph request failed with status ${status}.`;
  }
}

export function buildGraphUrl(
  path: string,
  query: Record<string, string | number | undefined> = {}
): string {
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  const url = new URL(`https://graph.microsoft.com/v1.0${normalizedPath}`);

  for (const [key, value] of Object.entries(query)) {
    if (value !== undefined) {
      url.searchParams.set(key, String(value));
    }
  }

  return url.toString();
}
