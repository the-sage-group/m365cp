import { Client } from "@microsoft/microsoft-graph-client";
import type { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

/**
 * Custom authentication provider for user-provided access tokens
 */
class AccessTokenAuthProvider implements AuthenticationProvider {
  private accessToken: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  async getAccessToken(): Promise<string> {
    return this.accessToken;
  }
}

/**
 * Create a Microsoft Graph client with a user-provided access token
 * Following best practices with middleware-based initialization
 */
export function createGraphClient(accessToken: string): Client {
  const authProvider = new AccessTokenAuthProvider(accessToken);

  return Client.initWithMiddleware({
    authProvider,
  });
}
