import type { OAuthTokenVerifier } from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import jwt from "jsonwebtoken";

/**
 * Pass-through token verifier for Microsoft Graph tokens.
 * We don't verify the token ourselves - Microsoft Graph API will do that.
 * This extracts the expiration time and passes the token through to the tools.
 */
export class MicrosoftGraphTokenVerifier implements OAuthTokenVerifier {
  async verifyAccessToken(token: string): Promise<AuthInfo> {
    // Decode JWT without verification (Microsoft Graph API will verify it)
    const decoded = jwt.decode(token) as any;

    if (!decoded || typeof decoded !== "object") {
      throw new Error("Invalid JWT token");
    }

    // Extract expiration time (exp claim is in seconds since epoch)
    const expiresAt = decoded.exp;
    if (typeof expiresAt !== "number") {
      throw new Error("Token missing exp claim");
    }

    // Extract scopes from scp or scope claim
    const scopes = decoded.scp
      ? decoded.scp.split(" ")
      : decoded.scope
        ? decoded.scope.split(" ")
        : [];

    return {
      token,
      clientId: decoded.appid || decoded.client_id || "microsoft-graph",
      scopes,
      expiresAt,
    };
  }
}
