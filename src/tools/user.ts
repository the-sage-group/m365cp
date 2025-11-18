import { createGraphClient } from "../graph/client.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { User } from "@microsoft/microsoft-graph-types";

/**
 * Register user-related tools
 */
export function registerUserTools(server: McpServer) {
  server.registerTool(
    "get_user_info",
    {
      title: "Get User Info",
      description: "Get information about the authenticated user",
      inputSchema: {},
    },
    async (_, extra) => {
      // Extract token from authInfo injected by requireBearerAuth middleware
      const accessToken = extra.authInfo?.token;
      if (!accessToken) {
        throw new Error("No access token available");
      }

      const client = createGraphClient(accessToken);
      const user: User = await client.api("/me").get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(user, null, 2),
          },
        ],
        structuredContent: user as Record<string, unknown>,
      };
    }
  );
}
