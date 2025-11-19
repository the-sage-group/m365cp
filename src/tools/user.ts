import { createGraphClient } from "../graph/client.js";
import type { User } from "@microsoft/microsoft-graph-types";

export const getUserInfo = {
  name: "get_user_info",
  schema: {
    title: "Get User Info",
    description: "Get information about the authenticated user",
    inputSchema: {},
  },
  handler: async (_: unknown, extra: { authInfo?: { token?: string } }) => {
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
          type: "text" as const,
          text: JSON.stringify(user, null, 2),
        },
      ],
      structuredContent: user as Record<string, unknown>,
    };
  },
};
