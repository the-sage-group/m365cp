import { GraphClient } from "../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";

export const getUserInfo = {
  name: "get_user_info",
  schema: {
    title: "Get User Info",
    description: "Get information about the authenticated user",
    inputSchema: {},
  },
  handler: (async (_, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const user = await client.getUser();

    return {
      content: [{ type: "text", text: JSON.stringify(user, null, 2) }],
    };
  }) satisfies ToolCallback<{}>,
};

export default [getUserInfo];
