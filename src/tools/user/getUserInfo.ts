import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { User } from "@microsoft/microsoft-graph-types";
import { toolNames } from "../names.js";

// ============================================================================
// Output Types
// ============================================================================

export type GetUserInfoResult = User;

// ============================================================================
// Tool Definition
// ============================================================================

export const getUserInfo = {
  name: toolNames.getUserInfo,
  schema: {
    title: "Get User Info",
    description: "Get information about the authenticated user",
    inputSchema: {},
  },
  handler: (async (_, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const user = await client.getUser();

    const result: GetUserInfoResult = user;

    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  }) satisfies ToolCallback<{}>,
};
