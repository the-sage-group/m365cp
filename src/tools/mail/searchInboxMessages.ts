import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { SearchHit } from "@microsoft/microsoft-graph-types";

// ============================================================================
// Output Types
// ============================================================================

export interface SearchInboxMessagesResult {
  count: number;
  results: SearchHit[];
}

// ============================================================================
// Tool Definition
// ============================================================================

export const searchInboxMessages = {
  name: "search_inbox_messages",
  schema: {
    title: "Search Inbox Messages",
    description:
      "Search for messages in the inbox. Returns full message details including conversationId.",
    inputSchema: z.object({
      query: z
        .string()
        .describe(
          "Search query to filter messages (searches subject, body, sender)"
        ),
      top: z
        .number()
        .optional()
        .default(10)
        .describe("Maximum number of messages to return (default: 10)"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const searchHits = await client.searchMessages(args.query);

    const result: SearchInboxMessagesResult = {
      count: searchHits.length,
      results: searchHits,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  }) satisfies ToolCallback<{
    query: z.ZodString;
    top: z.ZodOptional<z.ZodNumber>;
  }>,
};
