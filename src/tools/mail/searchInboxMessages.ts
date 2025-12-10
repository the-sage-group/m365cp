import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { SearchHit, Message } from "@microsoft/microsoft-graph-types";
import { toolNames } from "../names.js";

// ============================================================================
// Output Types
// ============================================================================

export interface MessageResult {
  id: string;
  conversationId?: string | null;
  subject?: string | null;
  from?: string | null;
  receivedDateTime?: string | null;
  bodyPreview?: string | null;
  webLink?: string | null;
}

export interface SearchInboxMessagesResult {
  count: number;
  results: MessageResult[];
}

// ============================================================================
// Tool Definition
// ============================================================================

export const searchInboxMessages = {
  name: toolNames.searchInboxMessages,
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

    const results: MessageResult[] = searchHits.map((hit) => {
      const message = hit.resource as Message;
      return {
        id: message.id!,
        conversationId: message.conversationId ?? undefined,
        subject: message.subject ?? undefined,
        from: message.from?.emailAddress?.address ?? undefined,
        receivedDateTime: message.receivedDateTime ?? undefined,
        bodyPreview: message.bodyPreview ?? undefined,
        webLink: message.webLink ?? undefined,
      };
    });

    const result: SearchInboxMessagesResult = {
      count: results.length,
      results,
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
