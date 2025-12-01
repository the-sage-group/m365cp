import { z } from "zod";
import { GraphClient } from "../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";

// ============================================================================
// Tools
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

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(
            {
              count: searchHits.length,
              results: searchHits,
            },
            null,
            2
          ),
        },
      ],
    };
  }) satisfies ToolCallback<{
    query: z.ZodString;
    top: z.ZodOptional<z.ZodNumber>;
  }>,
};

export const getConversation = {
  name: "get_conversation",
  schema: {
    title: "Get Conversation",
    description:
      "Get all messages and attachments from a conversation thread. Downloads all attachments and uploads them to Anthropic for analysis. Use search_inbox_messages first to find the conversationId.",
    inputSchema: z.object({
      conversationId: z
        .string()
        .describe("The conversation ID from a message search result"),
    }),
  },
  handler: (async (args, extra) => {
    console.log("Trying to get conversation", args);
    const client = new GraphClient(extra.authInfo!.token!);

    // Get all messages in the conversation
    const messages = await client.getConversationMessages(args.conversationId);

    console.log("Got conversation messages", messages.length);
    for (const message of messages) {
      // print a nice summary of each message
      console.log("Message: ", message.from);
      console.log("Subject: ", message.subject);
      console.log("Size: ", message.body?.content?.length);
      console.log("Attachments: ", message.attachments);
    }

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(
            {
              count: messages.length,
              results: messages,
            },
            null,
            2
          ),
        },
      ],
    };
  }) satisfies ToolCallback<{ conversationId: z.ZodString }>,
};

export default [searchInboxMessages, getConversation];
