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
      "Get all messages and attachments from a conversation thread. Downloads all file attachments and uploads them to OneDrive. Use search_inbox_messages first to find the conversationId.",
    inputSchema: z.object({
      conversationId: z
        .string()
        .describe("The conversation ID from a message search result"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);

    // Get all messages in the conversation
    const messages = await client.getConversationMessages(args.conversationId);

    // Process messages and upload attachments
    const processedMessages = await Promise.all(
      messages.map(async (message) => {
        const uploadedAttachments = [];

        // Process file attachments
        for (const attachment of message.attachments ?? []) {
          // Check if this is a FileAttachment with contentBytes
          if (
            (attachment as any)["@odata.type"] ===
              "#microsoft.graph.fileAttachment" &&
            "contentBytes" in attachment
          ) {
            try {
              // Decode base64 content
              const contentBytes = Buffer.from(
                (attachment as any).contentBytes,
                "base64"
              );

              // Upload to OneDrive in folder: attachments/{conversationId-prefix}
              const conversationPrefix = args.conversationId.substring(0, 8);
              const uploadedFile = await client.uploadFile(
                attachment.name || "attachment",
                contentBytes,
                `attachments/${conversationPrefix}`
              );
              console.log("Uploaded file:", uploadedFile);
              uploadedAttachments.push({
                name: attachment.name,
                size: attachment.size,
                contentType: attachment.contentType,
                fileId: uploadedFile.id,
                webUrl: uploadedFile.webUrl,
              });
            } catch (error) {
              uploadedAttachments.push({
                name: attachment.name,
                size: attachment.size,
                contentType: attachment.contentType,
                error: "Failed to upload to OneDrive",
              });
            }
          }
        }

        return {
          id: message.id,
          subject: message.subject,
          from: message.from,
          receivedDateTime: message.receivedDateTime,
          bodyPreview: message.bodyPreview,
          body: message.body,
          uploadedAttachments,
        };
      })
    );

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(
            {
              count: processedMessages.length,
              results: processedMessages,
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
