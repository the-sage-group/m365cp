import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type {
  Message,
  DriveItem,
  Attachment,
} from "@microsoft/microsoft-graph-types";

// ============================================================================
// Output Types
// ============================================================================

export interface UploadedAttachment {
  attachment: Attachment;
  driveItem: DriveItem;
  error?: string;
}

export interface ConversationMessage
  extends Pick<
    Message,
    "id" | "subject" | "from" | "receivedDateTime" | "bodyPreview" | "body"
  > {
  uploadedAttachments: UploadedAttachment[];
}

export interface GetConversationResult {
  count: number;
  results: ConversationMessage[];
}

// ============================================================================
// Tool Definition
// ============================================================================

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
        const uploadedAttachments: UploadedAttachment[] = [];

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
              const attachmentBytes = Buffer.from(
                (attachment as any).contentBytes,
                "base64"
              );

              // Upload to OneDrive in folder: attachments/{conversationId-prefix}
              const conversationPrefix = args.conversationId.substring(0, 8);
              const uploadedFile = await client.uploadFile(
                attachment.name || "attachment",
                attachmentBytes,
                `attachments/${conversationPrefix}`
              );

              // Strip contentBytes from attachment to avoid sending large data in response
              const { contentBytes, ...attachmentMetadata } = attachment as any;
              uploadedAttachments.push({
                attachment: attachmentMetadata,
                driveItem: uploadedFile,
              });
            } catch (error) {
              const { contentBytes, ...attachmentMetadata } = attachment as any;
              uploadedAttachments.push({
                attachment: attachmentMetadata,
                driveItem: {} as DriveItem,
                error: "Failed to upload to OneDrive",
              });
            }
          }
        }

        const conversationMessage: ConversationMessage = {
          id: message.id,
          subject: message.subject,
          from: message.from,
          receivedDateTime: message.receivedDateTime,
          bodyPreview: message.bodyPreview,
          body: message.body,
          uploadedAttachments,
        };

        return conversationMessage;
      })
    );

    const result: GetConversationResult = {
      count: processedMessages.length,
      results: processedMessages,
    };

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  }) satisfies ToolCallback<{ conversationId: z.ZodString }>,
};
