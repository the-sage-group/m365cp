import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type {
  Message,
  DriveItem,
  Attachment,
} from "@microsoft/microsoft-graph-types";
import puppeteer from "puppeteer";
import sanitize from "sanitize-filename";
import { toolNames } from "../names.js";

// ============================================================================
// Output Types
// ============================================================================

export interface DriveItemResult {
  id: string;
  name?: string | null;
  webUrl?: string | null;
  downloadUrl?: string | null;
  size?: number | null;
  mimeType?: string | null;
}

export interface AttachmentResult {
  id?: string | null;
  name?: string | null;
  contentType?: string | null;
  size?: number | null;
}

export interface UploadedAttachment {
  attachment: AttachmentResult;
  driveItem: DriveItemResult;
  error?: string;
}

export interface ConversationMessage {
  id?: string | null;
  subject?: string | null;
  from?: string | null;
  receivedDateTime?: string | null;
  bodyPreview?: string | null;
  uploadedAttachments: UploadedAttachment[];
}

export interface GetConversationResult {
  count: number;
  results: ConversationMessage[];
}

// ============================================================================
// Helper Functions
// ============================================================================

async function convertHtmlToPdf(message: Message): Promise<Buffer> {
  let result = message.body?.content ?? "";
  for (const att of message.attachments ?? []) {
    const fileAtt = att as any;
    if (
      fileAtt["@odata.type"] === "#microsoft.graph.fileAttachment" &&
      fileAtt.contentBytes &&
      fileAtt.contentId
    ) {
      const dataUrl = `data:${att.contentType || "image/png"};base64,${
        fileAtt.contentBytes
      }`;
      result = result.replace(
        new RegExp(`cid:${fileAtt.contentId}`, "g"),
        dataUrl
      );
    }
  }
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.setContent(result, { waitUntil: "networkidle0" });
  const pdf = await page.pdf({
    format: "A4",
    printBackground: true,
  });
  await browser.close();
  return Buffer.from(pdf);
}

// ============================================================================
// Tool Definition
// ============================================================================

export const getConversation = {
  name: toolNames.getConversation,
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

        // Helper to extract minimal driveItem fields
        const toDriveItemResult = (item: DriveItem): DriveItemResult => ({
          id: item.id!,
          name: item.name,
          webUrl: item.webUrl,
          downloadUrl: (item as any)["@microsoft.graph.downloadUrl"],
          size: item.size,
          mimeType: item.file?.mimeType,
        });

        // Convert HTML emails to PDF and upload to OneDrive
        if (message.body?.contentType?.includes("html")) {
          const pdfBytes = await convertHtmlToPdf(message);
          try {
            const pdfMeta: AttachmentResult = {
              name: `${message.subject || "email"}-${message.id}.pdf`,
              contentType: "application/pdf",
              size: pdfBytes.length,
            };
            const driveItem = await client.uploadFile(
              pdfMeta.name!,
              pdfBytes,
              `attachments/${sanitize(args.conversationId)}`
            );
            uploadedAttachments.push({
              attachment: pdfMeta,
              driveItem: toDriveItemResult(driveItem),
            });
          } catch (error) {
            console.error("Failed to upload PDF attachment:", error);
            uploadedAttachments.push({
              attachment: {},
              driveItem: { id: "" },
              error: "Failed to upload to OneDrive",
            });
          }
        }

        // Upload file attachments
        for (const attachment of message.attachments ?? []) {
          const fileAtt = attachment as any;
          if (
            fileAtt["@odata.type"] === "#microsoft.graph.fileAttachment" &&
            fileAtt.contentBytes
          ) {
            try {
              const driveItem = await client.uploadFile(
                attachment.name || "attachment",
                Buffer.from(fileAtt.contentBytes, "base64"),
                `attachments/${sanitize(args.conversationId)}`
              );
              uploadedAttachments.push({
                attachment: {
                  id: attachment.id,
                  name: attachment.name,
                  contentType: attachment.contentType,
                  size: attachment.size,
                },
                driveItem: toDriveItemResult(driveItem),
              });
            } catch (error) {
              uploadedAttachments.push({
                attachment: {
                  id: attachment.id,
                  name: attachment.name,
                  contentType: attachment.contentType,
                  size: attachment.size,
                },
                driveItem: { id: "" },
                error: "Failed to upload to OneDrive",
              });
            }
          }
        }

        return {
          id: message.id,
          subject: message.subject,
          from: message.from?.emailAddress?.address,
          receivedDateTime: message.receivedDateTime,
          bodyPreview: message.bodyPreview,
          uploadedAttachments,
        };
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
