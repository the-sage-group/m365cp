import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Message, ItemBody } from "@microsoft/microsoft-graph-types";
import { toolNames } from "../names.js";

// ============================================================================
// Input Schemas
// ============================================================================

const recipientSchema = z.object({
  address: z.string().email().describe("Email address of the recipient"),
  name: z.string().optional().describe("Display name of the recipient"),
});

// ============================================================================
// Output Types
// ============================================================================

export interface CreateDraftEmailResult {
  id: string;
  webLink?: string | null;
  subject?: string | null;
  isDraft: boolean;
  hasAttachments: boolean;
  attachmentCount: number;
}

// ============================================================================
// Tool Definition
// ============================================================================

export const createDraftEmail = {
  name: toolNames.createDraftEmail,
  schema: {
    title: "Create Draft Email",
    description:
      "Creates a draft email message in the user's Drafts folder. " +
      "Supports HTML content and file attachments from OneDrive (referenced by drive item ID). " +
      "The draft can be edited and sent later from Outlook.",
    inputSchema: z.object({
      subject: z.string().describe("Subject line of the email"),
      body: z.string().describe("Body content of the email (HTML or plain text)"),
      bodyContentType: z
        .enum(["html", "text"])
        .optional()
        .default("html")
        .describe('Content type of the body: "html" or "text" (default: "html")'),
      toRecipients: z
        .array(recipientSchema)
        .optional()
        .describe("Primary recipients of the email"),
      ccRecipients: z
        .array(recipientSchema)
        .optional()
        .describe("CC recipients of the email"),
      bccRecipients: z
        .array(recipientSchema)
        .optional()
        .describe("BCC recipients of the email"),
      importance: z
        .enum(["low", "normal", "high"])
        .optional()
        .default("normal")
        .describe('Importance level: "low", "normal", or "high"'),
      attachmentDriveItemIds: z
        .array(z.string())
        .optional()
        .describe(
          "OneDrive item IDs of files to attach to the email. " +
            "Files will be fetched from OneDrive and attached (max 3MB each)."
        ),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);

    // Build recipients
    const formatRecipients = (
      recipients?: Array<{ address: string; name?: string }>
    ) =>
      recipients?.map((r) => ({
        emailAddress: {
          address: r.address,
          name: r.name,
        },
      }));

    // Build the message object
    const message: Message = {
      subject: args.subject,
      body: {
        contentType: args.bodyContentType,
        content: args.body,
      } as ItemBody,
      toRecipients: formatRecipients(args.toRecipients),
      ccRecipients: formatRecipients(args.ccRecipients),
      bccRecipients: formatRecipients(args.bccRecipients),
      importance: args.importance,
    };

    // Create the draft with attachments
    const createdMessage = await client.createDraftEmail(
      message,
      args.attachmentDriveItemIds
    );

    const result: CreateDraftEmailResult = {
      id: createdMessage.id!,
      webLink: createdMessage.webLink,
      subject: createdMessage.subject,
      isDraft: createdMessage.isDraft ?? true,
      hasAttachments: createdMessage.hasAttachments ?? false,
      attachmentCount: createdMessage.attachments?.length ?? 0,
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
    subject: z.ZodString;
    body: z.ZodString;
    bodyContentType: z.ZodOptional<z.ZodDefault<z.ZodEnum<["html", "text"]>>>;
    toRecipients: z.ZodOptional<
      z.ZodArray<
        z.ZodObject<{
          address: z.ZodString;
          name: z.ZodOptional<z.ZodString>;
        }>
      >
    >;
    ccRecipients: z.ZodOptional<
      z.ZodArray<
        z.ZodObject<{
          address: z.ZodString;
          name: z.ZodOptional<z.ZodString>;
        }>
      >
    >;
    bccRecipients: z.ZodOptional<
      z.ZodArray<
        z.ZodObject<{
          address: z.ZodString;
          name: z.ZodOptional<z.ZodString>;
        }>
      >
    >;
    importance: z.ZodOptional<
      z.ZodDefault<z.ZodEnum<["low", "normal", "high"]>>
    >;
    attachmentDriveItemIds: z.ZodOptional<z.ZodArray<z.ZodString>>;
  }>,
};
