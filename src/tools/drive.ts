import { z } from "zod";
import { toFile } from "@anthropic-ai/sdk";
import { GraphClient } from "../graph/client.js";
import { getAnthropicClient } from "../anthropic/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";

export const getFile = {
  name: "get_file",
  schema: {
    title: "Get File",
    description:
      "Get a file from OneDrive by ID and upload it to Anthropic for analysis.",
    inputSchema: z.object({
      itemId: z.string().describe("The OneDrive item ID of the file"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const file = await client.getFileBytes(args.itemId);

    const anthropic = getAnthropicClient();
    const uploaded = await anthropic.beta.files.upload({
      file: await toFile(file.bytes, file.name, { type: file.mimeType }),
      betas: ["files-api-2025-04-14"],
    });

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            anthropicFileId: uploaded.id,
            name: file.name,
            type: file.mimeType,
            size: file.size,
            previewUrl: file.previewUrl,
          }),
        },
      ],
    };
  }) satisfies ToolCallback<{ itemId: z.ZodString }>,
};

export const searchFiles = {
  name: "search_files",
  schema: {
    title: "Search Files",
    description: "Search for files in OneDrive by name, content, or metadata.",
    inputSchema: z.object({
      query: z.string().describe("Search query"),
      top: z.number().optional().describe("Max results (default: 20)"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const items = await client.searchFiles(args.query, args.top);

    return {
      content: [{ type: "text", text: JSON.stringify(items, null, 2) }],
    };
  }) satisfies ToolCallback<{ query: z.ZodString; top: z.ZodOptional<z.ZodNumber> }>,
};

export default [getFile, searchFiles];
