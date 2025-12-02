import { z } from "zod";
import { GraphClient } from "../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";

export const getFile = {
  name: "get_file",
  schema: {
    title: "Get File",
    description:
      "Get a file from OneDrive by ID. Returns file metadata including OneDrive file ID, name, size, and URLs.",
    inputSchema: z.object({
      itemId: z.string().describe("The OneDrive item ID of the file"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);
    const file = await client.getFileBytes(args.itemId);

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            fileId: args.itemId,
            fileName: file.name,
            mimeType: file.mimeType,
            size: file.size.toString(),
            downloadUrl: file.downloadUrl,
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
  }) satisfies ToolCallback<{
    query: z.ZodString;
    top: z.ZodOptional<z.ZodNumber>;
  }>,
};

export default [getFile, searchFiles];
