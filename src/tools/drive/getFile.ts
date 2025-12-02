import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { DriveItem } from "@microsoft/microsoft-graph-types";

// ============================================================================
// Output Types
// ============================================================================

export type GetFileResult = DriveItem;

// ============================================================================
// Tool Definition
// ============================================================================

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
    const driveItem = await client.getFileMetadata(args.itemId);

    const result: GetFileResult = driveItem;

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(result),
        },
      ],
    };
  }) satisfies ToolCallback<{ itemId: z.ZodString }>,
};
