import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import { toolNames } from "../names.js";

// ============================================================================
// Output Types
// ============================================================================

export type MoveFileResult = DriveItem;

// ============================================================================
// Tool Definition
// ============================================================================

export const moveFile = {
  name: toolNames.moveFile,
  schema: {
    title: "Move File",
    description:
      "Move a file to a different folder in OneDrive. Optionally rename the file during the move.",
    inputSchema: z.object({
      itemId: z
        .string()
        .describe(
          "The ID of the file to move (from search or other operations)"
        ),
      destinationFolderPath: z
        .string()
        .describe(
          "The destination folder path (e.g., 'Documents/Archive' or 'Projects/2025')"
        ),
      newName: z
        .string()
        .optional()
        .describe("Optional: New name for the file"),
    }),
  },
  handler: (async (args, extra) => {
    console.log(args);
    const client = new GraphClient(extra.authInfo!.token!);
    let movedItem;
    try {
      movedItem = await client.moveFile(
        args.itemId,
        args.destinationFolderPath,
        args.newName
      );
    } catch (error) {
      console.error("Error moving file:", error);
      throw error;
    }

    const result: MoveFileResult = movedItem;

    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  }) satisfies ToolCallback<{
    itemId: z.ZodString;
    destinationFolderPath: z.ZodString;
    newName: z.ZodOptional<z.ZodString>;
  }>,
};
