import { z } from "zod";
import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import { toolNames } from "../names.js";

// ============================================================================
// Output Types
// ============================================================================

export interface MoveFileResult {
  id: string;
  name?: string | null;
  webUrl?: string | null;
  downloadUrl?: string | null;
  size?: number | null;
  movedTo: {
    id: string;
    name?: string | null;
    path?: string | null;
  };
}

// ============================================================================
// Tool Definition
// ============================================================================

export const moveFile = {
  name: toolNames.moveFile,
  schema: {
    title: "Move File",
    description:
      "Move a file to a different folder in OneDrive. The destination is found via fuzzy search - describe the folder naturally (e.g., 'truman's mexico trip folder', 'project documents', 'Q4 reports').",
    inputSchema: z.object({
      itemId: z
        .string()
        .describe(
          "The ID of the file to move (from search or other operations)"
        ),
      destinationQuery: z
        .string()
        .describe(
          "A fuzzy description of the destination folder (e.g., 'truman mexico trip', 'project docs 2025', 'archived invoices')"
        ),
      newName: z
        .string()
        .optional()
        .describe("Optional: New name for the file"),
    }),
  },
  handler: (async (args, extra) => {
    const client = new GraphClient(extra.authInfo!.token!);

    // Search for folders matching the destination query
    const searchResults = await client.searchFiles(args.destinationQuery, 25);

    // Filter to only folders
    const folders = searchResults.filter((item) => item.folder);

    if (folders.length === 0) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                error: "No matching folders found",
                query: args.destinationQuery,
                suggestion:
                  "Try a different search query or check the folder exists",
              },
              null,
              2
            ),
          },
        ],
      };
    }

    // Use the best match (first result - Graph API returns relevance-ranked results)
    const destinationFolder = folders[0];

    // Move the file to the found folder
    const movedItem = await client.moveFileById(
      args.itemId,
      destinationFolder.id!,
      args.newName
    );

    const result: MoveFileResult = {
      id: movedItem.id!,
      name: movedItem.name,
      webUrl: movedItem.webUrl,
      downloadUrl: (movedItem as any)["@microsoft.graph.downloadUrl"],
      size: movedItem.size,
      movedTo: {
        id: destinationFolder.id!,
        name: destinationFolder.name,
        path: destinationFolder.parentReference?.path
          ? `${destinationFolder.parentReference.path}/${destinationFolder.name}`
          : destinationFolder.name,
      },
    };

    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  }) satisfies ToolCallback<{
    itemId: z.ZodString;
    destinationQuery: z.ZodString;
    newName: z.ZodOptional<z.ZodString>;
  }>,
};
