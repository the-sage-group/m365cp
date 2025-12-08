import { z } from "zod";
import createDebug from "debug";
import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { DriveItem } from "@microsoft/microsoft-graph-types";

const debug = createDebug("m365:search_files");

// ============================================================================
// Output Types
// ============================================================================

export type SearchFilesResult = DriveItem[];

// ============================================================================
// Tool Definition
// ============================================================================

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
    debug("request: %O", args);
    const client = new GraphClient(extra.authInfo!.token!);
    const items = await client.searchFiles(args.query, args.top);

    const result: SearchFilesResult = items;
    debug("response: %d files", result.length);

    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  }) satisfies ToolCallback<{
    query: z.ZodString;
    top: z.ZodOptional<z.ZodNumber>;
  }>,
};
