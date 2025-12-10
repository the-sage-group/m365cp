import { z } from "zod";
import createDebug from "debug";
import { GraphClient } from "../../graph/client.js";
import type { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import { toolNames } from "../names.js";

const debug = createDebug("m365:search_files");

// ============================================================================
// Output Types
// ============================================================================

export interface FileResult {
  id: string;
  name?: string | null;
  webUrl?: string | null;
  downloadUrl?: string | null;
  size?: number | null;
  mimeType?: string | null;
}

export type SearchFilesResult = FileResult[];

// ============================================================================
// Tool Definition
// ============================================================================

export const searchFiles = {
  name: toolNames.searchFiles,
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

    const result: SearchFilesResult = items.map((item) => ({
      id: item.id!,
      name: item.name,
      webUrl: item.webUrl,
      downloadUrl: (item as any)["@microsoft.graph.downloadUrl"],
      size: item.size,
      mimeType: item.file?.mimeType,
    }));
    debug("response: %d files", result.length);

    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  }) satisfies ToolCallback<{
    query: z.ZodString;
    top: z.ZodOptional<z.ZodNumber>;
  }>,
};
