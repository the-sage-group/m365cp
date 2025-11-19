import { z } from "zod";
import { createGraphClient } from "../graph/client.js";
import type { DriveItem } from "@microsoft/microsoft-graph-types";

export const getFileDownloadUrl = {
  name: "get_file_download_url",
  schema: {
    title: "Get File Download URL",
    description:
      "Get a file's metadata and temporary download URL from OneDrive by ID.",
    inputSchema: z.object({
      itemId: z.string().describe("The ID of the file to retrieve"),
    }),
  },
  handler: async (
    args: { itemId: string },
    extra: { authInfo?: { token?: string } }
  ) => {
    const accessToken = extra.authInfo?.token;
    if (!accessToken) {
      throw new Error("No access token available");
    }

    const client = createGraphClient(accessToken);

    // Get metadata including download URL
    const metadata: DriveItem & { "@microsoft.graph.downloadUrl"?: string } =
      await client.api(`/me/drive/items/${args.itemId}`).get();

    const downloadUrl = metadata["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) {
      throw new Error("No download URL available for this item");
    }

    console.log(`File ready: ${metadata.name}, size: ${metadata.size} bytes`);
    console.log(`Download URL: ${downloadUrl}`);

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify({
            id: metadata.id,
            filename: metadata.name,
            mimeType: metadata.file?.mimeType || "application/octet-stream",
            size: metadata.size,
            downloadUrl: downloadUrl,
            createdDateTime: metadata.createdDateTime,
            lastModifiedDateTime: metadata.lastModifiedDateTime,
          }),
        },
      ],
    };
  },
};

export const searchFiles = {
  name: "search_files",
  schema: {
    title: "Search Files",
    description:
      "Search for files and folders in OneDrive. Searches across filename, metadata, and file content.",
    inputSchema: z.object({
      query: z
        .string()
        .describe(
          "Search query text to find items by name, content, or metadata"
        ),
      top: z
        .number()
        .optional()
        .describe("Maximum number of results to return (default: 20)"),
    }),
  },
  handler: async (
    args: { query: string; top?: number },
    extra: { authInfo?: { token?: string } }
  ) => {
    const accessToken = extra.authInfo?.token;
    if (!accessToken) {
      throw new Error("No access token available");
    }

    console.log("Search query: ", args.query);

    const client = createGraphClient(accessToken);
    const top = args.top || 20;

    const response = await client
      .api(`/me/drive/root/search(q='${args.query}')`)
      .top(top)
      .get();

    const items: DriveItem[] = response.value || [];

    return {
      content: [
        {
          type: "text" as const,
          text: JSON.stringify(items, null, 2),
        },
      ],
      structuredContent: {
        items,
        count: items.length,
        hasMore: !!response["@odata.nextLink"],
      },
    };
  },
};
