import { Client } from "@microsoft/microsoft-graph-client";
import type { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import type {
  DriveItem,
  Message,
  User,
  Attachment,
  SearchHit,
} from "@microsoft/microsoft-graph-types";
import "isomorphic-fetch";
import sanitize from "sanitize-filename";

// ============================================================================
// Types
// ============================================================================

export interface FileContent {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  bytes: Buffer;
  base64: string;
  previewUrl?: string;
  downloadUrl?: string;
}

// ============================================================================
// Client
// ============================================================================

class AccessTokenAuthProvider implements AuthenticationProvider {
  constructor(private accessToken: string) {}
  async getAccessToken(): Promise<string> {
    return this.accessToken;
  }
}

export class GraphClient {
  private client: Client;

  constructor(accessToken: string) {
    if (!accessToken) {
      throw new Error("No access token provided");
    }
    this.client = Client.initWithMiddleware({
      authProvider: new AccessTokenAuthProvider(accessToken),
    });
  }

  /** Raw API access for operations not covered by helper methods */
  api(path: string) {
    return this.client.api(path);
  }

  // ==========================================================================
  // User
  // ==========================================================================

  async getUser(): Promise<User> {
    return this.client.api("/me").get();
  }

  // ==========================================================================
  // Drive
  // ==========================================================================

  async getFileBytes(itemId: string): Promise<FileContent> {
    const metadata: DriveItem & { "@microsoft.graph.downloadUrl"?: string } =
      await this.client.api(`/me/drive/items/${itemId}`).get();

    const downloadUrl = metadata["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) {
      throw new Error("No download URL available for this item");
    }

    const response = await fetch(downloadUrl);
    if (!response.ok) {
      throw new Error(`Failed to download file: ${response.statusText}`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const bytes = Buffer.from(arrayBuffer);

    return {
      id: itemId,
      name: metadata.name || "download",
      mimeType: metadata.file?.mimeType || "application/octet-stream",
      size: metadata.size || 0,
      bytes,
      previewUrl: metadata.webUrl || undefined,
      downloadUrl: downloadUrl,
      base64: bytes.toString("base64"),
    };
  }

  async searchFiles(query: string, top = 20): Promise<DriveItem[]> {
    const response = await this.client
      .api("/me/drive/root/search")
      .query({ q: query })
      .top(top)
      .get();
    return response.value || [];
  }

  async getFileMetadata(itemId: string): Promise<DriveItem> {
    return this.client.api(`/me/drive/items/${itemId}`).get();
  }

  async uploadFile(
    fileName: string,
    content: Buffer,
    folderPath?: string
  ): Promise<DriveItem> {
    const sanitizedFileName = sanitize(fileName);
    const uploadPath = folderPath
      ? `/me/drive/root:/${folderPath}/${sanitizedFileName}:/content`
      : `/me/drive/root:/${sanitizedFileName}:/content`;

    return this.client.api(uploadPath).put(content);
  }

  async moveFileById(
    itemId: string,
    destinationFolderId: string,
    newName?: string
  ): Promise<DriveItem> {
    const updatePayload: any = {
      parentReference: {
        id: destinationFolderId,
      },
    };

    if (newName) {
      updatePayload.name = sanitize(newName);
    }

    return this.client.api(`/me/drive/items/${itemId}`).patch(updatePayload);
  }

  // ==========================================================================
  // Mail
  // ==========================================================================

  async getMessage(messageId: string): Promise<Message> {
    return this.client.api(`/me/messages/${messageId}`).get();
  }

  async searchMessages(query: string, top = 25): Promise<SearchHit[]> {
    const response = await this.client.api("/search/query").post({
      requests: [
        {
          entityTypes: ["message"],
          query: {
            queryString: query,
          },
          from: 0,
          size: top,
        },
      ],
    });

    return response.value?.[0]?.hitsContainers?.[0]?.hits || [];
  }

  async getMessageAttachments(messageId: string): Promise<Attachment[]> {
    const response = await this.client
      .api(`/me/messages/${messageId}/attachments`)
      .get();
    return response.value || [];
  }

  async getConversationMessages(conversationId: string): Promise<Message[]> {
    try {
      const response = await this.client
        .api("/me/messages")
        .filter(
          `receivedDateTime ge 1900-01-01T00:00:00Z and conversationId eq '${conversationId}'`
        )
        .orderby("receivedDateTime asc")
        .expand("attachments")
        .get();

      return response.value || [];
    } catch (error) {
      console.error("Error getting conversation messages:", error);
      throw error;
    }
  }
}
