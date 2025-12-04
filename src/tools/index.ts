// Import tools from new structure
import { getUserInfo } from "./user/getUserInfo.js";
import { getFile } from "./drive/getFile.js";
import { searchFiles } from "./drive/searchFiles.js";
import { moveFile } from "./drive/moveFile.js";
import { searchInboxMessages } from "./mail/searchInboxMessages.js";
import { getConversation } from "./mail/getConversation.js";

// Export individual tools
export {
  getUserInfo,
  getFile,
  searchFiles,
  moveFile,
  searchInboxMessages,
  getConversation,
};

// Re-export types for external consumers (e.g., mecha)
export type { GetUserInfoResult } from "./user/getUserInfo.js";
export type { GetFileResult } from "./drive/getFile.js";
export type { SearchFilesResult } from "./drive/searchFiles.js";
export type { MoveFileResult } from "./drive/moveFile.js";
export type { SearchInboxMessagesResult } from "./mail/searchInboxMessages.js";
export type {
  GetConversationResult,
  ConversationMessage,
  UploadedAttachment,
} from "./mail/getConversation.js";
