// Import tools from new structure
import { getUserInfo } from "./user/getUserInfo.js";
import { getFile } from "./drive/getFile.js";
import { searchFiles } from "./drive/searchFiles.js";
import { searchInboxMessages } from "./mail/searchInboxMessages.js";
import { getConversation } from "./mail/getConversation.js";

// Export aggregated array as default
const allTools = [getUserInfo, getFile, searchFiles, searchInboxMessages, getConversation];

export default allTools;

// Re-export types for external consumers (e.g., mecha)
export type { GetUserInfoResult } from "./user/getUserInfo.js";
export type { GetFileResult } from "./drive/getFile.js";
export type { SearchFilesResult } from "./drive/searchFiles.js";
export type { SearchInboxMessagesResult } from "./mail/searchInboxMessages.js";
export type { GetConversationResult, ConversationMessage, UploadedAttachment } from "./mail/getConversation.js";
