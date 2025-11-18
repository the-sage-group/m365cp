import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerUserTools } from "./user.js";

/**
 * Register all Microsoft Graph tools with the MCP server
 */
export function registerAllTools(server: McpServer) {
  registerUserTools(server);
  // Future tool categories can be added here:
  // registerMailTools(server);
  // registerCalendarTools(server);
  // registerFileTools(server);
}
