#!/usr/bin/env node

import "dotenv/config";
import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { requireBearerAuth } from "@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js";
import { MicrosoftGraphTokenVerifier } from "./auth/verifier.js";
import { ZodRawShape } from "zod";
import { ToolCallback } from "@modelcontextprotocol/sdk/server/mcp.js";

import tools from "./tools/index.js";

// Create MCP server (reused across requests)
const server = new McpServer({
  name: "mcp-microsoft365-graph",
  version: "1.0.0",
});

// Register all Graph API tools
for (const tool of tools) {
  server.registerTool(
    tool.name,
    tool.schema,
    tool.handler as ToolCallback<ZodRawShape>
  );
}

// Create Express app
const app = express();
app.use(express.json());

// Create token verifier
const tokenVerifier = new MicrosoftGraphTokenVerifier();

// MCP endpoint (stateless pattern - recommended for most use cases)
app.post(
  "/mcp",
  requireBearerAuth({ verifier: tokenVerifier }),
  async (req, res) => {
    try {
      // Create a new transport for each request to prevent request ID collisions
      // Different clients may use the same JSON-RPC request IDs
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
        enableJsonResponse: true,
      });

      res.on("close", () => {
        transport.close();
      });

      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("Error handling MCP request:", error);
    }
  }
);

// Health check endpoint
app.get("/health", async (req, res) => {
  res.json({
    status: "healthy",
    name: "mcp-microsoft365-graph",
    version: "1.0.0",
  });
});

// Start server
async function main() {
  const PORT = process.env.PORT || 3001;

  return new Promise<void>((resolve) => {
    const httpServer = app.listen(PORT, () => {
      console.log(`Microsoft 365 Graph MCP Server listening on port ${PORT}`);
      console.log(`MCP endpoint: http://localhost:${PORT}/mcp`);
      console.log(`Health check: http://localhost:${PORT}/health`);
      resolve();
    });

    // Handle graceful shutdown
    const shutdown = () => {
      console.log("Shutting down server...");
      httpServer.close(() => {
        console.log("HTTP server closed");
        process.exit(0);
      });
    };

    process.on("SIGTERM", shutdown);
    process.on("SIGINT", shutdown);
  });
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
