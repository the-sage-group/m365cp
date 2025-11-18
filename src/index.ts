#!/usr/bin/env node

import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";

// Create MCP server (reused across requests)
const server = new McpServer({
  name: "mcp-microsoft365-graph",
  version: "1.0.0",
});

// Register hello_world tool
server.registerTool(
  "hello_world",
  {
    title: "Hello World",
    description: "A silly test tool that says hello",
    inputSchema: {
      name: z.string().describe("Your name"),
      emoji: z.string().optional().describe("Optional emoji to include"),
    },
    outputSchema: {
      greeting: z.string(),
    },
  },
  async ({ name, emoji }) => {
    const greeting = emoji
      ? `Hello, ${name}! ${emoji}${emoji}${emoji}`
      : `Hello, ${name}! Welcome to the Microsoft 365 Graph MCP Server!`;

    const output = { greeting };
    return {
      content: [
        {
          type: "text",
          text: greeting,
        },
      ],
      structuredContent: output,
    };
  }
);

// Create Express app
const app = express();
app.use(express.json());

// MCP endpoint (stateless pattern - recommended for most use cases)
app.post("/mcp", async (req, res) => {
  console.log("Received POST request to /mcp");

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
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: {
          code: -32603,
          message: "Internal server error",
        },
        id: null,
      });
    }
  }
});

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
