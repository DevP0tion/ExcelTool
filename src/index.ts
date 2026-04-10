#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { dispose } from "./services/powershell.js";
import { registerWorkbookTools } from "./tools/workbook/index.js";
import { registerSheetTools } from "./tools/sheet/index.js";
import { registerCellTools } from "./tools/cell/index.js";
import { registerFormatTools } from "./tools/format/index.js";

const server = new McpServer({
  name: "xlmcp",
  version: "0.1.0",
});

// 도구 등록
registerWorkbookTools(server);
registerSheetTools(server);
registerCellTools(server);
registerFormatTools(server);

// stdio transport
const transport = new StdioServerTransport();

process.on("SIGINT", async () => {
  await dispose();
  process.exit(0);
});

await server.connect(transport);
