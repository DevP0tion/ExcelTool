#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { dispose } from "./services/powershell.js";
import { registerWorkbookTools } from "./tools/workbook/index.js";
import { registerSheetTools } from "./tools/sheet/index.js";
import { registerCellTools } from "./tools/cell/index.js";
import { registerFormatTools } from "./tools/format/index.js";
import { registerDataTools } from "./tools/data/index.js";
import { registerTableTools } from "./tools/table/index.js";
import { registerChartTools } from "./tools/chart/index.js";
import { registerPivotTools } from "./tools/pivot/index.js";
import { registerValidationTools } from "./tools/validation/index.js";
import { registerViewTools } from "./tools/view/index.js";

const server = new McpServer({
  name: "xlmcp",
  version: "0.2.0",
});

// 도구 등록
registerWorkbookTools(server);
registerSheetTools(server);
registerCellTools(server);
registerFormatTools(server);
registerDataTools(server);
registerTableTools(server);
registerChartTools(server);
registerPivotTools(server);
registerValidationTools(server);
registerViewTools(server);

// stdio transport
const transport = new StdioServerTransport();

process.on("SIGINT", async () => {
  await dispose();
  process.exit(0);
});

await server.connect(transport);
