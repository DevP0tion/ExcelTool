import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as createPivotTable } from "./create-pivot-table.tool.js";

export function registerPivotTools(server: McpServer) {
  createPivotTable(server);
}
