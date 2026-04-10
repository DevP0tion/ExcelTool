import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as listTables } from "./list-tables.tool.js";
import { register as createTable } from "./create-table.tool.js";
import { register as editTable } from "./edit-table.tool.js";
import { register as deleteTable } from "./delete-table.tool.js";

export function registerTableTools(server: McpServer) {
  listTables(server);
  createTable(server);
  editTable(server);
  deleteTable(server);
}
