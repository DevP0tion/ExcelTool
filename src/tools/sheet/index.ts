import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as listSheets } from "./list-sheets.tool.js";
import { register as createSheet } from "./create-sheet.tool.js";
import { register as deleteSheet } from "./delete-sheet.tool.js";
import { register as copySheet } from "./copy-sheet.tool.js";
import { register as renameSheet } from "./rename-sheet.tool.js";

export function registerSheetTools(server: McpServer) {
  listSheets(server);
  createSheet(server);
  deleteSheet(server);
  copySheet(server);
  renameSheet(server);
}
