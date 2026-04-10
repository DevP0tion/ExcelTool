import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as listOpenWorkbooks } from "./list-open-workbooks.tool.js";
import { register as getActiveWorkbook } from "./get-active-workbook.tool.js";
import { register as createWorkbook } from "./create-workbook.tool.js";
import { register as saveWorkbook } from "./save-workbook.tool.js";
import { register as closeWorkbook } from "./close-workbook.tool.js";

export function registerWorkbookTools(server: McpServer) {
  listOpenWorkbooks(server);
  getActiveWorkbook(server);
  createWorkbook(server);
  saveWorkbook(server);
  closeWorkbook(server);
}
