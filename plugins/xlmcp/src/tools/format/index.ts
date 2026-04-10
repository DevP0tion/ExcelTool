import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as formatRange } from "./format-range.tool.js";
import { register as setColumnWidth } from "./set-column-width.tool.js";
import { register as setRowHeight } from "./set-row-height.tool.js";
import { register as mergeCells } from "./merge-cells.tool.js";

export function registerFormatTools(server: McpServer) {
  formatRange(server);
  setColumnWidth(server);
  setRowHeight(server);
  mergeCells(server);
}
