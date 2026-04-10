import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as readCell } from "./read-cell.tool.js";
import { register as writeCell } from "./write-cell.tool.js";
import { register as readRange } from "./read-range.tool.js";
import { register as writeRange } from "./write-range.tool.js";

export function registerCellTools(server: McpServer) {
  readCell(server);
  writeCell(server);
  readRange(server);
  writeRange(server);
}
