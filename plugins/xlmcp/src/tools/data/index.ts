import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as insertDeleteRowsCols } from "./insert-delete-rows-cols.tool.js";
import { register as copyPasteRange } from "./copy-paste-range.tool.js";
import { register as findReplace } from "./find-replace.tool.js";
import { register as sortRange } from "./sort-range.tool.js";
import { register as autoFilter } from "./auto-filter.tool.js";

export function registerDataTools(server: McpServer) {
  insertDeleteRowsCols(server);
  copyPasteRange(server);
  findReplace(server);
  sortRange(server);
  autoFilter(server);
}
