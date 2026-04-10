import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as freezePanes } from "./freeze-panes.tool.js";
import { register as namedRange } from "./named-range.tool.js";

export function registerViewTools(server: McpServer) {
  freezePanes(server);
  namedRange(server);
}
