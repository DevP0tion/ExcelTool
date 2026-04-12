import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as insertVba } from "./insert-vba.tool.js";
import { register as listVba } from "./list-vba.tool.js";
import { register as manageVba } from "./manage-vba.tool.js";

export function registerVbaTools(server: McpServer) {
  insertVba(server);
  listVba(server);
  manageVba(server);
}
