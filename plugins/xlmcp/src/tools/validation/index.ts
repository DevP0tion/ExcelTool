import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as setDataValidation } from "./set-data-validation.tool.js";
import { register as setConditionalFormat } from "./set-conditional-format.tool.js";

export function registerValidationTools(server: McpServer) {
  setDataValidation(server);
  setConditionalFormat(server);
}
