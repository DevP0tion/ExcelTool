import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getPoolStatus } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_pool_status",
    {
      title: "Session Pool Status",
      description: `Returns PowerShell session pool status: sessions, queue, stats.`,
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      return textContent(getPoolStatus());
    }
  );
}
