import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getPoolStatus } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_pool_status",
    {
      title: "세션 풀 상태",
      description: `PowerShell 세션 풀의 현재 상태를 반환합니다.
세션 수, busy/alive 상태, 작업 큐 길이, 처리 통계를 확인할 수 있습니다.`,
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      return textContent(getPoolStatus());
    }
  );
}
