import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_sheets",
    {
      title: "시트 목록",
      description: "워크북의 모든 시트 이름을 반환합니다.",
      inputSchema: { workbook: workbookParam },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $names = @()
        foreach ($ws in $wb.Worksheets) { $names += $ws.Name }
        ConvertTo-Json @($names) -Compress
      `);
      return textContent({ sheets: JSON.parse(raw.trim()) });
    }
  );
}
