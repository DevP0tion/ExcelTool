import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_get_active_workbook",
    {
      title: "활성 워크북 정보",
      description: "현재 활성화된 워크북의 이름, 경로, 시트 수, 활성 시트 이름을 반환합니다.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const raw = await runPS(`
        $wb = $excel.ActiveWorkbook
        if (-not $wb) { throw "열려 있는 워크북이 없습니다." }
        @{
          Name = $wb.Name
          Path = $wb.FullName
          SheetCount = $wb.Worksheets.Count
          ActiveSheet = $wb.ActiveSheet.Name
        } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
