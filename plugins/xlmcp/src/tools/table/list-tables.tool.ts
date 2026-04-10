import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_tables",
    {
      title: "표 목록",
      description: "시트 내 모든 표(ListObject)의 이름, 범위, 스타일, 행/열 수를 반환합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $tables = @()
        foreach ($t in $ws.ListObjects) {
          $tables += @{
            Name = $t.Name
            Range = $t.Range.Address()
            Style = $t.TableStyle.Name
            Rows = $t.ListRows.Count
            Columns = $t.ListColumns.Count
          } | ConvertTo-Json -Compress
        }
        "[" + ($tables -join ",") + "]"
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
