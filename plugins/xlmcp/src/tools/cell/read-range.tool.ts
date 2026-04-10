import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_read_range",
    {
      title: "범위 읽기",
      description: "셀 범위의 값을 2D 배열로 반환합니다. 범위를 생략하면 UsedRange 전체를 읽습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().optional().describe("범위 주소 (예: A1:C10). 생략 시 UsedRange"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range
        ? `$ws.Range('${psEscape(range)}')`
        : `$ws.UsedRange`;
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = ${rangeExpr}
        $rows = $r.Rows.Count
        $cols = $r.Columns.Count
        $data = @()
        for ($i = 1; $i -le $rows; $i++) {
          $row = @()
          for ($j = 1; $j -le $cols; $j++) {
            $v = $r.Cells.Item($i, $j).Value2
            $row += if ($v -ne $null) { $v.ToString() } else { "" }
          }
          $data += ,@($row)
        }
        @{
          Range = $r.Address()
          Rows = $rows
          Cols = $cols
          Data = $data
        } | ConvertTo-Json -Depth 10 -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
