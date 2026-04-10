import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_range",
    {
      title: "범위 쓰기",
      description: "시작 셀부터 2D 배열 데이터를 입력합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        startCell: z.string().describe("시작 셀 주소 (예: A1)"),
        data: z
          .array(z.array(z.string()))
          .describe("2D 배열 데이터. 각 내부 배열이 한 행"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, startCell, data }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rows = data.length;
      const cols = data[0]?.length ?? 0;
      const psRows = data
        .map((row) => {
          const cells = row.map((v) => `'${psEscape(v)}'`).join(",");
          return `@(${cells})`;
        })
        .join(",");

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $start = $ws.Range('${psEscape(startCell)}')
        $endRow = $start.Row + ${rows} - 1
        $endCol = $start.Column + ${cols} - 1
        $endCell = $ws.Cells.Item($endRow, $endCol)
        $targetRange = $ws.Range($start, $endCell)
        $arr = New-Object 'object[,]' ${rows},${cols}
        $srcData = @(${psRows})
        for ($i = 0; $i -lt ${rows}; $i++) {
          $row = @($srcData[$i])
          for ($j = 0; $j -lt ${cols}; $j++) {
            $val = $row[$j]
            if ($val -match '^\=') {
              $ws.Cells.Item($start.Row + $i, $start.Column + $j).Formula = $val
            } else {
              $num = 0.0
              if ([double]::TryParse($val, [ref]$num)) {
                $arr[$i,$j] = $num
              } else {
                $arr[$i,$j] = $val
              }
            }
          }
        }
        $targetRange.Value2 = $arr
      `);
      return textContent({ success: true, rows, cols });
    }
  );
}
