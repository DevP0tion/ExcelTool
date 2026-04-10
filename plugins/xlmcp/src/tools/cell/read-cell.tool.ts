import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_read_cell",
    {
      title: "셀 읽기",
      description: "단일 셀의 값, 수식, 표시 텍스트를 반환합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("셀 주소 (예: A1, B3)"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, cell }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $c = $ws.Range('${psEscape(cell)}')
        @{
          Value = if ($c.Value2 -ne $null) { $c.Value2.ToString() } else { $null }
          Formula = $c.Formula
          Text = $c.Text
          NumberFormat = $c.NumberFormat
        } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
