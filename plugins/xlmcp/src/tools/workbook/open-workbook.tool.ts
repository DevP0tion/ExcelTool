import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_open_workbook",
    {
      title: "워크북 열기",
      description: "파일 경로로 워크북을 엽니다. 이미 열려 있으면 해당 워크북을 활성화합니다.",
      inputSchema: {
        filePath: z.string().describe("Excel 파일 절대 경로 (예: C:\\docs\\data.xlsx)"),
        readOnly: z.boolean().default(false).describe("읽기 전용으로 열기"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ filePath, readOnly }) => {
      const raw = await runPS(`
        $path = '${psEscape(filePath)}'
        $existing = $null
        foreach ($wb in $excel.Workbooks) {
          if ($wb.FullName -eq $path) { $existing = $wb; break }
        }
        if ($existing) {
          $existing.Activate()
          $wb = $existing
        } else {
          $wb = $excel.Workbooks.Open($path, [System.Reflection.Missing]::Value, $${readOnly})
        }
        @{
          Name = $wb.Name
          Path = $wb.FullName
          SheetCount = $wb.Worksheets.Count
          ReadOnly = [bool]$wb.ReadOnly
        } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
