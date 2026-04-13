import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_open_workbook",
    {
      title: "Open Workbook",
      description: "Open workbook by path. Activates if already open.",
      inputSchema: {
        filePath: z.string().describe("Absolute path to Excel file (e.g. C:\\docs\\data.xlsx)"),
        readOnly: z.boolean().default(false).describe("Open as read-only"),
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
      return textContent(parseJSON(raw));
    }
  );
}
