import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { textContent, parseJSON } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_get_active_workbook",
    {
      title: "Active Workbook Info",
      description: "Returns name, path, sheet count, and active sheet of the current workbook.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const raw = await runPS(`
        $wb = $excel.ActiveWorkbook
        if (-not $wb) { throw "No workbook is open." }
        $vbaTrusted = $false
        try { $null = $wb.VBProject.VBComponents.Count; $vbaTrusted = $true } catch {}
        @{
          Name = $wb.Name
          Path = $wb.FullName
          SheetCount = $wb.Worksheets.Count
          ActiveSheet = $wb.ActiveSheet.Name
          VbaAccessTrusted = $vbaTrusted
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
