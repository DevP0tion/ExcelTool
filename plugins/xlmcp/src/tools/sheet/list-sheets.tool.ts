import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_sheets",
    {
      title: "List Sheets",
      description: "Returns all sheet names.",
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
      return textContent({ sheets: parseJSON(raw) });
    }
  );
}
