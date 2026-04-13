import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_read_cell",
    {
      title: "Read Cell",
      description: "Returns value, formula, and display text of a cell.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("Cell address (e.g. A1, B3)"),
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
      return textContent(parseJSON(raw));
    }
  );
}
