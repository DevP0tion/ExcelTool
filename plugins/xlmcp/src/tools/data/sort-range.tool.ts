import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_sort_range",
    {
      title: "Sort Range",
      description: "Sort range by a column.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range to sort (e.g. A1:D100)"),
        sortBy: z.string().describe("Sort key column (e.g. B1)"),
        order: z.enum(["asc", "desc"]).default("asc").describe("Sort order"),
        hasHeader: z.boolean().default(true).describe("First row is header"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, sortBy, order, hasHeader }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      // xlAscending=1, xlDescending=2, xlYes=1, xlNo=2
      const xlOrder = order === "asc" ? 1 : 2;
      const xlHeader = hasHeader ? 1 : 2;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        $ws.Sort.SortFields.Clear()
        $ws.Sort.SortFields.Add($ws.Range('${psEscape(sortBy)}'), [Type]::Missing, ${xlOrder})
        $ws.Sort.SetRange($r)
        $ws.Sort.Header = ${xlHeader}
        $ws.Sort.Apply()
      `);
      return textContent({ success: true });
    }
  );
}
