import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_auto_filter",
    {
      title: "Auto Filter",
      description: "Set or remove auto filter. Apply criteria to a column.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Filter range (e.g. A1:D100)"),
        toggle: z.boolean().default(false).describe("Toggle filter on/off"),
        field: z.number().int().optional().describe("Column number to filter (1-based within range)"),
        criteria: z.string().optional().describe("Filter criteria (e.g. '>100')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, toggle, field, criteria }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (toggle) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
          else { $ws.Range('${psEscape(range)}').AutoFilter() }
        `);
      } else if (field && criteria) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').AutoFilter(${field}, '${psEscape(criteria)}')
        `);
      } else {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').AutoFilter()
        `);
      }
      return textContent({ success: true });
    }
  );
}
