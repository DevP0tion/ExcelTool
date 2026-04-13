import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_merge_cells",
    {
      title: "Merge/Unmerge Cells",
      description: "Merge or unmerge a cell range.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range address (e.g. A1:D1)"),
        unmerge: z.boolean().default(false).describe("Set true to unmerge"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, unmerge }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmd = unmerge ? `$r.UnMerge()` : `$r.Merge()`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
