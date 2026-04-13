import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_row_height",
    {
      title: "Set Row Height",
      description: "Set row height. Use 'auto' for auto-fit.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        rows: z.string().describe("Row range (e.g. '1:5', '3:3')"),
        height: z
          .union([z.number(), z.literal("auto")])
          .describe("Height (number) or 'auto'"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, rows, height }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmd =
        height === "auto"
          ? `$ws.Range('${psEscape(rows)}').EntireRow.AutoFit() | Out-Null`
          : `$ws.Range('${psEscape(rows)}').RowHeight = ${height}`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
