import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_column_width",
    {
      title: "Set Column Width",
      description: "Set column width. Use 'auto' for auto-fit.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        columns: z.string().describe("Column range (e.g. 'A:C', 'B:B')"),
        width: z
          .union([z.number(), z.literal("auto")])
          .describe("Width (number) or 'auto'"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, columns, width }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmd =
        width === "auto"
          ? `$ws.Range('${psEscape(columns)}').EntireColumn.AutoFit() | Out-Null`
          : `$ws.Range('${psEscape(columns)}').ColumnWidth = ${width}`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
