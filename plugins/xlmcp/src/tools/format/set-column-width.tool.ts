import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_column_width",
    {
      title: "열 너비 설정",
      description: "열 너비를 설정합니다. 'auto'로 자동 맞춤 가능.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        columns: z.string().describe("열 범위 (예: 'A:C', 'B:B')"),
        width: z
          .union([z.number(), z.literal("auto")])
          .describe("너비 값(숫자) 또는 'auto'"),
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
