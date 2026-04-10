import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_row_height",
    {
      title: "행 높이 설정",
      description: "행 높이를 설정합니다. 'auto'로 자동 맞춤 가능.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        rows: z.string().describe("행 범위 (예: '1:5', '3:3')"),
        height: z
          .union([z.number(), z.literal("auto")])
          .describe("높이 값(숫자) 또는 'auto'"),
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
