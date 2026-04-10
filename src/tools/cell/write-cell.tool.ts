import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_cell",
    {
      title: "셀 쓰기",
      description: "단일 셀에 값 또는 수식을 입력합니다. '='로 시작하면 수식으로 처리됩니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("셀 주소 (예: A1)"),
        value: z.string().describe("입력할 값 또는 수식 (예: '=SUM(A1:A10)')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, cell, value }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const isFormula = value.startsWith("=");
      const cmd = isFormula
        ? `$c.Formula = '${psEscape(value)}'`
        : `$c.Value2 = '${psEscape(value)}'`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $c = $ws.Range('${psEscape(cell)}')
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
