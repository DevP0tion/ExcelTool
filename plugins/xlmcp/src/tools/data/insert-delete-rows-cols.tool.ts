import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_insert_delete_rows_cols",
    {
      title: "행/열 삽입·삭제",
      description: "행 또는 열을 삽입하거나 삭제합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        target: z.enum(["row", "column"]).describe("대상: row 또는 column"),
        action: z.enum(["insert", "delete"]).describe("동작: insert 또는 delete"),
        index: z.number().int().describe("행 번호(1부터) 또는 열 번호(1부터)"),
        count: z.number().int().default(1).describe("삽입/삭제할 개수"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, target, action, index, count }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = target === "row"
        ? `$ws.Rows("${index}:${index + count - 1}")`
        : `$ws.Columns("${index}:${index + count - 1}")`;
      const cmd = action === "insert" ? "Insert()" : "Delete()";
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        ${rangeExpr}.${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
