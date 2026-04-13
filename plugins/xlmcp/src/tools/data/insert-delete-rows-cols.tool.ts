import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_insert_delete_rows_cols",
    {
      title: "Insert/Delete Rows & Columns",
      description: "Insert or delete rows/columns.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        target: z.enum(["row", "column"]).describe("Target: row or column"),
        action: z.enum(["insert", "delete"]).describe("Action: insert or delete"),
        index: z.union([z.number().int(), z.string()]).describe("Row number (1-based) or column number/letter (1 or 'A')"),
        count: z.number().int().default(1).describe("Count to insert/delete"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, target, action, index, count }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      let rangeExpr: string;
      if (target === "row") {
        const idx = Number(index);
        rangeExpr = `$ws.Rows("${idx}:${idx + count - 1}")`;
      } else {
        if (typeof index === "number") {
          rangeExpr = `$ws.Columns("${index}:${index + count - 1}")`;
        } else {
          // 문자 입력: "A", "AA" 등 → 열 번호 변환 → 오프셋 → 다시 문자로
          const startCol = colLetterToNum(index);
          const endCol = startCol + count - 1;
          rangeExpr = `$ws.Columns("${index.toUpperCase()}:${colNumToLetter(endCol)}")`;
        }
      }
      const cmd = action === "insert" ? "Insert()" : "Delete()";
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        ${rangeExpr}.${cmd}
      `, { exclusive: true });
      return textContent({ success: true });
    }
  );
}

// "A"→1, "Z"→26, "AA"→27, "AZ"→52
function colLetterToNum(letter: string): number {
  let num = 0;
  for (const ch of letter.toUpperCase()) {
    num = num * 26 + (ch.charCodeAt(0) - 64);
  }
  return num;
}

// 1→"A", 26→"Z", 27→"AA", 52→"AZ"
function colNumToLetter(num: number): string {
  let result = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}
