import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_cell",
    {
      title: "Write Cell",
      description: "Write value or formula to a cell. Starts with '=' for formula.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("Cell address (e.g. A1)"),
        value: z.string().describe("Value or formula (e.g. '=SUM(A1:A10)')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, cell, value }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const isFormula = value.startsWith("=");
      let cmd: string;
      if (isFormula) {
        cmd = `$c.Formula = '${psEscape(value)}'`;
      } else {
        const num = Number(value);
        cmd = value !== "" && !isNaN(num)
          ? `$c.Value2 = ${num}`
          : `$c.Value2 = '${psEscape(value)}'`;
      }
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
