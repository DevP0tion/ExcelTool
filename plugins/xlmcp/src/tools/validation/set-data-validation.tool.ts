import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_data_validation",
    {
      title: "Data Validation",
      description: "Set cell input rules: dropdown list, number range, text length, etc.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Target range (e.g. B2:B100)"),
        type: z
          .enum(["list", "whole", "decimal", "textLength", "custom", "clear"])
          .describe("Validation type. 'clear' removes rules"),
        formula: z
          .string()
          .optional()
          .describe("list: comma-separated or range ref. whole/decimal: min. custom: formula"),
        formula2: z.string().optional().describe("whole/decimal/textLength: max value"),
        errorMessage: z.string().optional().describe("Error message"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, type, formula, formula2, errorMessage }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (type === "clear") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').Validation.Delete()
        `);
        return textContent({ success: true });
      }

      // xlValidateList=3, xlValidateWholeNumber=1, xlValidateDecimal=2, xlValidateTextLength=6, xlValidateCustom=7
      const typeMap: Record<string, number> = { list: 3, whole: 1, decimal: 2, textLength: 6, custom: 7 };
      const f1 = formula ? `'${psEscape(formula)}'` : "''";
      const f2 = formula2 ? `,'${psEscape(formula2)}'` : "";
      // Validation.Add(Type, AlertStyle, Operator, Formula1, Formula2)
      // xlBetween=1, AlertStyle는 Missing 처리
      const operator = (type === "list" || type === "custom") ? "[Type]::Missing" : "1"; // xlBetween
      const errCmd = errorMessage
        ? `$ws.Range('${psEscape(range)}').Validation.ErrorMessage = '${psEscape(errorMessage)}'`
        : "";

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        $r.Validation.Delete()
        $r.Validation.Add(${typeMap[type]}, [Type]::Missing, ${operator}, ${f1}${f2})
        ${errCmd}
      `);
      return textContent({ success: true });
    }
  );
}
