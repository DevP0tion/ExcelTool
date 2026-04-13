import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, hexToRgb, rgbToOle } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_conditional_format",
    {
      title: "Conditional Formatting",
      description: "Add conditional format rule. Use 'clear' to remove all.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Target range (e.g. B2:B100)"),
        type: z
          .enum(["cellValue", "expression", "colorScale", "dataBar", "clear"])
          .describe("Rule type. 'clear' removes all rules"),
        operator: z
          .enum(["greaterThan", "lessThan", "equal", "between", "notEqual"])
          .optional()
          .describe("Comparison operator for cellValue type"),
        formula: z.string().optional().describe("Condition value or formula (e.g. '100', '=TODAY()')"),
        formula2: z.string().optional().describe("Second value for 'between' operator"),
        fontColor: z.string().optional().describe("Font color RGB hex when met"),
        bgColor: z.string().optional().describe("Background color RGB hex when met"),
        bold: z.boolean().optional().describe("Bold when met"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, type, operator, formula, formula2, fontColor, bgColor, bold }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (type === "clear") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').FormatConditions.Delete()
        `);
        return textContent({ success: true });
      }

      if (type === "colorScale") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').FormatConditions.AddColorScale(3) | Out-Null
        `);
        return textContent({ success: true });
      }

      if (type === "dataBar") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').FormatConditions.AddDatabar() | Out-Null
        `);
        return textContent({ success: true });
      }

      // cellValue or expression
      // xlCellValue=1, xlExpression=2
      const xlType = type === "cellValue" ? 1 : 2;
      // xlGreater=5, xlLess=6, xlEqual=3, xlBetween=1, xlNotEqual=4
      const opMap: Record<string, number> = { greaterThan: 5, lessThan: 6, equal: 3, between: 1, notEqual: 4 };
      const xlOp = operator ? opMap[operator] : 3;
      const f1 = formula ? `'${psEscape(formula)}'` : "''";
      const f2 = formula2 ? `,'${psEscape(formula2)}'` : "";

      const fmtCmds: string[] = [];
      if (fontColor) {
        fmtCmds.push(`$fc.Font.Color = ${rgbToOle(hexToRgb(fontColor))}`);
      }
      if (bgColor) {
        fmtCmds.push(`$fc.Interior.Color = ${rgbToOle(hexToRgb(bgColor))}`);
      }
      if (bold !== undefined) fmtCmds.push(`$fc.Font.Bold = $${bold}`);

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        $fc = $r.FormatConditions.Add(${xlType}, ${xlOp}, ${f1}${f2})
        ${fmtCmds.join("\n        ")}
      `);
      return textContent({ success: true });
    }
  );
}
