import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_conditional_format",
    {
      title: "조건부 서식",
      description: "범위에 조건부 서식 규칙을 추가합니다. clear 타입으로 기존 규칙을 제거할 수 있습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("대상 범위 (예: B2:B100)"),
        type: z
          .enum(["cellValue", "expression", "colorScale", "dataBar", "clear"])
          .describe("규칙 유형. clear는 기존 규칙 모두 제거"),
        operator: z
          .enum(["greaterThan", "lessThan", "equal", "between", "notEqual"])
          .optional()
          .describe("cellValue 유형의 비교 연산자"),
        formula: z.string().optional().describe("조건 값 또는 수식 (예: '100', '=TODAY()')"),
        formula2: z.string().optional().describe("between 연산자의 두 번째 값"),
        fontColor: z.string().optional().describe("조건 충족 시 폰트 색상 RGB hex"),
        bgColor: z.string().optional().describe("조건 충족 시 배경 색상 RGB hex"),
        bold: z.boolean().optional().describe("조건 충족 시 굵게"),
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
        const [r, g, b_] = hexToRgb(fontColor);
        fmtCmds.push(`$fc.Font.Color = ${r + g * 256 + b_ * 65536}`);
      }
      if (bgColor) {
        const [r, g, b_] = hexToRgb(bgColor);
        fmtCmds.push(`$fc.Interior.Color = ${r + g * 256 + b_ * 65536}`);
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

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [parseInt(h.substring(0, 2), 16), parseInt(h.substring(2, 4), 16), parseInt(h.substring(4, 6), 16)];
}
