import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_set_data_validation",
    {
      title: "데이터 유효성 검사",
      description: "셀에 입력 규칙을 설정합니다. 드롭다운 목록, 숫자 범위, 텍스트 길이 등을 제한할 수 있습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("대상 범위 (예: B2:B100)"),
        type: z
          .enum(["list", "whole", "decimal", "textLength", "custom", "clear"])
          .describe("유효성 유형. clear는 기존 규칙 제거"),
        formula: z
          .string()
          .optional()
          .describe("list: 쉼표 구분 값 (예: 'A,B,C') 또는 범위참조. whole/decimal: 최소값. custom: 수식"),
        formula2: z.string().optional().describe("whole/decimal/textLength: 최대값"),
        errorMessage: z.string().optional().describe("오류 메시지"),
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
      // xlBetween=1 for range types, xlValidAlertStop=1
      const operator = (type === "list" || type === "custom") ? "" : ", 1"; // xlBetween
      const errCmd = errorMessage
        ? `$ws.Range('${psEscape(range)}').Validation.ErrorMessage = '${psEscape(errorMessage)}'`
        : "";

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        $r.Validation.Delete()
        $r.Validation.Add(${typeMap[type]}${operator}, [Type]::Missing, ${f1}${f2})
        ${errCmd}
      `);
      return textContent({ success: true });
    }
  );
}
