import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_paste_range",
    {
      title: "범위 복사/붙여넣기",
      description: "범위를 복사하여 대상 위치에 붙여넣습니다. 값만, 수식만, 서식만 등 옵션 지정 가능.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("원본 범위 (예: A1:C10)"),
        destCell: z.string().describe("붙여넣기 시작 셀 (예: E1)"),
        destSheet: z.string().optional().describe("대상 시트. 생략 시 같은 시트"),
        pasteType: z
          .enum(["all", "values", "formulas", "formats"])
          .default("all")
          .describe("붙여넣기 유형"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, destSheet, pasteType }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const dstShName = destSheet ? `'${psEscape(destSheet)}'` : shName;
      // xlPasteAll=-4104, xlPasteValues=-4163, xlPasteFormulas=-4123, xlPasteFormats=-4122
      const pasteMap: Record<string, number> = { all: -4104, values: -4163, formulas: -4123, formats: -4122 };
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $srcWs = Resolve-Sheet $wb ${shName}
        $dstWs = Resolve-Sheet $wb ${dstShName}
        $srcWs.Range('${psEscape(sourceRange)}').Copy()
        $dstWs.Range('${psEscape(destCell)}').PasteSpecial(${pasteMap[pasteType]})
        $excel.CutCopyMode = $false
      `, { exclusive: true });
      return textContent({ success: true });
    }
  );
}
