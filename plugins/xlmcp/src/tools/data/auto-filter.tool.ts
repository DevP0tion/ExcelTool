import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_auto_filter",
    {
      title: "자동 필터",
      description: "범위에 자동 필터를 설정하거나 해제합니다. criteria를 지정하면 특정 열에 필터 조건을 적용합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("필터 범위 (예: A1:D100)"),
        toggle: z.boolean().default(false).describe("true이면 필터 토글 (있으면 해제, 없으면 설정)"),
        field: z.number().int().optional().describe("필터 적용할 열 번호 (범위 내 1부터)"),
        criteria: z.string().optional().describe("필터 조건 (예: '>100', 'A사')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, toggle, field, criteria }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (toggle) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
          else { $ws.Range('${psEscape(range)}').AutoFilter() }
        `);
      } else if (field && criteria) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').AutoFilter(${field}, '${psEscape(criteria)}')
        `);
      } else {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Range('${psEscape(range)}').AutoFilter()
        `);
      }
      return textContent({ success: true });
    }
  );
}
