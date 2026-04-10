import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_freeze_panes",
    {
      title: "틀 고정",
      description: "셀 위치 기준으로 틀을 고정하거나 해제합니다. cell을 지정하면 해당 셀의 위쪽 행과 왼쪽 열이 고정됩니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().optional().describe("고정 기준 셀 (예: B3 → 1~2행, A열 고정). 생략 시 해제"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, cell }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      if (cell) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Activate()
          $ws.Range('${psEscape(cell)}').Select()
          $excel.ActiveWindow.FreezePanes = $false
          $excel.ActiveWindow.FreezePanes = $true
        `);
      } else {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Activate()
          $excel.ActiveWindow.FreezePanes = $false
        `);
      }
      return textContent({ success: true });
    }
  );
}
