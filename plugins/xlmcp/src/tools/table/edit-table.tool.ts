import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_edit_table",
    {
      title: "표 편집",
      description: "기존 표의 이름, 스타일, 크기를 변경하거나 행/열을 추가합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("표 이름"),
        newName: z.string().optional().describe("새 표 이름"),
        style: z.string().optional().describe("표 스타일 (예: 'TableStyleLight1')"),
        resize: z.string().optional().describe("새 범위로 크기 변경 (예: 'A1:E20')"),
        addRow: z.boolean().optional().describe("true이면 맨 아래에 행 추가"),
        addColumn: z.string().optional().describe("추가할 열 이름"),
        showTotals: z.boolean().optional().describe("요약 행 표시/숨기기"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, name, newName, style, resize, addRow, addColumn, showTotals }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmds: string[] = [];
      if (newName) cmds.push(`$t.Name = '${psEscape(newName)}'`);
      if (style) cmds.push(`$t.TableStyle = '${psEscape(style)}'`);
      if (resize) cmds.push(`$t.Resize($ws.Range('${psEscape(resize)}'))`);
      if (addRow) cmds.push(`$t.ListRows.Add() | Out-Null`);
      if (addColumn) cmds.push(`$t.ListColumns.Add().Name = '${psEscape(addColumn)}'`);
      if (showTotals !== undefined) cmds.push(`$t.ShowTotals = $${showTotals}`);

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $t = $ws.ListObjects.Item('${psEscape(name)}')
        ${cmds.join("\n        ")}
      `);
      return textContent({ success: true });
    }
  );
}
