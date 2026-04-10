import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_sheet",
    {
      title: "시트 추가",
      description: "새 시트를 추가합니다.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("새 시트 이름"),
        after: z.string().optional().describe("이 시트 뒤에 추가. 생략 시 맨 뒤"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, name, after }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const afterCmd = after
        ? `$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item('${psEscape(after)}'))`
        : `$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        ${afterCmd}
        $ws.Name = '${psEscape(name)}'
      `);
      return textContent({ success: true, name });
    }
  );
}
