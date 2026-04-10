import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_delete_table",
    {
      title: "표 삭제",
      description: "표(ListObject)를 삭제합니다. 데이터는 유지하고 표 서식만 제거합니다. clearData를 true로 하면 데이터도 삭제합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("삭제할 표 이름"),
        clearData: z.boolean().default(false).describe("true이면 데이터도 함께 삭제"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, name, clearData }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmd = clearData
        ? `$t.Delete()`
        : `$t.Unlist()`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $t = $ws.ListObjects.Item('${psEscape(name)}')
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
