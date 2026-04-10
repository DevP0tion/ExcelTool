import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_rename_sheet",
    {
      title: "시트 이름 변경",
      description: "시트 이름을 변경합니다.",
      inputSchema: {
        workbook: workbookParam,
        oldName: z.string().describe("현재 시트 이름"),
        newName: z.string().describe("새 시트 이름"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, oldName, newName }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Worksheets.Item('${psEscape(oldName)}').Name = '${psEscape(newName)}'
      `);
      return textContent({ success: true });
    }
  );
}
