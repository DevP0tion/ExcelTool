import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_close_workbook",
    {
      title: "워크북 닫기",
      description: "워크북을 닫습니다. save 옵션으로 저장 여부를 지정합니다.",
      inputSchema: {
        workbook: workbookParam,
        save: z.boolean().default(false).describe("닫기 전 저장 여부"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, save }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Close(${save ? "$true" : "$false"})
      `);
      return textContent({ success: true });
    }
  );
}
