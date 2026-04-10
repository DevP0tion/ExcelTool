import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_save_workbook",
    {
      title: "워크북 저장",
      description: "워크북을 저장합니다. savePath를 지정하면 다른 이름으로 저장합니다.",
      inputSchema: {
        workbook: workbookParam,
        savePath: z.string().optional().describe("다른 이름으로 저장할 절대 경로"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, savePath }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const saveCmd = savePath
        ? `$wb.SaveAs('${psEscape(savePath)}')`
        : `$wb.Save()`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        ${saveCmd}
      `);
      return textContent({ success: true });
    }
  );
}
