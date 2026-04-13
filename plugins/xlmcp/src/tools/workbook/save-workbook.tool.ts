import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_save_workbook",
    {
      title: "Save Workbook",
      description: "Save workbook. Use savePath for Save As.",
      inputSchema: {
        workbook: workbookParam,
        savePath: z.string().optional().describe("Absolute path for Save As"),
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
