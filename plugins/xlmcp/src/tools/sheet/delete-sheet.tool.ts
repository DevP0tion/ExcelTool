import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_delete_sheet",
    {
      title: "Delete Sheet",
      description: "Delete a sheet.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("Sheet name to delete"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, name }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Worksheets.Item('${psEscape(name)}').Delete()
      `);
      return textContent({ success: true });
    }
  );
}
