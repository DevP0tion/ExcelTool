import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_rename_sheet",
    {
      title: "Rename Sheet",
      description: "Rename a sheet.",
      inputSchema: {
        workbook: workbookParam,
        oldName: z.string().describe("Current sheet name"),
        newName: z.string().describe("New sheet name"),
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
