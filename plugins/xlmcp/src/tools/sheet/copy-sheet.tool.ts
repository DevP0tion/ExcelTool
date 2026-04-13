import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_sheet",
    {
      title: "Copy Sheet",
      description: "Copy sheet within the same workbook.",
      inputSchema: {
        workbook: workbookParam,
        source: z.string().describe("Source sheet name"),
        newName: z.string().optional().describe("Name for the copy. Auto-named if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, source, newName }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const renameCmd = newName
        ? `$wb.ActiveSheet.Name = '${psEscape(newName)}'`
        : "";
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $src = $wb.Worksheets.Item('${psEscape(source)}')
        $src.Copy([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        ${renameCmd}
        $wb.ActiveSheet.Name
      `);
      return textContent({ success: true, name: raw.trim() });
    }
  );
}
