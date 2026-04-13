import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_sheet",
    {
      title: "Add Sheet",
      description: "Add a new sheet.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("New sheet name"),
        after: z.string().optional().describe("Add after this sheet. Appends to end if omitted"),
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
