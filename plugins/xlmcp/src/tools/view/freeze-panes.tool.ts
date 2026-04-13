import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_freeze_panes",
    {
      title: "Freeze Panes",
      description: "Freeze/unfreeze panes at cell position.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().optional().describe("Freeze cell (e.g. B3). Omit to unfreeze"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, cell }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      if (cell) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Activate()
          $ws.Range('${psEscape(cell)}').Select()
          $excel.ActiveWindow.FreezePanes = $false
          $excel.ActiveWindow.FreezePanes = $true
        `, { exclusive: true });
      } else {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Activate()
          $excel.ActiveWindow.FreezePanes = $false
        `, { exclusive: true });
      }
      return textContent({ success: true });
    }
  );
}
