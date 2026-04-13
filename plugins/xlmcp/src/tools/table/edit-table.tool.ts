import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_edit_table",
    {
      title: "Edit Table",
      description: "Edit table name, style, size, or add rows/columns.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("Table name"),
        newName: z.string().optional().describe("New table name"),
        style: z.string().optional().describe("Table style (e.g. 'TableStyleLight1')"),
        resize: z.string().optional().describe("Resize to new range (e.g. 'A1:E20')"),
        addRow: z.boolean().optional().describe("Add row at bottom"),
        addColumn: z.string().optional().describe("Column name to add"),
        showTotals: z.boolean().optional().describe("Show/hide totals row"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, name, newName, style, resize, addRow, addColumn, showTotals }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmds: string[] = [];
      if (newName) cmds.push(`$t.Name = '${psEscape(newName)}'`);
      if (style) cmds.push(`$t.TableStyle = '${psEscape(style)}'`);
      if (resize) cmds.push(`$t.Resize($ws.Range('${psEscape(resize)}'))`);
      if (addRow) cmds.push(`$t.ListRows.Add() | Out-Null`);
      if (addColumn) cmds.push(`$t.ListColumns.Add().Name = '${psEscape(addColumn)}'`);
      if (showTotals !== undefined) cmds.push(`$t.ShowTotals = $${showTotals}`);

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $t = $ws.ListObjects.Item('${psEscape(name)}')
        ${cmds.join("\n        ")}
      `);
      return textContent({ success: true });
    }
  );
}
