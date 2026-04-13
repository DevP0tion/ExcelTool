import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_delete_table",
    {
      title: "Delete Table",
      description: "Remove table. Data preserved unless clearData is true.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("Table name to delete"),
        clearData: z.boolean().default(false).describe("Also delete data"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, name, clearData }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmd = clearData
        ? `$t.Delete()`
        : `$t.Unlist()`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $t = $ws.ListObjects.Item('${psEscape(name)}')
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );
}
