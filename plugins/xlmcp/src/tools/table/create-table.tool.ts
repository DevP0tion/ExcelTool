import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_table",
    {
      title: "Create Table",
      description: "Convert range to Excel table (ListObject).",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range to convert (e.g. A1:D10)"),
        name: z.string().optional().describe("Table name. Auto if omitted"),
        hasHeader: z.boolean().default(true).describe("Use first row as header"),
        style: z.string().optional().describe("Table style (e.g. 'TableStyleMedium2')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, range, name, hasHeader, style }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      // xlYes=1, xlNo=2
      const xlHeader = hasHeader ? 1 : 2;
      const nameCmd = name ? `$t.Name = '${psEscape(name)}'` : "";
      const styleCmd = style ? `$t.TableStyle = '${psEscape(style)}'` : "";
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        $t = $ws.ListObjects.Add(1, $r, [Type]::Missing, ${xlHeader})
        ${nameCmd}
        ${styleCmd}
        @{ Name = $t.Name; Range = $t.Range.Address() } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
