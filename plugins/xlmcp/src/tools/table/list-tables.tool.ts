import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_tables",
    {
      title: "List Tables",
      description: "Returns all tables: name, range, style, row/column count.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $tables = @()
        foreach ($t in $ws.ListObjects) {
          $tables += @{
            Name = $t.Name
            Range = $t.Range.Address()
            Style = if ($t.TableStyle) { $t.TableStyle.Name } else { $null }
            Rows = $t.ListRows.Count
            Columns = $t.ListColumns.Count
          }
        }
        ConvertTo-Json @($tables) -Depth 3 -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
