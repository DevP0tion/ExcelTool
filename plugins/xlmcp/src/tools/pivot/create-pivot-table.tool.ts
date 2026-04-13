import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_pivot_table",
    {
      title: "Create Pivot Table",
      description: "Create pivot table from data range.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("Source range (e.g. A1:D100)"),
        destCell: z.string().describe("Pivot table location (e.g. F1 or Sheet2!A1)"),
        name: z.string().optional().describe("Pivot table name"),
        rowFields: z.array(z.string()).optional().describe("Row field names"),
        columnFields: z.array(z.string()).optional().describe("Column field names"),
        dataFields: z
          .array(
            z.object({
              field: z.string().describe("Field name"),
              function: z
                .enum(["sum", "count", "average", "max", "min"])
                .default("sum")
                .describe("Aggregate function"),
            })
          )
          .optional()
          .describe("Data field array"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, name, rowFields, columnFields, dataFields }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const pvtName = name ? `'${psEscape(name)}'` : "'PivotTable1'";
      // xlSum=-4157, xlCount=-4112, xlAverage=-4106, xlMax=-4136, xlMin=-4139
      const fnMap: Record<string, number> = { sum: -4157, count: -4112, average: -4106, max: -4136, min: -4139 };

      const fieldCmds: string[] = [];
      if (rowFields) {
        for (const f of rowFields) {
          fieldCmds.push(`$pf = $pvt.PivotFields('${psEscape(f)}'); $pf.Orientation = 1`); // xlRowField
        }
      }
      if (columnFields) {
        for (const f of columnFields) {
          fieldCmds.push(`$pf = $pvt.PivotFields('${psEscape(f)}'); $pf.Orientation = 2`); // xlColumnField
        }
      }
      if (dataFields) {
        for (const d of dataFields) {
          fieldCmds.push(`$pf = $pvt.PivotFields('${psEscape(d.field)}'); $pf.Orientation = 4; $pf.Function = ${fnMap[d.function]}`); // xlDataField
        }
      }

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $src = $ws.Range('${psEscape(sourceRange)}')
        $cache = $wb.PivotCaches().Create(1, $src)
        $dest = $wb.Worksheets.Item($ws.Name).Range('${psEscape(destCell)}')
        $pvt = $cache.CreatePivotTable($dest, ${pvtName})
        ${fieldCmds.join("\n        ")}
        @{ Name = $pvt.Name; Location = $pvt.TableRange1.Address() } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
