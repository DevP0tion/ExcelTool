import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_chart",
    {
      title: "Create Chart",
      description: "Create chart from data range.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        dataRange: z.string().describe("Chart data range (e.g. A1:D10)"),
        chartType: z
          .enum(["line", "bar", "column", "pie", "scatter", "area"])
          .default("column")
          .describe("Chart type"),
        title: z.string().optional().describe("Chart title"),
        position: z.string().optional().describe("Chart position cell (e.g. F1). Auto if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, dataRange, chartType, title, position }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      // xlLine=4, xlBar=2, xlColumnClustered=51, xlPie=5, xlXYScatter=-4169, xlArea=1
      const typeMap: Record<string, number> = {
        line: 4, bar: 2, column: 51, pie: 5, scatter: -4169, area: 1,
      };
      const titleCmd = title ? `$chart.Chart.HasTitle = $true; $chart.Chart.ChartTitle.Text = '${psEscape(title)}'` : "";
      const posCmd = position
        ? `$pos = $ws.Range('${psEscape(position)}'); $chart.Left = $pos.Left; $chart.Top = $pos.Top`
        : "";
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(dataRange)}')
        $chart = $ws.Shapes.AddChart2([Type]::Missing, ${typeMap[chartType]}, [Type]::Missing, [Type]::Missing, 400, 300)
        $chart.Chart.SetSourceData($r)
        ${titleCmd}
        ${posCmd}
        @{ Name = $chart.Name } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
