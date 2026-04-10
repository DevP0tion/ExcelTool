import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_chart",
    {
      title: "차트 생성",
      description: "데이터 범위로 차트를 생성합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        dataRange: z.string().describe("차트 데이터 범위 (예: A1:D10)"),
        chartType: z
          .enum(["line", "bar", "column", "pie", "scatter", "area"])
          .default("column")
          .describe("차트 유형"),
        title: z.string().optional().describe("차트 제목"),
        position: z.string().optional().describe("차트 위치 셀 (예: F1). 생략 시 자동"),
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
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
