import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_read_cell_format",
    {
      title: "셀 서식 읽기",
      description: "셀의 현재 서식 정보를 반환합니다. 폰트, 배경색, 정렬, 테두리, 병합 여부 등을 포함합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("셀 주소 (예: A1)"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, cell }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $c = $ws.Range('${psEscape(cell)}')

        function OleToHex($color) {
          if ($color -eq $null) { return $null }
          $r = $color -band 0xFF
          $g = ($color -shr 8) -band 0xFF
          $b = ($color -shr 16) -band 0xFF
          return "{0:X2}{1:X2}{2:X2}" -f $r, $g, $b
        }

        $hAlignMap = @{
          -4131 = "left"
          -4108 = "center"
          -4152 = "right"
          1 = "general"
        }
        $vAlignMap = @{
          -4160 = "top"
          -4108 = "center"
          -4107 = "bottom"
        }
        $borderWeightMap = @{
          1 = "hairline"
          2 = "thin"
          -4138 = "medium"
          4 = "thick"
        }

        $borders = @{}
        $borderNames = @{7="left";8="top";9="bottom";10="right"}
        foreach ($idx in $borderNames.Keys) {
          $b = $c.Borders.Item($idx)
          if ($b.LineStyle -ne -4142) {
            $w = $borderWeightMap[[int]$b.Weight]
            if (-not $w) { $w = "thin" }
            $borders[$borderNames[$idx]] = @{
              style = $w
              color = OleToHex $b.Color
            }
          }
        }

        $bgColor = $null
        if ($c.Interior.Pattern -ne -4142) {
          $bgColor = OleToHex $c.Interior.Color
        }

        $ha = $hAlignMap[[int]$c.HorizontalAlignment]
        if (-not $ha) { $ha = "general" }
        $va = $vAlignMap[[int]$c.VerticalAlignment]
        if (-not $va) { $va = "bottom" }

        @{
          Font = @{
            Name = $c.Font.Name
            Size = $c.Font.Size
            Bold = [bool]$c.Font.Bold
            Italic = [bool]$c.Font.Italic
            Color = OleToHex $c.Font.Color
          }
          BgColor = $bgColor
          HAlign = $ha
          VAlign = $va
          WrapText = [bool]$c.WrapText
          NumberFormat = $c.NumberFormat
          MergeCells = [bool]$c.MergeCells
          Borders = $borders
        } | ConvertTo-Json -Depth 5 -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
