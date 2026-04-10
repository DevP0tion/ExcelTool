import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    parseInt(h.substring(0, 2), 16),
    parseInt(h.substring(2, 4), 16),
    parseInt(h.substring(4, 6), 16),
  ];
}

function rgbToOle([r, g, b]: [number, number, number]): number {
  return r + g * 256 + b * 65536;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_format_range",
    {
      title: "범위 서식",
      description:
        "범위에 서식을 적용합니다. 폰트(이름/크기/굵기/기울임/색상), 배경색, 정렬, 테두리, 표시 형식 등을 지정할 수 있습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("범위 주소 (예: A1:D10)"),
        fontName: z.string().optional().describe("폰트 이름 (예: '맑은 고딕')"),
        fontSize: z.number().optional().describe("폰트 크기"),
        bold: z.boolean().optional().describe("굵게"),
        italic: z.boolean().optional().describe("기울임"),
        fontColor: z.string().optional().describe("폰트 색상 RGB hex (예: 'FF0000')"),
        bgColor: z.string().optional().describe("배경 색상 RGB hex (예: 'FFFF00')"),
        hAlign: z
          .enum(["left", "center", "right"])
          .optional()
          .describe("가로 정렬"),
        vAlign: z
          .enum(["top", "center", "bottom"])
          .optional()
          .describe("세로 정렬"),
        wrapText: z.boolean().optional().describe("텍스트 줄바꿈"),
        numberFormat: z
          .string()
          .optional()
          .describe("표시 형식 (예: '#,##0', 'yyyy-mm-dd')"),
        border: z
          .enum(["thin", "medium", "thick", "none"])
          .optional()
          .describe("전체 테두리 스타일 (간편 옵션)"),
        borderEdges: z
          .object({
            left: z.enum(["thin", "medium", "thick", "none"]).optional(),
            right: z.enum(["thin", "medium", "thick", "none"]).optional(),
            top: z.enum(["thin", "medium", "thick", "none"]).optional(),
            bottom: z.enum(["thin", "medium", "thick", "none"]).optional(),
            outline: z.enum(["thin", "medium", "thick", "none"]).optional(),
            inside: z.enum(["thin", "medium", "thick", "none"]).optional(),
          })
          .optional()
          .describe("개별 테두리 제어. border와 함께 사용 시 borderEdges가 우선"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async (params) => {
      const { workbook, sheet, range, ...fmt } = params;
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      const cmds: string[] = [];

      if (fmt.fontName) cmds.push(`$r.Font.Name = '${psEscape(fmt.fontName)}'`);
      if (fmt.fontSize) cmds.push(`$r.Font.Size = ${fmt.fontSize}`);
      if (fmt.bold !== undefined) cmds.push(`$r.Font.Bold = $${fmt.bold}`);
      if (fmt.italic !== undefined) cmds.push(`$r.Font.Italic = $${fmt.italic}`);
      if (fmt.fontColor) {
        const rgb = hexToRgb(fmt.fontColor);
        cmds.push(`$r.Font.Color = ${rgbToOle(rgb)}`);
      }
      if (fmt.bgColor) {
        const rgb = hexToRgb(fmt.bgColor);
        cmds.push(`$r.Interior.Color = ${rgbToOle(rgb)}`);
      }
      if (fmt.hAlign) {
        const map = { left: -4131, center: -4108, right: -4152 };
        cmds.push(`$r.HorizontalAlignment = ${map[fmt.hAlign]}`);
      }
      if (fmt.vAlign) {
        const map = { top: -4160, center: -4108, bottom: -4107 };
        cmds.push(`$r.VerticalAlignment = ${map[fmt.vAlign]}`);
      }
      if (fmt.wrapText !== undefined) cmds.push(`$r.WrapText = $${fmt.wrapText}`);
      if (fmt.numberFormat) cmds.push(`$r.NumberFormat = '${psEscape(fmt.numberFormat)}'`);

      const weightMap: Record<string, number> = { thin: 2, medium: -4138, thick: 4 };
      const edgeIdxMap: Record<string, number[]> = {
        left: [7], right: [10], top: [8], bottom: [9],
        outline: [7, 8, 9, 10],
        inside: [11, 12],
      };

      function applyBorder(indices: number[], style: string) {
        if (style === "none") {
          for (const idx of indices) cmds.push(`$r.Borders.Item(${idx}).LineStyle = -4142`);
        } else {
          for (const idx of indices) {
            cmds.push(`$r.Borders.Item(${idx}).LineStyle = 1`);
            cmds.push(`$r.Borders.Item(${idx}).Weight = ${weightMap[style]}`);
          }
        }
      }

      if (fmt.borderEdges) {
        for (const [edge, style] of Object.entries(fmt.borderEdges)) {
          if (!style || !edgeIdxMap[edge]) continue;
          applyBorder(edgeIdxMap[edge], style);
        }
      } else if (fmt.border) {
        if (fmt.border === "none") {
          cmds.push(`$r.Borders.LineStyle = -4142`);
        } else {
          for (let i = 7; i <= 12; i++) {
            cmds.push(`$r.Borders.Item(${i}).LineStyle = 1`);
            cmds.push(`$r.Borders.Item(${i}).Weight = ${weightMap[fmt.border]}`);
          }
        }
      }

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        ${cmds.join("\n        ")}
      `);
      return textContent({ success: true });
    }
  );
}
