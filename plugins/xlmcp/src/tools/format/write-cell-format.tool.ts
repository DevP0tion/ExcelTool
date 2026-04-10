import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const borderSchema = z
  .object({
    style: z.enum(["hairline", "thin", "medium", "thick"]).optional(),
    color: z.string().optional().describe("RGB hex (예: 'FF0000')"),
  })
  .optional();

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_cell_format",
    {
      title: "셀 서식 일괄 적용",
      description:
        "read_cell_format의 출력 형식과 동일한 구조로 서식을 일괄 적용합니다. 범위에도 적용 가능합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("범위 주소 (예: A1 또는 A1:D10)"),
        font: z
          .object({
            name: z.string().optional(),
            size: z.number().optional(),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            color: z.string().optional().describe("RGB hex"),
          })
          .optional()
          .describe("폰트 설정"),
        bgColor: z.string().optional().describe("배경 색상 RGB hex"),
        hAlign: z.enum(["left", "center", "right", "general"]).optional(),
        vAlign: z.enum(["top", "center", "bottom"]).optional(),
        wrapText: z.boolean().optional(),
        numberFormat: z.string().optional(),
        borders: z
          .object({
            left: borderSchema,
            top: borderSchema,
            right: borderSchema,
            bottom: borderSchema,
          })
          .optional()
          .describe("개별 테두리 설정"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async (params) => {
      const { workbook, sheet, range, font, bgColor, hAlign, vAlign, wrapText, numberFormat, borders } = params;
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      const cmds: string[] = [];

      if (font) {
        if (font.name) cmds.push(`$r.Font.Name = '${psEscape(font.name)}'`);
        if (font.size) cmds.push(`$r.Font.Size = ${font.size}`);
        if (font.bold !== undefined) cmds.push(`$r.Font.Bold = $${font.bold}`);
        if (font.italic !== undefined) cmds.push(`$r.Font.Italic = $${font.italic}`);
        if (font.color) {
          const [r, g, b] = hexToRgb(font.color);
          cmds.push(`$r.Font.Color = ${r + g * 256 + b * 65536}`);
        }
      }

      if (bgColor) {
        const [r, g, b] = hexToRgb(bgColor);
        cmds.push(`$r.Interior.Color = ${r + g * 256 + b * 65536}`);
      }

      if (hAlign) {
        const map: Record<string, number> = { left: -4131, center: -4108, right: -4152, general: 1 };
        cmds.push(`$r.HorizontalAlignment = ${map[hAlign]}`);
      }
      if (vAlign) {
        const map: Record<string, number> = { top: -4160, center: -4108, bottom: -4107 };
        cmds.push(`$r.VerticalAlignment = ${map[vAlign]}`);
      }
      if (wrapText !== undefined) cmds.push(`$r.WrapText = $${wrapText}`);
      if (numberFormat) cmds.push(`$r.NumberFormat = '${psEscape(numberFormat)}'`);

      if (borders) {
        const idxMap: Record<string, number> = { left: 7, top: 8, bottom: 9, right: 10 };
        const weightMap: Record<string, number> = { hairline: 1, thin: 2, medium: -4138, thick: 4 };
        for (const [side, cfg] of Object.entries(borders)) {
          if (!cfg) continue;
          const idx = idxMap[side];
          cmds.push(`$r.Borders.Item(${idx}).LineStyle = 1`);
          if (cfg.style) cmds.push(`$r.Borders.Item(${idx}).Weight = ${weightMap[cfg.style]}`);
          if (cfg.color) {
            const [r, g, b] = hexToRgb(cfg.color);
            cmds.push(`$r.Borders.Item(${idx}).Color = ${r + g * 256 + b * 65536}`);
          }
        }
      }

      if (cmds.length === 0) return textContent({ success: true, message: "변경 사항 없음" });

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

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    parseInt(h.substring(0, 2), 16),
    parseInt(h.substring(2, 4), 16),
    parseInt(h.substring(4, 6), 16),
  ];
}
