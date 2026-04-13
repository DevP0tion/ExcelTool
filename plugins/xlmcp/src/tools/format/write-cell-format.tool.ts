import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, hexToRgb, rgbToOle } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const borderSchema = z
  .object({
    style: z.enum(["hairline", "thin", "medium", "thick"]).optional(),
    color: z.string().optional().describe("RGB hex (e.g. 'FF0000')"),
  })
  .optional();

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_cell_format",
    {
      title: "Write Cell Format",
      description:
        "Apply formatting using read_cell_format structure. Works on ranges.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range address (e.g. A1 or A1:D10)"),
        font: z
          .object({
            name: z.string().optional(),
            size: z.number().optional(),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            color: z.string().optional().describe("RGB hex"),
          })
          .optional()
          .describe("Font settings"),
        bgColor: z.string().optional().describe("Background color RGB hex"),
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
          .describe("Per-edge border settings"),
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
          cmds.push(`$r.Font.Color = ${rgbToOle(hexToRgb(font.color))}`);
        }
      }

      if (bgColor) {
        cmds.push(`$r.Interior.Color = ${rgbToOle(hexToRgb(bgColor))}`);
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
            cmds.push(`$r.Borders.Item(${idx}).Color = ${rgbToOle(hexToRgb(cfg.color))}`);
          }
        }
      }

      if (cmds.length === 0) return textContent({ success: true, message: "No changes" });

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
