import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, hexToRgb, rgbToOle } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_format_range",
    {
      title: "Format Range",
      description:
        "Apply formatting: font, background, alignment, borders, number format.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range address (e.g. A1:D10)"),
        fontName: z.string().optional().describe("Font name (e.g. 'Arial')"),
        fontSize: z.number().optional().describe("Font size"),
        bold: z.boolean().optional().describe("Bold"),
        italic: z.boolean().optional().describe("Italic"),
        fontColor: z.string().optional().describe("Font color RGB hex (e.g. 'FF0000')"),
        bgColor: z.string().optional().describe("Background color RGB hex (e.g. 'FFFF00')"),
        hAlign: z
          .enum(["left", "center", "right"])
          .optional()
          .describe("Horizontal alignment"),
        vAlign: z
          .enum(["top", "center", "bottom"])
          .optional()
          .describe("Vertical alignment"),
        wrapText: z.boolean().optional().describe("Wrap text"),
        numberFormat: z
          .string()
          .optional()
          .describe("Number format (e.g. '#,##0', 'yyyy-mm-dd')"),
        border: z
          .enum(["thin", "medium", "thick", "none"])
          .optional()
          .describe("Border style for all edges (shortcut)"),
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
          .describe("Per-edge border control. Overrides border if both set"),
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
