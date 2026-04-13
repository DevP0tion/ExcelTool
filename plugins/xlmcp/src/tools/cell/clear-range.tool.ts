import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_clear_range",
    {
      title: "Clear Range",
      description: "Clear content, formats, or all in a range.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("Range address (e.g. A1:D10)"),
        mode: z
          .enum(["values", "formats", "all"])
          .default("all")
          .describe("Target: values, formats, or all"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, range, mode }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const cmdMap: Record<string, string> = {
        values: "$r.ClearContents()",
        formats: "$r.ClearFormats()",
        all: "$r.Clear()",
      };
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')
        ${cmdMap[mode]}
      `);
      return textContent({ success: true });
    }
  );
}
