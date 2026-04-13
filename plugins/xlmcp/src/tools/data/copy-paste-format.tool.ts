import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_paste_format",
    {
      title: "Copy/Paste Format",
      description: "Copy/paste formatting (font, color, borders). Uses clipboard (exclusive).",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("Source range (e.g. A1:C10)"),
        destCell: z.string().describe("Paste start cell (e.g. E1)"),
        destSheet: z.string().optional().describe("Destination sheet. Same sheet if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, destSheet }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const dstShName = destSheet ? `'${psEscape(destSheet)}'` : shName;
      // xlPasteFormats = -4122
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $srcWs = Resolve-Sheet $wb ${shName}
        $dstWs = Resolve-Sheet $wb ${dstShName}
        $srcWs.Range('${psEscape(sourceRange)}').Copy()
        $dstWs.Range('${psEscape(destCell)}').PasteSpecial(-4122)
        $excel.CutCopyMode = $false
      `, { exclusive: true });
      return textContent({ success: true });
    }
  );
}
