import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_sheet",
    {
      title: "시트 복사",
      description: "시트를 복사합니다. 같은 워크북 내에서 복사됩니다.",
      inputSchema: {
        workbook: workbookParam,
        source: z.string().describe("원본 시트 이름"),
        newName: z.string().optional().describe("복사본 시트 이름. 생략 시 자동 이름"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, source, newName }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const renameCmd = newName
        ? `$wb.ActiveSheet.Name = '${psEscape(newName)}'`
        : "";
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $src = $wb.Worksheets.Item('${psEscape(source)}')
        $src.Copy([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        ${renameCmd}
        $wb.ActiveSheet.Name
      `);
      return textContent({ success: true, name: raw.trim() });
    }
  );
}
