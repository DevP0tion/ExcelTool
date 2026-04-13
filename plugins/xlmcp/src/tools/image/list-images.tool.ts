import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_images",
    {
      title: "List Images",
      description: "Returns list of images (Pictures) in the sheet.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $imgs = @()
        foreach ($s in $ws.Shapes) {
          # Type 13 = msoPicture, 11 = msoLinkedPicture
          if ($s.Type -eq 13 -or $s.Type -eq 11) {
            $imgs += @{
              Name = $s.Name
              Width = [math]::Round($s.Width, 1)
              Height = [math]::Round($s.Height, 1)
              Left = [math]::Round($s.Left, 1)
              Top = [math]::Round($s.Top, 1)
            }
          }
        }
        ConvertTo-Json @($imgs) -Depth 3 -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
