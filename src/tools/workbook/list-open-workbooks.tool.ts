import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_open_workbooks",
    {
      title: "열린 워크북 목록",
      description: "현재 Excel에 열려 있는 모든 워크북의 이름과 경로를 반환합니다.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const raw = await runPS(`
        $result = @()
        foreach ($wb in $excel.Workbooks) {
          $result += @{ Name = $wb.Name; Path = $wb.FullName; Sheets = $wb.Worksheets.Count } | ConvertTo-Json -Compress
        }
        "[" + ($result -join ",") + "]"
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
