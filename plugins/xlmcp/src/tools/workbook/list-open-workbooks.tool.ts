import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { runPS } from "../../services/powershell.js";
import { textContent, parseJSON } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_open_workbooks",
    {
      title: "List Open Workbooks",
      description: "Returns names and paths of all open workbooks.",
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
      return textContent(parseJSON(raw));
    }
  );
}
