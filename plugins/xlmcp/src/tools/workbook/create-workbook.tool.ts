import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_workbook",
    {
      title: "새 워크북 생성",
      description: "새 빈 워크북을 생성합니다. savePath를 지정하면 즉시 저장합니다.",
      inputSchema: {
        savePath: z.string().optional().describe("저장할 절대 경로 (.xlsx). 생략 시 저장하지 않음"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ savePath }) => {
      const saveCmd = savePath
        ? `$wb.SaveAs('${psEscape(savePath)}')`
        : "";
      const raw = await runPS(`
        $wb = $excel.Workbooks.Add()
        ${saveCmd}
        @{ Name = $wb.Name; Path = $wb.FullName } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );
}
