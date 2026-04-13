import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_create_workbook",
    {
      title: "Create Workbook",
      description: "Create a new empty workbook. Saves immediately if savePath is given.",
      inputSchema: {
        savePath: z.string().optional().describe("Absolute save path (.xlsx). Not saved if omitted"),
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
      return textContent(parseJSON(raw));
    }
  );
}
