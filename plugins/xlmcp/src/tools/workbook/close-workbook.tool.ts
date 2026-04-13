import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_close_workbook",
    {
      title: "Close Workbook",
      description: "Close workbook. Set save to true to save before closing.",
      inputSchema: {
        workbook: workbookParam,
        save: z.boolean().default(false).describe("Save before closing"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, save }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Close(${save ? "$true" : "$false"})
      `);
      return textContent({ success: true });
    }
  );
}
