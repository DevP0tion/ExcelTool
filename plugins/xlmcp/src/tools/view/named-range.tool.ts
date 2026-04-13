import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_named_range",
    {
      title: "Named Ranges",
      description: "List, create, or delete named ranges.",
      inputSchema: {
        workbook: workbookParam,
        action: z.enum(["list", "add", "delete"]).describe("Action: list, add, or delete"),
        name: z.string().optional().describe("Name (required for add/delete)"),
        refersTo: z.string().optional().describe("Reference (required for add, e.g. '=Sheet1!$A$1:$D$10')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, action, name, refersTo }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';

      if (action === "list") {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $names = @()
          foreach ($n in $wb.Names) {
            $names += @{ Name = $n.Name; RefersTo = $n.RefersTo; Visible = [bool]$n.Visible }
          }
          ConvertTo-Json @($names) -Depth 3 -Compress
        `);
        return textContent(parseJSON(raw));
      }

      if (action === "add" && name && refersTo) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $wb.Names.Add('${psEscape(name)}', '${psEscape(refersTo)}')
        `);
        return textContent({ success: true });
      }

      if (action === "delete" && name) {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $wb.Names.Item('${psEscape(name)}').Delete()
        `);
        return textContent({ success: true });
      }

      return textContent({ error: "name and refersTo (for add) are required." });
    }
  );
}
