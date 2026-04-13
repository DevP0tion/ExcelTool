import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_find_replace",
    {
      title: "Find/Replace",
      description: "Find or replace text. Find-only if replace is omitted.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        find: z.string().describe("Text to find"),
        replace: z.string().optional().describe("Replacement text. Find-only if omitted"),
        matchCase: z.boolean().default(false).describe("Case-sensitive"),
        range: z.string().optional().describe("Search range. Entire sheet if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, find, replace, matchCase, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range ? `$ws.Range('${psEscape(range)}')` : `$ws.UsedRange`;

      if (replace !== undefined) {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = ${rangeExpr}
          $replaced = $r.Replace('${psEscape(find)}', '${psEscape(replace)}', [Type]::Missing, [Type]::Missing, ${matchCase ? "$true" : "$false"})
          @{ Success = $replaced } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      } else {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = ${rangeExpr}
          $results = @()
          $first = $r.Find('${psEscape(find)}', [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, ${matchCase ? "$true" : "$false"})
          if ($first) {
            $current = $first
            do {
              $results += @{ Cell = $current.Address(); Value = $current.Value2 }
              $current = $r.FindNext($current)
            } while ($current -and $current.Address() -ne $first.Address())
          }
          @{ Count = $results.Count; Matches = $results } | ConvertTo-Json -Depth 5 -Compress
        `);
        return textContent(parseJSON(raw));
      }
    }
  );
}
