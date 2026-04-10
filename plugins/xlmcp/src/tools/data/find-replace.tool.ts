import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_find_replace",
    {
      title: "찾기/바꾸기",
      description: "시트 내에서 텍스트를 찾거나 바꿉니다. replace를 생략하면 찾기만 수행합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        find: z.string().describe("찾을 텍스트"),
        replace: z.string().optional().describe("바꿀 텍스트. 생략 시 찾기만 수행"),
        matchCase: z.boolean().default(false).describe("대소문자 구분"),
        range: z.string().optional().describe("검색 범위. 생략 시 전체 시트"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, find, replace, matchCase, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range ? `$ws.Range('${psEscape(range)}')` : `$ws.Cells`;

      if (replace !== undefined) {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = ${rangeExpr}
          $replaced = $r.Replace('${psEscape(find)}', '${psEscape(replace)}', [Type]::Missing, [Type]::Missing, ${matchCase ? "$true" : "$false"})
          @{ Success = $replaced } | ConvertTo-Json -Compress
        `);
        return textContent(JSON.parse(raw.trim()));
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
        return textContent(JSON.parse(raw.trim()));
      }
    }
  );
}
