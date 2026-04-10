import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_named_range",
    {
      title: "이름 정의 관리",
      description: "이름 정의(Named Range)를 조회, 생성, 삭제합니다.",
      inputSchema: {
        workbook: workbookParam,
        action: z.enum(["list", "add", "delete"]).describe("동작: list(조회), add(생성), delete(삭제)"),
        name: z.string().optional().describe("이름 (add/delete 시 필수)"),
        refersTo: z.string().optional().describe("참조 범위 (add 시 필수, 예: '=Sheet1!$A$1:$D$10')"),
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
            $names += @{ Name = $n.Name; RefersTo = $n.RefersTo; Visible = [bool]$n.Visible } | ConvertTo-Json -Compress
          }
          "[" + ($names -join ",") + "]"
        `);
        return textContent(JSON.parse(raw.trim()));
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

      return textContent({ error: "name과 refersTo(add 시)가 필요합니다." });
    }
  );
}
