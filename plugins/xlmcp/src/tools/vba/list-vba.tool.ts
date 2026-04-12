import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

const TYPE_NAMES: Record<number, string> = { 1: "module", 2: "classModule", 3: "form", 100: "document" };

export function register(server: McpServer) {
  server.registerTool(
    "excel_list_vba",
    {
      title: "VBA 모듈 목록",
      description: `VBA 모듈 목록을 반환합니다. name을 지정하면 해당 모듈의 소스 코드도 반환합니다.
⚠️ "프로그래밍 방식 VBA 프로젝트 액세스 신뢰"가 켜져 있어야 합니다.`,
      inputSchema: {
        workbook: workbookParam,
        name: z.string().optional().describe("특정 모듈 이름. 지정 시 해당 모듈의 코드 반환"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, name }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';

      if (name) {
        // 특정 모듈 코드 읽기
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $comp = $wb.VBProject.VBComponents.Item('${psEscape(name)}')
          $cm = $comp.CodeModule
          $code = ''
          if ($cm.CountOfLines -gt 0) {
            $code = $cm.Lines(1, $cm.CountOfLines)
          }
          @{
            Name = $comp.Name
            Type = $comp.Type
            Lines = $cm.CountOfLines
            Code = $code
          } | ConvertTo-Json -Depth 3 -Compress
        `);
        const parsed = parseJSON<{ Name: string; Type: number; Lines: number; Code: string }>(raw);
        return textContent({
          ...parsed,
          Type: TYPE_NAMES[parsed.Type] ?? String(parsed.Type),
        });
      }

      // 전체 목록
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $list = @()
        foreach ($comp in $wb.VBProject.VBComponents) {
          $list += @{
            Name = $comp.Name
            Type = $comp.Type
            Lines = $comp.CodeModule.CountOfLines
          }
        }
        ConvertTo-Json @($list) -Depth 3 -Compress
      `);
      const parsed = parseJSON<Array<{ Name: string; Type: number; Lines: number }>>(raw);
      return textContent(
        parsed.map((m) => ({
          ...m,
          Type: TYPE_NAMES[m.Type] ?? String(m.Type),
        }))
      );
    }
  );
}
