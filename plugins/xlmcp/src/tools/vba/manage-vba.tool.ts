import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_manage_vba",
    {
      title: "VBA 모듈 관리",
      description: `VBA 모듈을 삭제, 코드 교체, 코드 추가합니다.
Sheet/ThisWorkbook 등 Document 모듈은 삭제 불가하며 코드만 제거됩니다.
⚠️ "프로그래밍 방식 VBA 프로젝트 액세스 신뢰"가 켜져 있어야 합니다.`,
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("모듈 이름"),
        action: z
          .enum(["delete", "replace", "append"])
          .describe("delete(삭제), replace(코드 전체 교체), append(코드 끝에 추가)"),
        code: z.string().optional().describe("replace/append 시 VBA 코드"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, name, action, code }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';

      if (action === "delete") {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $comp = $wb.VBProject.VBComponents.Item('${psEscape(name)}')
          # Document(100) 타입은 삭제 불가 → 코드만 제거
          if ($comp.Type -eq 100) {
            if ($comp.CodeModule.CountOfLines -gt 0) {
              $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
            }
            @{ Action = 'code_cleared'; Name = $comp.Name; Message = 'Document 모듈은 삭제 불가, 코드만 제거됨' } | ConvertTo-Json -Compress
          } else {
            $wb.VBProject.VBComponents.Remove($comp)
            @{ Action = 'deleted'; Name = '${psEscape(name)}' } | ConvertTo-Json -Compress
          }
        `);
        return textContent(parseJSON(raw));
      }

      // replace / append — code 필수
      if (!code) throw new Error(`${action} 시 code 파라미터가 필요합니다.`);

      const tmpPath = join(tmpdir(), `xlmcp_vba_${randomUUID()}.txt`);
      writeFileSync(tmpPath, code.trim(), "utf-8");
      const escapedPath = tmpPath.replace(/\\/g, "\\\\");

      try {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $comp = $wb.VBProject.VBComponents.Item('${psEscape(name)}')
          $cm = $comp.CodeModule
          $newCode = [System.IO.File]::ReadAllText('${escapedPath}', (New-Object System.Text.UTF8Encoding $false))
          ${action === "replace"
            ? `if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
          $cm.InsertLines(1, $newCode)`
            : `$existing = ''
          if ($cm.CountOfLines -gt 0) { $existing = $cm.Lines(1, $cm.CountOfLines) }
          if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
          $cm.InsertLines(1, $existing + [char]13 + [char]10 + $newCode)`
          }
          @{
            Action = '${action}'
            Name = $comp.Name
            LineCount = $cm.CountOfLines
          } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      } finally {
        try { unlinkSync(tmpPath); } catch { /* ignore */ }
      }
    }
  );
}
