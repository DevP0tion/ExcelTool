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
      title: "Manage VBA Module",
      description: "Delete/replace/append VBA module code. Document modules: code cleared only.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("Module name"),
        action: z
          .enum(["delete", "replace", "append"])
          .describe("delete, replace (all code), or append (to end)"),
        code: z.string().optional().describe("VBA code for replace/append"),
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
            @{ Action = 'code_cleared'; Name = $comp.Name; Message = 'Document module cannot be deleted. Code cleared only.' } | ConvertTo-Json -Compress
          } else {
            $wb.VBProject.VBComponents.Remove($comp)
            @{ Action = 'deleted'; Name = '${psEscape(name)}' } | ConvertTo-Json -Compress
          }
        `);
        return textContent(parseJSON(raw));
      }

      // replace / append — code 필수
      if (!code) throw new Error(`Code parameter is required for ${action}.`);

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
