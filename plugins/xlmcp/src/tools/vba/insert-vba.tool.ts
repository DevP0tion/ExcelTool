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
    "excel_insert_vba",
    {
      title: "Add VBA Module",
      description: "Add VBA module with code. Requires VBA project access trust.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("Module name (e.g. MyMacro)"),
        code: z.string().describe("VBA source code"),
        type: z
          .enum(["module", "classModule", "form"])
          .default("module")
          .describe("Type: module, classModule, or form"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, name, code, type }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      // vbext_ct_StdModule=1, vbext_ct_ClassModule=2, vbext_ct_MSForm=3
      const typeMap: Record<string, number> = { module: 1, classModule: 2, form: 3 };

      // VBA 코드는 줄바꿈 보존 필요 → 임시 파일 경유
      const tmpPath = join(tmpdir(), `xlmcp_vba_${randomUUID()}.txt`);
      writeFileSync(tmpPath, code.trim(), "utf-8");
      const escapedPath = tmpPath.replace(/\\/g, "\\\\");

      try {
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $comp = $wb.VBProject.VBComponents.Add(${typeMap[type]})
          $comp.Name = '${psEscape(name)}'
          $code = [System.IO.File]::ReadAllText('${escapedPath}', (New-Object System.Text.UTF8Encoding $false))
          $comp.CodeModule.InsertLines(1, $code)
          @{
            Name = $comp.Name
            Type = '${type}'
            LineCount = $comp.CodeModule.CountOfLines
          } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      } finally {
        try { unlinkSync(tmpPath); } catch { /* ignore */ }
      }
    }
  );
}
