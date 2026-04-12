import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_paste_range",
    {
      title: "범위 복사/붙여넣기 (값·수식)",
      description: `범위의 값 또는 수식을 복사하여 대상 위치에 붙여넣습니다.
시스템 클립보드를 사용하지 않으므로 다른 작업과 안전하게 병렬 실행됩니다.

⚠️ 이 도구는 값(values)과 수식(formulas)만 복사합니다.
서식(폰트, 색상, 테두리 등)을 복사하려면 excel_copy_paste_format을 사용하세요.`,
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("원본 범위 (예: A1:C10)"),
        destCell: z.string().describe("붙여넣기 시작 셀 (예: E1)"),
        destSheet: z.string().optional().describe("대상 시트. 생략 시 같은 시트"),
        pasteType: z
          .enum(["values", "formulas"])
          .default("values")
          .describe("values: 계산된 값만 복사. formulas: 수식 원본 복사"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, destSheet, pasteType }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const dstShName = destSheet ? `'${psEscape(destSheet)}'` : shName;

      // 1. 소스 읽기
      const prop = pasteType === "formulas" ? "Formula" : "Value2";
      const readRaw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(sourceRange)}')
        $rows = $r.Rows.Count
        $cols = $r.Columns.Count
        $val = $r.${prop}
        $data = @()
        if ($rows -eq 1 -and $cols -eq 1) {
          $v = $val
          $data = ,@(,$(if ($v -ne $null) { $v } else { $null }))
        } elseif ($rows -eq 1) {
          $row = @()
          for ($j = 1; $j -le $cols; $j++) {
            $v = $val[1,$j]
            $row += $(if ($v -ne $null) { $v } else { $null })
          }
          $data = ,@($row)
        } else {
          for ($i = 1; $i -le $rows; $i++) {
            $row = @()
            for ($j = 1; $j -le $cols; $j++) {
              $v = $val[$i,$j]
              $row += $(if ($v -ne $null) { $v } else { $null })
            }
            $data += ,@($row)
          }
        }
        @{ Rows = $rows; Cols = $cols; Data = $data } | ConvertTo-Json -Depth 10 -Compress
      `);

      const { Rows: rows, Cols: cols, Data: data } = parseJSON<{
        Rows: number;
        Cols: number;
        Data: (string | number | null)[][];
      }>(readRaw);

      // 2. 대상에 쓰기
      if (pasteType === "formulas") {
        // 수식은 셀마다 다를 수 있으므로 배열 벌크 쓰기
        const formulaCmds: string[] = [];
        for (let i = 0; i < rows; i++) {
          for (let j = 0; j < cols; j++) {
            const v = data[i][j];
            if (v != null && String(v).startsWith("=")) {
              formulaCmds.push(
                `$dstWs.Cells.Item($dst.Row + ${i}, $dst.Column + ${j}).Formula = '${psEscape(String(v))}'`
              );
            } else if (v != null) {
              formulaCmds.push(
                `$dstWs.Cells.Item($dst.Row + ${i}, $dst.Column + ${j}).Value2 = '${psEscape(String(v))}'`
              );
            }
          }
        }
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $dstWs = Resolve-Sheet $wb ${dstShName}
          $dst = $dstWs.Range('${psEscape(destCell)}')
          ${formulaCmds.join("\n          ")}
        `);
      } else {
        // values: JSON 임시 파일 경유 벌크 쓰기
        const tmpPath = join(tmpdir(), `xlmcp_cp_${randomUUID()}.json`);
        writeFileSync(tmpPath, JSON.stringify(data));
        const escapedPath = tmpPath.replace(/\\/g, "\\\\");

        try {
          await runPS(`
            $wb = Resolve-Workbook ${wbName}
            $dstWs = Resolve-Sheet $wb ${dstShName}
            $dst = $dstWs.Range('${psEscape(destCell)}')
            $endCell = $dstWs.Cells.Item($dst.Row + ${rows} - 1, $dst.Column + ${cols} - 1)
            $targetRange = $dstWs.Range($dst, $endCell)
            $json = Get-Content '${escapedPath}' -Raw -Encoding UTF8
            $srcData = $json | ConvertFrom-Json
            $arr = New-Object 'object[,]' ${rows},${cols}
            for ($i = 0; $i -lt ${rows}; $i++) {
              for ($j = 0; $j -lt ${cols}; $j++) {
                $v = $srcData[$i][$j]
                if ($v -ne $null) { $arr[$i,$j] = $v }
              }
            }
            $targetRange.Value2 = $arr
          `);
        } finally {
          try { unlinkSync(tmpPath); } catch { /* ignore */ }
        }
      }

      return textContent({ success: true, rows, cols, pasteType });
    }
  );
}
