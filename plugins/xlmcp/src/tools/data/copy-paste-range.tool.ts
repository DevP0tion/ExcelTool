import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const DEFAULT_CHUNK_SIZE = 30;

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
        chunkSize: z.number().int().optional().describe("청크 분할 행수. 기본 30"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, destSheet, pasteType, chunkSize: cs }) => {
      const chunkSize = cs ?? DEFAULT_CHUNK_SIZE;
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

      // 2. 쓰기 — values와 formulas 모두 JSON 파일 + 벌크 경로
      const assignProp = pasteType === "formulas" ? "Formula" : "Value2";

      if (rows < chunkSize) {
        // 소규모: 단일 벌크
        await writeBulk(wbName, dstShName, destCell, data, rows, cols, assignProp);
      } else {
        // 대규모: 청크 분할 + 병렬 + Calculation 억제
        await writeChunked(wbName, dstShName, destCell, data, rows, cols, assignProp, chunkSize);
      }

      return textContent({ success: true, rows, cols, pasteType });
    }
  );
}

// ── 단일 벌크 쓰기 ──
async function writeBulk(
  wbName: string,
  shName: string,
  destCell: string,
  data: (string | number | null)[][],
  rows: number,
  cols: number,
  prop: string
): Promise<void> {
  const tmpPath = join(tmpdir(), `xlmcp_cp_${randomUUID()}.json`);
  writeFileSync(tmpPath, JSON.stringify(data));
  const escapedPath = tmpPath.replace(/\\/g, "\\\\");

  try {
    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $dstWs = Resolve-Sheet $wb ${shName}
      $dst = $dstWs.Range('${psEscape(destCell)}')
      $targetRange = $dstWs.Range($dst, $dstWs.Cells.Item($dst.Row + ${rows} - 1, $dst.Column + ${cols} - 1))
      $json = Get-Content '${escapedPath}' -Raw -Encoding UTF8
      $srcData = $json | ConvertFrom-Json
      $arr = New-Object 'object[,]' ${rows},${cols}
      for ($i = 0; $i -lt ${rows}; $i++) {
        for ($j = 0; $j -lt ${cols}; $j++) {
          $v = $srcData[$i][$j]
          if ($v -ne $null) { $arr[$i,$j] = $v }
        }
      }
      $targetRange.${prop} = $arr
    `);
  } finally {
    try { unlinkSync(tmpPath); } catch { /* ignore */ }
  }
}

// ── 청크 분할 + 병렬 + Calculation 억제 ──
async function writeChunked(
  wbName: string,
  shName: string,
  destCell: string,
  data: (string | number | null)[][],
  rows: number,
  cols: number,
  prop: string,
  chunkSize: number
): Promise<void> {
  // Calculation/ScreenUpdating 억제
  await runPS(`
    $wb = Resolve-Workbook ${wbName}
    $excel.ScreenUpdating = $false
    $excel.Calculation = -4135
  `);

  // 청크 생성 + 임시 파일
  const chunks: { offset: number; chunkRows: number }[] = [];
  const batchId = randomUUID();
  const tmpFiles: string[] = [];

  for (let offset = 0; offset < rows; offset += chunkSize) {
    const chunkRows = Math.min(chunkSize, rows - offset);
    chunks.push({ offset, chunkRows });
    const chunkData = data.slice(offset, offset + chunkRows);
    const filePath = join(tmpdir(), `xlmcp_cp_${batchId}_${chunks.length - 1}.json`);
    writeFileSync(filePath, JSON.stringify(chunkData));
    tmpFiles.push(filePath);
  }

  try {
    // 병렬 쓰기
    await Promise.all(
      chunks.map((chunk, i) => {
        const escapedPath = tmpFiles[i].replace(/\\/g, "\\\\");
        return runPS(`
          $wb = Resolve-Workbook ${wbName}
          $dstWs = Resolve-Sheet $wb ${shName}
          $dst = $dstWs.Range('${psEscape(destCell)}')
          $chunkStart = $dstWs.Cells.Item($dst.Row + ${chunk.offset}, $dst.Column)
          $chunkEnd = $dstWs.Cells.Item($dst.Row + ${chunk.offset} + ${chunk.chunkRows} - 1, $dst.Column + ${cols} - 1)
          $targetRange = $dstWs.Range($chunkStart, $chunkEnd)
          $json = Get-Content '${escapedPath}' -Raw -Encoding UTF8
          $srcData = $json | ConvertFrom-Json
          $arr = New-Object 'object[,]' ${chunk.chunkRows},${cols}
          for ($i = 0; $i -lt ${chunk.chunkRows}; $i++) {
            for ($j = 0; $j -lt ${cols}; $j++) {
              $v = $srcData[$i][$j]
              if ($v -ne $null) { $arr[$i,$j] = $v }
            }
          }
          $targetRange.${prop} = $arr
        `);
      })
    );
  } finally {
    // 임시 파일 정리
    for (const f of tmpFiles) {
      try { unlinkSync(f); } catch { /* ignore */ }
    }
    // Calculation/ScreenUpdating 복원
    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $excel.Calculation = -4105
      $excel.ScreenUpdating = $true
    `).catch(() => { /* 복원 실패 무시 */ });
  }
}
