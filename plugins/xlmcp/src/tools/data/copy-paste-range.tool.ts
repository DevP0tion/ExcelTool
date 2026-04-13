import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { writeFileSync, readFileSync, unlinkSync } from "fs";
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
      title: "Copy/Paste Range (Values & Formulas)",
      description: "Copy/paste values or formulas. No clipboard, parallel-safe.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("Source range (e.g. A1:C10)"),
        destCell: z.string().describe("Paste start cell (e.g. E1)"),
        destSheet: z.string().optional().describe("Destination sheet. Same sheet if omitted"),
        pasteType: z
          .enum(["values", "formulas"])
          .default("values")
          .describe("values: computed values. formulas: original formulas"),
        chunkSize: z.number().int().optional().describe("Chunk size in rows. Default 30"),
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

      // 1. 메타 조회
      const metaRaw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(sourceRange)}')
        @{ Rows = $r.Rows.Count; Cols = $r.Columns.Count; StartRow = $r.Row; StartCol = $r.Column } | ConvertTo-Json -Compress
      `);
      const meta = parseJSON<{ Rows: number; Cols: number; StartRow: number; StartCol: number }>(metaRaw);
      const { Rows: rows, Cols: cols, StartRow: startRow, StartCol: startCol } = meta;

      // 2. 소스 읽기 (임시 파일 + 청크 분할)
      const data = rows < chunkSize
        ? await readSource(wbName, shName, `$ws.Range('${psEscape(sourceRange)}')`, rows, cols, prop)
        : await readSourceChunked(wbName, shName, rows, cols, startRow, startCol, prop, chunkSize);

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

// ── 소스 읽기: 단일 + 임시 파일 ──
async function readSource(
  wbName: string, shName: string, rangeExpr: string,
  rows: number, cols: number, prop: string
): Promise<(string | number | null)[][]> {
  const tmpPath = join(tmpdir(), `xlmcp_cpr_${randomUUID()}.json`);
  const escapedPath = tmpPath.replace(/\\/g, "\\\\");
  try {
    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $ws = Resolve-Sheet $wb ${shName}
      $r = ${rangeExpr}
      $values = $r.${prop}
      ${buildReadScript(rows, cols)}
      $json = ConvertTo-Json @($data) -Depth 5 -Compress
      [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
    `);
    return JSON.parse(readFileSync(tmpPath, "utf-8"));
  } finally {
    try { unlinkSync(tmpPath); } catch { /* ignore */ }
  }
}

// ── 소스 읽기: 청크 분할 + 병렬 ──
async function readSourceChunked(
  wbName: string, shName: string, rows: number, cols: number,
  startRow: number, startCol: number, prop: string, chunkSize: number
): Promise<(string | number | null)[][]> {
  const chunks: { offset: number; chunkRows: number }[] = [];
  for (let offset = 0; offset < rows; offset += chunkSize) {
    chunks.push({ offset, chunkRows: Math.min(chunkSize, rows - offset) });
  }
  const batchId = randomUUID();
  const tmpFiles = chunks.map((_, i) => join(tmpdir(), `xlmcp_cpr_${batchId}_${i}.json`));

  try {
    await Promise.all(
      chunks.map((chunk, i) => {
        const escapedPath = tmpFiles[i].replace(/\\/g, "\\\\");
        const r1 = startRow + chunk.offset;
        const r2 = r1 + chunk.chunkRows - 1;
        const c2 = startCol + cols - 1;
        return runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = $ws.Range($ws.Cells.Item(${r1}, ${startCol}), $ws.Cells.Item(${r2}, ${c2}))
          $values = $r.${prop}
          ${buildReadScript(chunk.chunkRows, cols)}
          $json = ConvertTo-Json @($data) -Depth 5 -Compress
          [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
        `);
      })
    );
    const allData: (string | number | null)[][] = [];
    for (const f of tmpFiles) {
      allData.push(...JSON.parse(readFileSync(f, "utf-8")));
    }
    return allData;
  } finally {
    for (const f of tmpFiles) {
      try { unlinkSync(f); } catch { /* ignore */ }
    }
  }
}

// ── PS 읽기 스크립트 (Rank 방어) ──
function buildReadScript(rows: number, cols: number): string {
  return `
      $data = @()
      if ($values -isnot [System.Array]) {
        $data = ,@(,$(if ($values -ne $null) { $values } else { $null }))
      } elseif ($values.Rank -eq 2) {
        for ($i = 1; $i -le ${rows}; $i++) {
          $row = @()
          for ($j = 1; $j -le ${cols}; $j++) {
            $v = $values[$i,$j]
            $row += $(if ($v -ne $null) { $v } else { $null })
          }
          $data += ,@($row)
        }
      } else {
        $row = @()
        for ($k = 0; $k -lt $values.Length; $k++) {
          $v = $values[$k]
          $row += $(if ($v -ne $null) { $v } else { $null })
        }
        if (${rows} -eq 1) {
          $data = ,@($row)
        } else {
          foreach ($v in $row) { $data += ,@(,$v) }
        }
      }`;
}
