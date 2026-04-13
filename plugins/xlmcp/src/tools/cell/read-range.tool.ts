import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { readFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const DEFAULT_CHUNK_SIZE = 30;

export function register(server: McpServer) {
  server.registerTool(
    "excel_read_range",
    {
      title: "Read Range",
      description:
        "Read cell values as 2D array. Reads UsedRange if omitted. Auto chunked parallel for large data.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().optional().describe("Range address (e.g. A1:C10). UsedRange if omitted"),
        chunkSize: z.number().int().optional().describe("Chunk size in rows for parallel read. Default 30"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, range, chunkSize: cs }) => {
      const chunkSize = cs ?? DEFAULT_CHUNK_SIZE;
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range
        ? `$ws.Range('${psEscape(range)}')`
        : `$ws.UsedRange`;

      // 1. 범위 메타 조회
      const metaRaw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = ${rangeExpr}
        @{
          Rows = $r.Rows.Count
          Cols = $r.Columns.Count
          StartRow = $r.Row
          StartCol = $r.Column
          Address = $r.Address()
        } | ConvertTo-Json -Compress
      `);
      const meta = parseJSON<{
        Rows: number;
        Cols: number;
        StartRow: number;
        StartCol: number;
        Address: string;
      }>(metaRaw);

      const { Rows: rows, Cols: cols, StartRow: startRow, StartCol: startCol, Address: addr } = meta;

      if (rows === 0 || cols === 0) {
        return textContent({ Range: addr, Rows: 0, Cols: 0, Data: [] });
      }

      // 2. 소규모: 단일 읽기 + 임시 파일 출력
      if (rows < chunkSize) {
        const data = await readSingle(wbName, shName, rangeExpr, rows, cols);
        return textContent({ Range: addr, Rows: rows, Cols: cols, Data: data });
      }

      // 3. 대규모: 청크 분할 병렬 읽기
      const data = await readChunked(wbName, shName, rows, cols, startRow, startCol, chunkSize);
      return textContent({ Range: addr, Rows: rows, Cols: cols, Data: data });
    }
  );
}

// ── 소규모: 단일 읽기 + 임시 파일 ──
async function readSingle(
  wbName: string,
  shName: string,
  rangeExpr: string,
  rows: number,
  cols: number
): Promise<unknown[][]> {
  const tmpPath = join(tmpdir(), `xlmcp_read_${randomUUID()}.json`);
  const escapedPath = tmpPath.replace(/\\/g, "\\\\");

  try {
    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $ws = Resolve-Sheet $wb ${shName}
      $r = ${rangeExpr}
      $values = $r.Value2
      ${buildReadScript(rows, cols)}
      $json = ConvertTo-Json @($data) -Depth 5 -Compress
      [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
    `);
    return JSON.parse(readFileSync(tmpPath, "utf-8"));
  } finally {
    try { unlinkSync(tmpPath); } catch { /* ignore */ }
  }
}

// ── 대규모: 청크 분할 병렬 읽기 ──
async function readChunked(
  wbName: string,
  shName: string,
  rows: number,
  cols: number,
  startRow: number,
  startCol: number,
  chunkSize: number
): Promise<unknown[][]> {
  // 청크 정보 생성
  const chunks: { offset: number; chunkRows: number }[] = [];
  for (let offset = 0; offset < rows; offset += chunkSize) {
    chunks.push({ offset, chunkRows: Math.min(chunkSize, rows - offset) });
  }

  // 임시 파일 경로 생성
  const batchId = randomUUID();
  const tmpFiles = chunks.map((_, i) =>
    join(tmpdir(), `xlmcp_read_${batchId}_${i}.json`)
  );

  try {
    // 병렬 읽기
    await Promise.all(
      chunks.map((chunk, i) => {
        const escapedPath = tmpFiles[i].replace(/\\/g, "\\\\");
        const chunkStartRow = startRow + chunk.offset;
        const chunkEndRow = chunkStartRow + chunk.chunkRows - 1;
        const endCol = startCol + cols - 1;

        return runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = $ws.Range($ws.Cells.Item(${chunkStartRow}, ${startCol}), $ws.Cells.Item(${chunkEndRow}, ${endCol}))
          $values = $r.Value2
          ${buildReadScript(chunk.chunkRows, cols)}
          $json = ConvertTo-Json @($data) -Depth 5 -Compress
          [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
        `);
      })
    );

    // TS에서 파일 병합
    const allData: unknown[][] = [];
    for (const f of tmpFiles) {
      const chunkData: unknown[][] = JSON.parse(readFileSync(f, "utf-8"));
      allData.push(...chunkData);
    }
    return allData;
  } finally {
    for (const f of tmpFiles) {
      try { unlinkSync(f); } catch { /* ignore */ }
    }
  }
}

// ── PS 읽기 스크립트 생성 (배열 Rank 방어) ──
function buildReadScript(rows: number, cols: number): string {
  return `
      $data = @()
      if ($values -isnot [System.Array]) {
        # 스칼라 (1행 1열)
        $data = ,@(,$(if ($values -ne $null) { $values } else { $null }))
      } elseif ($values.Rank -eq 2) {
        # 2D 배열 (정상 경로)
        for ($i = 1; $i -le ${rows}; $i++) {
          $row = @()
          for ($j = 1; $j -le ${cols}; $j++) {
            $v = $values[$i,$j]
            $row += $(if ($v -ne $null) { $v } else { $null })
          }
          $data += ,@($row)
        }
      } else {
        # 1D 배열 (1행 또는 1열이 언래핑된 경우)
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
