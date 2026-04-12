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
    "excel_read_range_formulas",
    {
      title: "범위 수식 읽기",
      description:
        "셀 범위의 수식을 2D 배열로 반환합니다. 수식이 없는 셀은 값을 반환합니다. 대용량 시 자동 청크 분할 병렬 읽기.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().optional().describe("범위 주소 (예: A1:C10). 생략 시 UsedRange"),
        chunkSize: z.number().int().optional().describe("청크 분할 행수. 기본 30"),
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

      // 메타 조회
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
        Rows: number; Cols: number; StartRow: number; StartCol: number; Address: string;
      }>(metaRaw);

      const { Rows: rows, Cols: cols, StartRow: startRow, StartCol: startCol, Address: addr } = meta;

      if (rows === 0 || cols === 0) {
        return textContent({ Range: addr, Rows: 0, Cols: 0, Data: [] });
      }

      if (rows < chunkSize) {
        const data = await readFormulaSingle(wbName, shName, rangeExpr, rows, cols);
        return textContent({ Range: addr, Rows: rows, Cols: cols, Data: data });
      }

      const data = await readFormulaChunked(wbName, shName, rows, cols, startRow, startCol, chunkSize);
      return textContent({ Range: addr, Rows: rows, Cols: cols, Data: data });
    }
  );
}

async function readFormulaSingle(
  wbName: string, shName: string, rangeExpr: string, rows: number, cols: number
): Promise<unknown[][]> {
  const tmpPath = join(tmpdir(), `xlmcp_readf_${randomUUID()}.json`);
  const escapedPath = tmpPath.replace(/\\/g, "\\\\");
  try {
    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $ws = Resolve-Sheet $wb ${shName}
      $r = ${rangeExpr}
      $values = $r.Formula
      ${buildReadScript(rows, cols)}
      $json = ConvertTo-Json @($data) -Depth 5 -Compress
      [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
    `);
    return JSON.parse(readFileSync(tmpPath, "utf-8"));
  } finally {
    try { unlinkSync(tmpPath); } catch { /* ignore */ }
  }
}

async function readFormulaChunked(
  wbName: string, shName: string, rows: number, cols: number,
  startRow: number, startCol: number, chunkSize: number
): Promise<unknown[][]> {
  const chunks: { offset: number; chunkRows: number }[] = [];
  for (let offset = 0; offset < rows; offset += chunkSize) {
    chunks.push({ offset, chunkRows: Math.min(chunkSize, rows - offset) });
  }

  const batchId = randomUUID();
  const tmpFiles = chunks.map((_, i) => join(tmpdir(), `xlmcp_readf_${batchId}_${i}.json`));

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
          $values = $r.Formula
          ${buildReadScript(chunk.chunkRows, cols)}
          $json = ConvertTo-Json @($data) -Depth 5 -Compress
          [System.IO.File]::WriteAllText('${escapedPath}', $json, (New-Object System.Text.UTF8Encoding $false))
        `);
      })
    );

    const allData: unknown[][] = [];
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
