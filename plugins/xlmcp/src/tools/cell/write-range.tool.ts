import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { tmpdir } from "os";
import { writeFileSync, unlinkSync } from "fs";
import { join } from "path";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const CHUNK_THRESHOLD = 500; // 이 행수 이상이면 청크 분할
const CHUNK_SIZE = 500;

interface FormulaEntry {
  rowOffset: number;
  colOffset: number;
  formula: string;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_range",
    {
      title: "범위 쓰기",
      description:
        "시작 셀부터 2D 배열 데이터를 입력합니다. 대용량 데이터는 자동 청크 분할 + 병렬 쓰기로 가속됩니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        startCell: z.string().describe("시작 셀 주소 (예: A1)"),
        data: z
          .array(z.array(z.string()))
          .describe("2D 배열 데이터. 각 내부 배열이 한 행"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, startCell, data }) => {
      const rows = data.length;
      const cols = data[0]?.length ?? 0;
      if (rows === 0 || cols === 0) {
        return textContent({ success: true, rows: 0, cols: 0 });
      }

      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // 소규모: 기존 인라인 방식
      if (rows < CHUNK_THRESHOLD) {
        return writeInline(wbName, shName, startCell, data, rows, cols);
      }

      // 대규모: 청크 분할 + 임시 파일 + 병렬 쓰기
      return writeChunked(wbName, shName, startCell, data, rows, cols);
    }
  );
}

// ── 소규모: 인라인 (기존 방식) ──
async function writeInline(
  wbName: string,
  shName: string,
  startCell: string,
  data: string[][],
  rows: number,
  cols: number
) {
  const formulas: FormulaEntry[] = [];
  const psRows = data
    .map((row, ri) => {
      const cells = row.map((v, ci) => {
        if (v.startsWith("=")) {
          formulas.push({ rowOffset: ri, colOffset: ci, formula: v });
          return "''";
        }
        return `'${psEscape(v)}'`;
      });
      return `@(${cells.join(",")})`;
    })
    .join(",");

  const formulaCmds = formulas
    .map(
      (f) =>
        `$ws.Cells.Item($start.Row + ${f.rowOffset}, $start.Column + ${f.colOffset}).Formula = '${psEscape(f.formula)}'`
    )
    .join("\n        ");

  await runPS(`
    $wb = Resolve-Workbook ${wbName}
    $ws = Resolve-Sheet $wb ${shName}
    $start = $ws.Range('${psEscape(startCell)}')
    $endRow = $start.Row + ${rows} - 1
    $endCol = $start.Column + ${cols} - 1
    $targetRange = $ws.Range($start, $ws.Cells.Item($endRow, $endCol))
    $arr = New-Object 'object[,]' ${rows},${cols}
    $srcData = @(${psRows})
    for ($i = 0; $i -lt ${rows}; $i++) {
      $row = @($srcData[$i])
      for ($j = 0; $j -lt ${cols}; $j++) {
        $val = $row[$j]
        $num = 0.0
        if ([double]::TryParse($val, [ref]$num)) {
          $arr[$i,$j] = $num
        } else {
          $arr[$i,$j] = $val
        }
      }
    }
    $targetRange.Value2 = $arr
    ${formulaCmds}
  `);

  return textContent({ success: true, rows, cols });
}

// ── 대규모: 청크 분할 + 임시 파일 + 병렬 ──
async function writeChunked(
  wbName: string,
  shName: string,
  startCell: string,
  data: string[][],
  rows: number,
  cols: number
) {
  // 청크 분할
  const chunks: { rowOffset: number; chunkData: (string | number | null)[][] }[] = [];
  const formulas: FormulaEntry[] = [];

  for (let offset = 0; offset < rows; offset += CHUNK_SIZE) {
    const end = Math.min(offset + CHUNK_SIZE, rows);
    const chunkData: (string | number | null)[][] = [];

    for (let ri = offset; ri < end; ri++) {
      const row: (string | number | null)[] = [];
      for (let ci = 0; ci < cols; ci++) {
        const v = data[ri][ci];
        if (v.startsWith("=")) {
          formulas.push({ rowOffset: ri, colOffset: ci, formula: v });
          row.push(null);
        } else {
          const num = Number(v);
          row.push(v !== "" && !isNaN(num) ? num : v);
        }
      }
      chunkData.push(row);
    }

    chunks.push({ rowOffset: offset, chunkData });
  }

  // 임시 파일 생성
  const tmpFiles: string[] = [];
  const timestamp = Date.now();

  for (let i = 0; i < chunks.length; i++) {
    const filePath = join(tmpdir(), `xlmcp_chunk_${timestamp}_${i}.json`);
    writeFileSync(filePath, JSON.stringify(chunks[i].chunkData));
    tmpFiles.push(filePath);
  }

  try {
    // 병렬 쓰기 (General Pool 라운드 로빈)
    await Promise.all(
      chunks.map((chunk, i) => {
        const chunkRows = chunk.chunkData.length;
        const tmpPath = tmpFiles[i].replace(/\\/g, "\\\\");

        return runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $start = $ws.Range('${psEscape(startCell)}')
          $chunkStart = $ws.Cells.Item($start.Row + ${chunk.rowOffset}, $start.Column)
          $chunkEnd = $ws.Cells.Item($start.Row + ${chunk.rowOffset} + ${chunkRows} - 1, $start.Column + ${cols} - 1)
          $targetRange = $ws.Range($chunkStart, $chunkEnd)

          $json = Get-Content '${tmpPath}' -Raw -Encoding UTF8
          $data = $json | ConvertFrom-Json
          $arr = New-Object 'object[,]' ${chunkRows},${cols}
          for ($i = 0; $i -lt ${chunkRows}; $i++) {
            for ($j = 0; $j -lt ${cols}; $j++) {
              $v = $data[$i][$j]
              if ($v -ne $null) { $arr[$i,$j] = $v }
            }
          }
          $targetRange.Value2 = $arr
        `);
      })
    );

    // 수식 적용 (청크 완료 후 순차)
    if (formulas.length > 0) {
      const formulaCmds = formulas
        .map(
          (f) =>
            `$ws.Cells.Item($start.Row + ${f.rowOffset}, $start.Column + ${f.colOffset}).Formula = '${psEscape(f.formula)}'`
        )
        .join("\n        ");

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $start = $ws.Range('${psEscape(startCell)}')
        ${formulaCmds}
      `);
    }
  } finally {
    // 임시 파일 정리
    for (const f of tmpFiles) {
      try {
        unlinkSync(f);
      } catch {
        // ignore
      }
    }
  }

  return textContent({
    success: true,
    rows,
    cols,
    chunks: chunks.length,
    mode: "parallel",
  });
}
