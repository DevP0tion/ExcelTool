import { PowerShell } from "node-powershell";

let shell: PowerShell | null = null;

export async function getShell(): Promise<PowerShell> {
  if (!shell) {
    shell = new PowerShell({
      executableOptions: {
        "-ExecutionPolicy": "Bypass",
        "-NoProfile": true,
      },
    });
    // 실행 중인 Excel에 연결 시도, 없으면 새로 생성
    await shell.invoke(`
      try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
      } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
      }
      $excel.DisplayAlerts = $false

      function Resolve-Workbook {
        param([string]$Name)
        if ($Name -and $Name -ne "") {
          return $excel.Workbooks.Item($Name)
        }
        if (-not $excel.ActiveWorkbook) {
          throw "열려 있는 워크북이 없습니다."
        }
        return $excel.ActiveWorkbook
      }

      function Resolve-Sheet {
        param($wb, [string]$SheetName)
        if ($SheetName -and $SheetName -ne "") {
          return $wb.Worksheets.Item($SheetName)
        }
        return $wb.ActiveSheet
      }
    `);
  }
  return shell;
}

export async function runPS(script: string): Promise<string> {
  const ps = await getShell();
  const wrapped = `
    try {
      ${script}
    } catch {
      [Console]::Error.WriteLine(($_ | ConvertTo-Json -Compress))
      throw $_
    }
  `;
  try {
    const result = await ps.invoke(wrapped);
    return result.raw ?? "";
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    // PowerShell InvocationError에서 구조화된 메시지 추출 시도
    try {
      const parsed = JSON.parse(msg);
      throw new Error(JSON.stringify({
        error: true,
        message: parsed.Exception?.Message ?? parsed.FullyQualifiedErrorId ?? msg,
        type: parsed.Exception?.GetType?.()?.Name ?? "PowerShellError",
      }));
    } catch {
      throw new Error(JSON.stringify({
        error: true,
        message: msg.replace(/\r?\n/g, " ").trim(),
        type: "PowerShellError",
      }));
    }
  }
}

export async function dispose(): Promise<void> {
  if (shell) {
    try {
      await shell.invoke(`
        if ($excel) {
          [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
      `);
    } catch { /* ignore */ }
    await shell.dispose();
    shell = null;
  }
}
