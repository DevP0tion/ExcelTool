import Shell from "node-powershell";

let shell: InstanceType<typeof Shell> | null = null;

export async function getShell(): Promise<InstanceType<typeof Shell>> {
  if (!shell) {
    shell = new Shell({
      executionPolicy: "Bypass",
      noProfile: true,
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
  const result = await ps.invoke(script);
  return result.raw ?? "";
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
    shell.dispose();
    shell = null;
  }
}
