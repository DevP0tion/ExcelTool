import Shell from "node-powershell";

let shell: InstanceType<typeof Shell> | null = null;

export async function getShell(): Promise<InstanceType<typeof Shell>> {
  if (!shell) {
    shell = new Shell({
      executionPolicy: "Bypass",
      noProfile: true,
    });
    // Excel COM 초기화
    await shell.invoke(`
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $false
      $excel.DisplayAlerts = $false
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
          $excel.Quit()
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
      `);
    } catch { /* ignore */ }
    shell.dispose();
    shell = null;
  }
}
