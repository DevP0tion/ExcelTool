<#
  XLMCP resolver verification harness.

  Purpose: validate the attach-only / workbook-bearing / ghost-safe resolver on the
  ACTUAL machine where the bug occurs, in node-powershell's exact shell context
  (64-bit Windows PowerShell 5.1). tsc/build only check types -- THIS is the real
  behavioral verification, mapped to the bug report's AC7 acceptance criteria.

  Run (in the repo, on the bug machine):
      powershell -NoProfile -File plugins\xlmcp\scripts\verify-resolver.ps1
  Optional (tests create + ownership-quit; run with Excel CLOSED):
      powershell -NoProfile -File plugins\xlmcp\scripts\verify-resolver.ps1 -RunCreateTest

  SAFETY: read-only by default. It attaches to your running Excel to inspect it,
  never closing or quitting a user instance. With -RunCreateTest it may create ONE
  instance and quit it again ONLY IF the harness itself created it and it is empty.

  The block between "XLMCP RESOLVER BEGIN/END" is the exact PowerShell that ships in
  INIT_SCRIPT (src/services/powershell.ts). Keep the two in sync.
#>
param([switch]$RunCreateTest)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$ErrorActionPreference = 'Continue'

# ===================== XLMCP RESOLVER BEGIN (mirrors INIT_SCRIPT) =====================
$global:XlmcpOwnsExcel = $false
$global:excel = $null
$global:XlmcpExcelClsid = '{00024500-0000-0000-C000-000000000046}'

# ROT helper for cross-instance discovery. Best-effort: resolver degrades to
# GetActiveObject-only if Add-Type is unavailable in this shell.
$global:XlmcpRotReady = $false
try {
  if (-not ([System.Management.Automation.PSTypeName]'XlmcpRot').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public class XlmcpRot {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int r, out IRunningObjectTable t);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int r, out IBindCtx c);
  // Running objects whose moniker display name contains needle (case-insensitive).
  public static List<object> GetByFilter(string needle) {
    var res = new List<object>();
    IRunningObjectTable rot = null;
    if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return res;
    IEnumMoniker en = null; rot.EnumRunning(out en); if (en == null) return res; en.Reset();
    IMoniker[] m = new IMoniker[1];
    while (en.Next(1, m, IntPtr.Zero) == 0) {
      IBindCtx ctx = null; string dn = null;
      try {
        CreateBindCtx(0, out ctx);
        try { m[0].GetDisplayName(ctx, null, out dn); } catch {}
        if (dn != null && needle != null && dn.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) {
          object o = null;
          try { if (rot.GetObject(m[0], out o) == 0 && o != null) res.Add(o); } catch {}
        }
      } finally {
        if (ctx != null) Marshal.ReleaseComObject(ctx);
        if (m[0] != null) Marshal.ReleaseComObject(m[0]);
      }
    }
    Marshal.ReleaseComObject(en);
    Marshal.ReleaseComObject(rot);
    return res;
  }
  // Diagnostics only: moniker display names.
  public static List<string> List() {
    var res = new List<string>();
    IRunningObjectTable rot = null;
    if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return res;
    IEnumMoniker en = null; rot.EnumRunning(out en); if (en == null) return res; en.Reset();
    IMoniker[] m = new IMoniker[1];
    while (en.Next(1, m, IntPtr.Zero) == 0) {
      IBindCtx ctx = null; string dn = null;
      try { CreateBindCtx(0, out ctx); try { m[0].GetDisplayName(ctx, null, out dn); } catch {} res.Add(dn); }
      finally { if (ctx != null) Marshal.ReleaseComObject(ctx); if (m[0] != null) Marshal.ReleaseComObject(m[0]); }
    }
    Marshal.ReleaseComObject(en); Marshal.ReleaseComObject(rot);
    return res;
  }
}
'@
  }
  $global:XlmcpRotReady = $true
} catch { $global:XlmcpRotReady = $false }

function Get-XlmcpWbCount {
  param($app)
  if (-not $app) { return -1 }
  try { return [int]$app.Workbooks.Count } catch { return -1 }
}

function Test-XlmcpExcelRunning {
  # True if any Excel process exists. Lets callers tell "nothing open" from "running but unbindable".
  try { return (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) } catch { return $false }
}

function Find-WorkbookBearingExcel {
  # Excel.Application with >=1 workbook, or $null. NEVER creates.
  $candidates = New-Object System.Collections.ArrayList
  try {
    $a = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    if ($a) { [void]$candidates.Add($a) }
  } catch {}
  if ($global:XlmcpRotReady) {
    try { foreach ($a in [XlmcpRot]::GetByFilter($global:XlmcpExcelClsid)) { if ($a) { [void]$candidates.Add($a) } } } catch {}
    try { foreach ($wb in [XlmcpRot]::GetByFilter('.xls')) { try { if ($wb.Application) { [void]$candidates.Add($wb.Application) } } catch {} } } catch {}
  }
  # Most workbooks wins (first-wins on tie). Assumes a single user Excel instance
  # (the report's scenario); two equal-count instances could diverge across sessions.
  $best = $null; $bestCount = 0
  foreach ($c in $candidates) {
    $n = Get-XlmcpWbCount $c
    if ($n -gt $bestCount) { $bestCount = $n; $best = $c }
  }
  return $best
}

function Find-AnyExcel {
  # ANY running Excel.Application (even empty), or $null. NEVER creates.
  try {
    $a = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    if ($a) { return $a }
  } catch {}
  if ($global:XlmcpRotReady) {
    try { foreach ($a in [XlmcpRot]::GetByFilter($global:XlmcpExcelClsid)) { if ($a) { return $a } } } catch {}
    try { foreach ($wb in [XlmcpRot]::GetByFilter('.xls')) { try { if ($wb.Application) { return $wb.Application } } catch {} } } catch {}
  }
  return $null
}

function Ensure-Excel {
  param([switch]$AllowCreate)
  # Keep current binding while it still holds workbooks.
  if ($global:excel) {
    $n = Get-XlmcpWbCount $global:excel
    if ($n -gt 0) { return $global:excel }
    if ($n -lt 0) { $global:excel = $null }   # dead RCW -> drop and re-resolve
  }
  # Prefer the user's workbook-bearing instance.
  $wbApp = Find-WorkbookBearingExcel
  if ($wbApp) {
    $global:excel = $wbApp
    try { $global:excel.DisplayAlerts = $false } catch {}
    return $global:excel
  }
  if ($AllowCreate) {
    if (-not $global:excel) { $global:excel = Find-AnyExcel }   # attach before creating
    if (-not $global:excel) {
      $global:excel = New-Object -ComObject Excel.Application
      $global:excel.Visible = $true
      $global:XlmcpOwnsExcel = $true                            # only OUR creations are ownable
    }
    try { $global:excel.DisplayAlerts = $false } catch {}
    return $global:excel
  }
  return $null   # attach-only: nothing open -> caller emits a clean message
}

function Resolve-Workbook {
  param([string]$Name)
  $app = Ensure-Excel
  if (-not $app) {
    if (Test-XlmcpExcelRunning) {
      throw "Excel is running but XLMCP could not bind to a workbook-bearing instance (the workbook may not be registered in the COM Running Object Table). Try re-saving the workbook, or close and reopen it via excel_open_workbook."
    }
    throw "No workbook is open. Open a workbook in Excel first, or use excel_open_workbook / excel_create_workbook."
  }
  if ($Name -and $Name -ne "") { return $app.Workbooks.Item($Name) }
  if ($app.ActiveWorkbook) { return $app.ActiveWorkbook }
  if ((Get-XlmcpWbCount $app) -gt 0) { return $app.Workbooks.Item(1) }
  throw "No workbook is open."
}

function Resolve-Sheet {
  param($wb, [string]$SheetName)
  if ($SheetName -and $SheetName -ne "") { return $wb.Worksheets.Item($SheetName) }
  return $wb.ActiveSheet
}
# Disposal (mirrors Session.dispose): quit ONLY harness/XLMCP-created empty instances.
function Invoke-XlmcpDispose {
  if ($global:excel) {
    if ($global:XlmcpOwnsExcel) {
      try { if ((Get-XlmcpWbCount $global:excel) -eq 0) { $global:excel.Quit() } } catch {}
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:excel) | Out-Null } catch {}
    $global:excel = $null
  }
}
# ====================== XLMCP RESOLVER END (mirrors INIT_SCRIPT) ======================

# ----------------------------- test runner -----------------------------
$script:pass = 0; $script:fail = 0; $script:info = 0
function Show-Result {
  param([string]$name, [string]$status, [string]$detail = "")
  $tag = switch ($status) { 'PASS' { '[PASS]' } 'FAIL' { '[FAIL]' } default { '[INFO]' } }
  if ($status -eq 'PASS') { $script:pass++ } elseif ($status -eq 'FAIL') { $script:fail++ } else { $script:info++ }
  if ($detail) { Write-Output ("{0} {1} -- {2}" -f $tag, $name, $detail) } else { Write-Output ("{0} {1}" -f $tag, $name) }
}

Write-Output "==================== XLMCP resolver verification ===================="
Write-Output ("PSVersion={0}  PSEdition={1}  Is64BitProcess={2}" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [Environment]::Is64BitProcess)
if ($env:NODE_POWERSHELL) { Write-Output ("NODE_POWERSHELL={0} (server may use a different shell than this run)" -f $env:NODE_POWERSHELL) }
Write-Output ("EXCEL processes now: {0}" -f @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count)
Write-Output "---------------------------------------------------------------------"

# A) Add-Type / ROT helper available in this shell (Tier-2 mechanism viability)
if ($global:XlmcpRotReady) { Show-Result "A. Add-Type CSharp ROT helper compiles in this shell" "PASS" }
else { Show-Result "A. Add-Type CSharp ROT helper compiles in this shell" "FAIL" "resolver will run GetActiveObject-only (degraded)" }

# B) ROT enumeration (diagnostic)
if ($global:XlmcpRotReady) {
  try {
    $names = @([XlmcpRot]::List())
    Show-Result "B. ROT enumeration" "INFO" ("{0} entries" -f $names.Count)
    foreach ($n in $names) { if ($n) { Write-Output ("       ROT: " + $n) } }
  } catch { Show-Result "B. ROT enumeration" "FAIL" $_.Exception.Message }
}

# C) CORE: select the workbook-bearing instance (report AC7.4) -- read-only
$wbApp = Find-WorkbookBearingExcel
if ($wbApp) {
  $cnt = Get-XlmcpWbCount $wbApp
  $list = @()
  try { foreach ($wb in $wbApp.Workbooks) { try { $list += $wb.FullName } catch {} } } catch {}
  Show-Result "C. Find-WorkbookBearingExcel selects a workbook-bearing instance" "PASS" ("Workbooks.Count={0}" -f $cnt)
  foreach ($f in $list) { Write-Output ("       WB: " + $f) }
} else {
  Show-Result "C. Find-WorkbookBearingExcel selects a workbook-bearing instance" "INFO" "no open workbook found right now -- open a workbook in Excel and re-run to verify AC7.1/AC7.4"
}

# D) Resolve-Workbook end-to-end (the path the failing tools use) -- read-only
try {
  $wb = Resolve-Workbook ""
  Show-Result "D. Resolve-Workbook (active)" "PASS" ("resolved: {0}" -f $wb.FullName)
} catch {
  if ($wbApp) { Show-Result "D. Resolve-Workbook (active)" "FAIL" $_.Exception.Message }
  else { Show-Result "D. Resolve-Workbook (active)" "INFO" ("clean message when nothing open: " + $_.Exception.Message) }
}

# E) attach-only never auto-creates, and emits a clean message when nothing open (Q2 policy)
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
$null = Ensure-Excel                      # attach-only
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
if ($after -le $before) { Show-Result "E. attach-only Ensure-Excel creates no process" "PASS" ("EXCEL {0} -> {1}" -f $before, $after) }
else { Show-Result "E. attach-only Ensure-Excel creates no process" "FAIL" ("EXCEL {0} -> {1} (a process was spawned!)" -f $before, $after) }

# F) opt-in: create + ownership-quit path (run with Excel CLOSED to exercise creation)
if ($RunCreateTest) {
  $b2 = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
  $ownedBefore = $global:XlmcpOwnsExcel
  $app = Ensure-Excel -AllowCreate
  if ($app) {
    if ($global:XlmcpOwnsExcel -and -not $ownedBefore) {
      Show-Result "F. Ensure-Excel -AllowCreate created an instance we own" "PASS" ("owns=true; Workbooks.Count={0}" -f (Get-XlmcpWbCount $app))
      Invoke-XlmcpDispose                 # should Quit (owned + empty)
      Start-Sleep -Milliseconds 600
      $a2 = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
      if ($a2 -le $b2) { Show-Result "F. dispose quits the owned empty instance (no ghost)" "PASS" ("EXCEL {0} -> {1}" -f $b2, $a2) }
      else { Show-Result "F. dispose quits the owned empty instance (no ghost)" "FAIL" ("EXCEL {0} -> {1}" -f $b2, $a2) }
    } else {
      Show-Result "F. Ensure-Excel -AllowCreate attached to an existing instance" "INFO" "owns=false (a user instance was already running; create path not exercised -- close Excel and re-run with -RunCreateTest)"
    }
  } else {
    Show-Result "F. Ensure-Excel -AllowCreate" "FAIL" "returned null (could not attach or create -- Excel may not be installed/registered on this machine)"
  }
} else {
  Show-Result "F. create + ownership-quit path" "INFO" "skipped (pass -RunCreateTest, ideally with Excel closed)"
}

Write-Output "---------------------------------------------------------------------"
Write-Output ("RESULT: PASS={0} FAIL={1} INFO={2}" -f $script:pass, $script:fail, $script:info)
Write-Output ""
Write-Output "Manual checks not automatable here:"
Write-Output "  AC7.1  Start server with Excel CLOSED, then open a workbook -> excel_list_open_workbooks must return it."
Write-Output "  AC7.2  Restart the MCP server several times -> 'Get-Process EXCEL' count must NOT grow monotonically."
Write-Output "  AC7.3  After server shutdown, a workbook YOU opened must still be open (never Quit)."
Write-Output "===================================================================="
