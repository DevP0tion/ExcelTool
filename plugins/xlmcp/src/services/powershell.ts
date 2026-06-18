import { PowerShell } from "node-powershell";
import { env } from "./env.js";

// ── 설정 ──
const POOL_SIZE = Math.max(1, parseInt(env("POOL_SIZE") ?? "4", 10) || 4);
const HEARTBEAT_INTERVAL = 10_000;
const INVOKE_TIMEOUT = 30_000;

// ── PS 초기화 스크립트 ──
// attach-only / 워크북 보유 인스턴스 우선 / 유령-안전 리졸버.
// INIT은 인스턴스를 생성하지 않는다(지연 바인딩). 첫 도구 호출의 invoke 래퍼가 Ensure-Excel을 돌린다.
// 이 블록은 scripts/verify-resolver.ps1 의 "XLMCP RESOLVER" 블록과 동일하게 유지할 것.
// 주의: C# here-string 의 종결자 '@ 와 본문은 반드시 column 0 에서 시작해야 한다(PowerShell here-string 규칙).
const INIT_SCRIPT = `
$global:XlmcpOwnsExcel = $false
$global:excel = $null
$global:XlmcpExcelClsid = '{00024500-0000-0000-C000-000000000046}'

# ROT helper for cross-instance discovery. Best-effort: degrades to GetActiveObject-only if Add-Type is unavailable.
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
  # Keep current binding while it still holds workbooks (fast steady-state path).
  if ($global:excel) {
    $n = Get-XlmcpWbCount $global:excel
    if ($n -gt 0) { return $global:excel }
    if ($n -lt 0) { $global:excel = $null }
  }
  # Prefer the user's workbook-bearing instance.
  $wbApp = Find-WorkbookBearingExcel
  if ($wbApp) {
    $global:excel = $wbApp
    try { $global:excel.DisplayAlerts = $false } catch {}
    return $global:excel
  }
  if ($AllowCreate) {
    if (-not $global:excel) { $global:excel = Find-AnyExcel }   # attach to an existing instance before creating
    if (-not $global:excel) {
      $global:excel = New-Object -ComObject Excel.Application
      $global:excel.Visible = $true
      $global:XlmcpOwnsExcel = $true                            # only OUR creations are ownable
    }
    try { $global:excel.DisplayAlerts = $false } catch {}
    return $global:excel
  }
  return $null   # attach-only: nothing open -> caller throws a clean message
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

function Invoke-XlmcpDispose {
  # Quit ONLY XLMCP-created (owned) empty instances. NEVER quit a user instance (preserve data).
  if ($global:excel) {
    if ($global:XlmcpOwnsExcel) {
      try { if ((Get-XlmcpWbCount $global:excel) -eq 0) { $global:excel.Quit() } } catch {}
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:excel) | Out-Null } catch {}
    $global:excel = $null
  }
}
`;

// ── 개별 세션 ──
class Session {
  public ps: PowerShell;
  public busy = false;
  public alive = true;

  constructor(public readonly id: number) {
    this.ps = new PowerShell({
      executableOptions: {
        "-ExecutionPolicy": "Bypass",
        "-NoProfile": true,
      },
    });
  }

  async init(): Promise<void> {
    await this.ps.invoke(INIT_SCRIPT);
  }

  async invoke(script: string, timeoutMs: number): Promise<string> {
    this.busy = true;
    const wrapped = `
      try {
        # Re-bind to the workbook-bearing instance every call so all sessions converge on the user's instance.
        # attach-only: if nothing is open, \$global:excel stays \$null and each tool handles it.
        Ensure-Excel | Out-Null
        ${script}
      } catch {
        try { [Console]::Error.WriteLine(($_ | ConvertTo-Json -Compress)) } catch { [Console]::Error.WriteLine($_.Exception.Message) }
        throw $_
      } finally {
        # COM 참조 정리 (변수가 존재하는 경우만)
        foreach ($__v in @('r','c','src','srcWs','dstWs','dst','dest','targetRange','start','chunkStart','chunkEnd','pos','t','pvt','pf','cache','chart','chartObj','fc','first','current','n','existing','pic','s','comp','cm')) {
          try {
            $__obj = Get-Variable -Name $__v -ValueOnly -ErrorAction SilentlyContinue
            if ($__obj -ne $null) {
              [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($__obj)
              Set-Variable -Name $__v -Value $null -ErrorAction SilentlyContinue
            }
          } catch {}
        }
      }
    `;
    try {
      const result = await Session.withTimeout(this.ps.invoke(wrapped), timeoutMs);
      return result.raw ?? "";
    } finally {
      this.busy = false;
    }
  }

  async healthCheck(): Promise<boolean> {
    if (this.busy || !this.alive) return this.alive;
    try {
      // PS 프로세스 생존만 확인 (Excel 비의존).
      // attach-only 지연 바인딩에서 $excel 은 워크북 작업 전까지 $null 이므로
      // $excel.Version 로 검사하면 워크북 미오픈 상태를 "죽음"으로 오판해 세션이 끝없이 재생성된다.
      await Session.withTimeout(this.ps.invoke("$PID"), 5000);
      return true;
    } catch {
      this.alive = false;
      return false;
    }
  }

  async dispose(): Promise<void> {
    try {
      // 정리 invoke에도 타임아웃 적용 (블로킹 PS 방어)
      // 소유권 기반: 우리가 만든 빈 인스턴스만 Quit. 사용자 인스턴스는 보존(데이터 유실 방지).
      await Session.withTimeout(
        this.ps.invoke("if (Get-Command Invoke-XlmcpDispose -ErrorAction SilentlyContinue) { Invoke-XlmcpDispose }"),
        5000
      );
    } catch { /* ignore */ }
    try {
      await this.ps.dispose();
    } catch { /* ignore */ }
    this.alive = false;
  }

  isProcessDead(err: unknown): boolean {
    const msg = err instanceof Error ? err.message : String(err);
    return (
      msg.includes("process exited") ||
      msg.includes("invoke called after") ||
      msg.includes("EPIPE") ||
      msg.includes("Timeout")
    );
  }

  static async create(id: number): Promise<Session> {
    const session = new Session(id);
    try {
      await session.init();
    } catch (err) {
      // INIT_SCRIPT 실패 시 PS 프로세스 정리
      session.alive = false;
      try { await session.ps.dispose(); } catch { /* ignore */ }
      throw err;
    }
    return session;
  }

  private static withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error(`Timeout: ${ms}ms exceeded`)), ms);
      promise.then(
        (v) => { clearTimeout(timer); resolve(v); },
        (e) => { clearTimeout(timer); reject(e); }
      );
    });
  }
}

// ── 작업 큐 항목 ──
interface QueuedTask {
  id: number;
  script: string;
  resolve: (v: string) => void;
  reject: (e: Error) => void;
  enqueuedAt: number;
}

// ── 세션 풀 ──
class SessionPool {
  private generalPool: Session[] = [];
  private exclusiveSession: Session | null = null;
  private roundRobinIndex = 0;
  private initialized = false;
  private nextGeneralId = 0;
  private pendingCreations = 0; // 생성 중인 세션 수 (경합 방지)
  private nextTaskId = 1;

  // exclusive
  private exclusiveRunning = false;
  private exclusiveQueue: Array<{
    script: string;
    resolve: (v: string) => void;
    reject: (e: Error) => void;
  }> = [];
  private generalActiveCount = 0;
  private generalDrainResolve: (() => void) | null = null;

  // 이벤트 기반 대기 (폴링 제거)
  private exclusiveEndWaiters: (() => void)[] = [];
  private generalQuietWaiters: (() => void)[] = [];

  // 작업 큐
  private generalQueue: QueuedTask[] = [];
  private totalProcessed = 0;
  private totalQueued = 0;

  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  // ── 초기화 (1개만 생성, 실패 시 정리) ──
  async init(): Promise<void> {
    if (this.initialized) return;

    let general: Session | null = null;
    let exclusive: Session | null = null;

    try {
      general = await Session.create(this.nextGeneralId++);
      exclusive = await Session.create(100);
    } catch (err) {
      // 부분 성공 세션 정리
      if (general) {
        this.nextGeneralId--;
        await general.dispose();
      }
      if (exclusive) await exclusive.dispose();
      throw err;
    }

    this.generalPool = [general];
    this.exclusiveSession = exclusive;
    this.heartbeatTimer = setInterval(() => this.heartbeat(), HEARTBEAT_INTERVAL);
    this.initialized = true;
  }

  // ── 일반 실행 ──
  async executeGeneral(script: string): Promise<string> {
    await this.init();
    if (this.exclusiveRunning) {
      await this.waitForExclusiveEnd();
    }

    // 유휴 세션 탐색
    const idle = this.findIdle();
    if (idle) {
      return this.invokeOnSession(idle, script, false);
    }

    // 상한 미도달 → 새 세션 생성 (pendingCreations로 동시 생성 경합 방지)
    if (this.generalPool.length + this.pendingCreations < POOL_SIZE) {
      this.pendingCreations++;
      try {
        const newSession = await Session.create(this.nextGeneralId++);
        this.generalPool.push(newSession);
        return this.invokeOnSession(newSession, script, false);
      } catch (err) {
        this.nextGeneralId--;
        throw err;
      } finally {
        this.pendingCreations--;
      }
    }

    // 상한 도달 → 큐에 대기
    this.totalQueued++;
    return new Promise<string>((resolve, reject) => {
      this.generalQueue.push({ id: this.nextTaskId++, script, resolve, reject, enqueuedAt: Date.now() });
    });
  }

  // ── exclusive 실행 ──
  async executeExclusive(script: string): Promise<string> {
    await this.init();
    if (this.exclusiveRunning) {
      return new Promise<string>((resolve, reject) => {
        this.exclusiveQueue.push({ script, resolve, reject });
      });
    }
    return this.runExclusive(script);
  }

  private async runExclusive(script: string): Promise<string> {
    this.exclusiveRunning = true;
    if (this.generalActiveCount > 0) {
      await new Promise<void>((resolve) => {
        this.generalDrainResolve = resolve;
      });
    }
    try {
      return await this.invokeOnSession(this.exclusiveSession!, script, true);
    } finally {
      const next = this.exclusiveQueue.shift();

      if (next && this.generalQueue.length > 0) {
        // general 큐 우선: exclusive 해제 → general flush → 완료 대기 → exclusive 재개
        this.exclusiveRunning = false;
        this.signalExclusiveEnd();
        this.flushGeneralQueue();
        this.waitForGeneralQuiet().then(() => {
          this.runExclusive(next.script).then(next.resolve, next.reject);
        });
      } else if (next) {
        // general 큐 없음: 바로 다음 exclusive 실행
        this.runExclusive(next.script).then(next.resolve, next.reject);
      } else {
        this.exclusiveRunning = false;
        this.signalExclusiveEnd();
        this.flushGeneralQueue();
      }
    }
  }

  // ── 세션에서 실행 (사망 시 1회 재시도) ──
  private async invokeOnSession(
    session: Session,
    script: string,
    isExclusive: boolean
  ): Promise<string> {
    if (!isExclusive) this.generalActiveCount++;
    try {
      const result = await session.invoke(script, INVOKE_TIMEOUT);
      this.totalProcessed++;
      return result;
    } catch (err: unknown) {
      if (session.isProcessDead(err)) {
        await this.recoverSession(session, isExclusive);
        // 복구된 세션으로 1회 재시도
        const recovered = isExclusive
          ? this.exclusiveSession
          : this.generalPool.find((s) => s.id === session.id);
        if (recovered && recovered.alive) {
          try {
            const result = await recovered.invoke(script, INVOKE_TIMEOUT);
            this.totalProcessed++;
            return result;
          } catch {
            // 재시도도 실패 → 원본 에러 throw
          }
        }
      }
      throw SessionPool.formatError(err);
    } finally {
      if (!isExclusive) {
        this.generalActiveCount--;
        if (this.exclusiveRunning && this.generalActiveCount === 0 && this.generalDrainResolve) {
          this.generalDrainResolve();
          this.generalDrainResolve = null;
        }
        // 큐에서 다음 작업 디스패치 (복구된 세션이면 풀에서 탐색)
        const dispatchTarget = session.alive
          ? session
          : this.generalPool.find((s) => s.id === session.id) ?? session;
        this.dispatchFromQueue(dispatchTarget);
        // general quiet 시그널 (큐 비고 활성 0이면)
        this.signalGeneralQuiet();
      }
    }
  }

  // ── 큐 디스패치 (세션 1개 완료 시) ──
  private dispatchFromQueue(freedSession: Session): void {
    if (this.generalQueue.length === 0) return;
    if (this.exclusiveRunning) return;
    if (freedSession.busy || !freedSession.alive) return;

    const task = this.generalQueue.shift()!;
    this.invokeOnSession(freedSession, task.script, false)
      .then(task.resolve, task.reject);
  }

  // ── 큐 일괄 디스패치 (exclusive 완료 시) ──
  private flushGeneralQueue(): void {
    if (this.generalQueue.length === 0) return;
    for (const session of this.generalPool) {
      if (this.generalQueue.length === 0) break;
      if (!session.busy && session.alive) {
        const task = this.generalQueue.shift()!;
        this.invokeOnSession(session, task.script, false)
          .then(task.resolve, task.reject);
      }
    }
  }

  // ── 유휴 세션 탐색 ──
  private findIdle(): Session | null {
    const poolSize = this.generalPool.length;
    for (let i = 0; i < poolSize; i++) {
      const idx = (this.roundRobinIndex + i) % poolSize;
      if (!this.generalPool[idx].busy && this.generalPool[idx].alive) {
        this.roundRobinIndex = (idx + 1) % poolSize;
        return this.generalPool[idx];
      }
    }
    return null;
  }

  // ── exclusive 종료 대기 (이벤트 기반) ──
  private waitForExclusiveEnd(): Promise<void> {
    if (!this.exclusiveRunning) return Promise.resolve();
    return new Promise<void>((resolve) => {
      this.exclusiveEndWaiters.push(resolve);
    });
  }

  // ── general 큐 + 활성 작업 완료 대기 (이벤트 기반) ──
  private waitForGeneralQuiet(): Promise<void> {
    if (this.generalQueue.length === 0 && this.generalActiveCount === 0) return Promise.resolve();
    return new Promise<void>((resolve) => {
      this.generalQuietWaiters.push(resolve);
    });
  }

  // ── 이벤트 시그널 ──
  private signalExclusiveEnd(): void {
    const waiters = this.exclusiveEndWaiters.splice(0);
    for (const resolve of waiters) resolve();
  }

  private signalGeneralQuiet(): void {
    if (this.generalQueue.length === 0 && this.generalActiveCount === 0) {
      const waiters = this.generalQuietWaiters.splice(0);
      for (const resolve of waiters) resolve();
    }
  }

  // ── 세션 복구 ──
  private async recoverSession(session: Session, isExclusive: boolean): Promise<void> {
    await session.dispose();
    try {
      const newSession = await Session.create(session.id);
      if (isExclusive) {
        this.exclusiveSession = newSession;
      } else {
        const idx = this.generalPool.findIndex((s) => s.id === session.id);
        if (idx !== -1) this.generalPool[idx] = newSession;
      }
    } catch {
      // 재생성 실패 → 다음 호출 시 재시도
    }
  }

  // ── heartbeat (병렬 체크) ──
  private async heartbeat(): Promise<void> {
    const checks = this.generalPool.map(async (s) => {
      const alive = await s.healthCheck();
      if (!alive) await this.recoverSession(s, false);
    });
    if (this.exclusiveSession) {
      const exSess = this.exclusiveSession;
      checks.push(
        exSess.healthCheck().then(async (alive) => {
          if (!alive) await this.recoverSession(exSess, true);
        })
      );
    }
    await Promise.all(checks);
  }

  // ── 상태 조회 ──
  getStatus() {
    const now = Date.now();
    return {
      poolMaxSize: POOL_SIZE,
      poolCurrentSize: this.generalPool.length,
      pendingCreations: this.pendingCreations,
      sessions: this.generalPool.map((s) => ({
        id: s.id,
        busy: s.busy,
        alive: s.alive,
      })),
      exclusive: this.exclusiveSession
        ? { id: this.exclusiveSession.id, busy: this.exclusiveSession.busy, alive: this.exclusiveSession.alive }
        : null,
      exclusiveRunning: this.exclusiveRunning,
      generalActiveCount: this.generalActiveCount,
      generalQueue: this.generalQueue.map((t) => ({
        id: t.id,
        waitingMs: now - t.enqueuedAt,
      })),
      exclusiveQueueLength: this.exclusiveQueue.length,
      totalProcessed: this.totalProcessed,
      totalQueued: this.totalQueued,
    };
  }

  // ── 작업 취소 (단건) ──
  cancelTask(taskId: number): boolean {
    const idx = this.generalQueue.findIndex((t) => t.id === taskId);
    if (idx === -1) return false;
    const [task] = this.generalQueue.splice(idx, 1);
    task.reject(new Error(JSON.stringify({ error: true, message: `Task #${taskId} cancelled`, type: "Cancelled" })));
    return true;
  }

  // ── 작업 취소 (전체) ──
  cancelAllTasks(): number {
    const count = this.generalQueue.length;
    for (const task of this.generalQueue) {
      task.reject(new Error(JSON.stringify({ error: true, message: `Task #${task.id} cancelled`, type: "Cancelled" })));
    }
    this.generalQueue = [];
    return count;
  }

  // ── 종료 ──
  async dispose(): Promise<void> {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
    // 큐 잔여 작업 reject
    for (const task of this.generalQueue) {
      task.reject(new Error(JSON.stringify({ error: true, message: "Pool disposed", type: "PoolDisposed" })));
    }
    this.generalQueue = [];
    await Promise.all([
      ...this.generalPool.map((s) => s.dispose()),
      this.exclusiveSession?.dispose() ?? Promise.resolve(),
    ]);
    this.generalPool = [];
    this.exclusiveSession = null;
    this.initialized = false;
  }

  // ── 에러 포맷 ──
  private static formatError(err: unknown): Error {
    const msg = err instanceof Error ? err.message : String(err);
    const cleaned = msg.replace(/\r?\n/g, " ").trim();
    let errorMessage = cleaned;
    const jsonStart = cleaned.indexOf("{");
    const jsonEnd = cleaned.lastIndexOf("}");
    if (jsonStart !== -1 && jsonEnd > jsonStart) {
      try {
        const parsed = JSON.parse(cleaned.slice(jsonStart, jsonEnd + 1));
        errorMessage = parsed.Exception?.Message ?? parsed.FullyQualifiedErrorId ?? cleaned;
      } catch { /* 원본 사용 */ }
    }
    return new Error(JSON.stringify({ error: true, message: errorMessage, type: "PowerShellError" }));
  }
}

// ── 싱글턴 인스턴스 ──
const pool = new SessionPool();

// ── 외부 API ──
export interface RunPSOptions {
  exclusive?: boolean;
}

export async function runPS(script: string, options?: RunPSOptions): Promise<string> {
  if (options?.exclusive) return pool.executeExclusive(script);
  return pool.executeGeneral(script);
}

export function getPoolStatus() {
  return pool.getStatus();
}

export function cancelTask(taskId: number): boolean {
  return pool.cancelTask(taskId);
}

export function cancelAllTasks(): number {
  return pool.cancelAllTasks();
}

export async function dispose(): Promise<void> {
  await pool.dispose();
}
