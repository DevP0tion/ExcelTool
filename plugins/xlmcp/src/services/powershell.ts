import { PowerShell } from "node-powershell";

// ── 설정 ──
const POOL_SIZE = Math.max(1, parseInt(process.env.XLMCP_POOL_SIZE ?? "4", 10) || 4);
const HEARTBEAT_INTERVAL = 10_000;
const INVOKE_TIMEOUT = 30_000;

// ── PS 초기화 스크립트 ──
const INIT_SCRIPT = `
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
        ${script}
      } catch {
        [Console]::Error.WriteLine(($_ | ConvertTo-Json -Compress))
        throw $_
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
      await Session.withTimeout(this.ps.invoke("$excel.Version"), 5000);
      return true;
    } catch {
      this.alive = false;
      return false;
    }
  }

  async dispose(): Promise<void> {
    try {
      await this.ps.invoke(`
        if ($excel) {
          [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
      `);
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
      msg.includes("타임아웃")
    );
  }

  static async create(id: number): Promise<Session> {
    const session = new Session(id);
    await session.init();
    return session;
  }

  private static withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error(`타임아웃: ${ms}ms 초과`)), ms);
      promise.then(
        (v) => { clearTimeout(timer); resolve(v); },
        (e) => { clearTimeout(timer); reject(e); }
      );
    });
  }
}

// ── 세션 풀 ──
class SessionPool {
  private generalPool: Session[] = [];
  private exclusiveSession: Session | null = null;
  private roundRobinIndex = 0;
  private initialized = false;

  private exclusiveRunning = false;
  private exclusiveQueue: Array<{
    script: string;
    resolve: (v: string) => void;
    reject: (e: Error) => void;
  }> = [];
  private generalActiveCount = 0;
  private generalDrainResolve: (() => void) | null = null;

  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  private nextGeneralId = 0;

  // ── 초기화 (1개만 생성) ──
  async init(): Promise<void> {
    if (this.initialized) return;

    const first = await Session.create(this.nextGeneralId++);
    this.generalPool = [first];

    this.exclusiveSession = await Session.create(100);
    this.heartbeatTimer = setInterval(() => this.heartbeat(), HEARTBEAT_INTERVAL);
    this.initialized = true;
  }

  // ── 일반 실행 ──
  async executeGeneral(script: string): Promise<string> {
    await this.init();
    if (this.exclusiveRunning) {
      await this.waitForExclusiveEnd();
    }
    const session = await this.pickGeneral();
    return this.invokeOnSession(session, script, false);
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
      if (next) {
        this.runExclusive(next.script).then(next.resolve, next.reject);
      } else {
        this.exclusiveRunning = false;
      }
    }
  }

  // ── 세션에서 실행 ──
  private async invokeOnSession(
    session: Session,
    script: string,
    isExclusive: boolean
  ): Promise<string> {
    if (!isExclusive) this.generalActiveCount++;
    try {
      return await session.invoke(script, INVOKE_TIMEOUT);
    } catch (err: unknown) {
      if (session.isProcessDead(err)) {
        await this.recoverSession(session, isExclusive);
      }
      throw SessionPool.formatError(err);
    } finally {
      if (!isExclusive) {
        this.generalActiveCount--;
        if (this.exclusiveRunning && this.generalActiveCount === 0 && this.generalDrainResolve) {
          this.generalDrainResolve();
          this.generalDrainResolve = null;
        }
      }
    }
  }

  // ── 세션 선택 (lazy growth) ──
  private async pickGeneral(): Promise<Session> {
    const poolSize = this.generalPool.length;

    // 1. 유휴 세션 탐색
    for (let i = 0; i < poolSize; i++) {
      const idx = (this.roundRobinIndex + i) % poolSize;
      if (!this.generalPool[idx].busy && this.generalPool[idx].alive) {
        this.roundRobinIndex = (idx + 1) % poolSize;
        return this.generalPool[idx];
      }
    }

    // 2. 모두 busy + 상한 미도달 → 새 세션 생성
    if (poolSize < POOL_SIZE) {
      const newSession = await Session.create(this.nextGeneralId++);
      this.generalPool.push(newSession);
      return newSession;
    }

    // 3. 상한 도달 → 라운드 로빈 (node-powershell 내부 큐에 의존)
    const session = this.generalPool[this.roundRobinIndex];
    this.roundRobinIndex = (this.roundRobinIndex + 1) % poolSize;
    return session;
  }

  // ── exclusive 대기 ──
  private waitForExclusiveEnd(): Promise<void> {
    return new Promise<void>((resolve) => {
      const check = () => {
        if (!this.exclusiveRunning) resolve();
        else setTimeout(check, 50);
      };
      check();
    });
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

  // ── heartbeat ──
  private async heartbeat(): Promise<void> {
    for (const s of this.generalPool) {
      const alive = await s.healthCheck();
      if (!alive) await this.recoverSession(s, false);
    }
    if (this.exclusiveSession) {
      const alive = await this.exclusiveSession.healthCheck();
      if (!alive) await this.recoverSession(this.exclusiveSession, true);
    }
  }

  // ── 종료 ──
  async dispose(): Promise<void> {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
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

export async function dispose(): Promise<void> {
  await pool.dispose();
}
