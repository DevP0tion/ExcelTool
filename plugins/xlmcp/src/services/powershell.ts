import { PowerShell } from "node-powershell";

// ── 설정 ──
const POOL_SIZE = 4;
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
interface Session {
  ps: PowerShell;
  busy: boolean;
  alive: boolean;
  id: number;
}

async function createSession(id: number): Promise<Session> {
  const ps = new PowerShell({
    executableOptions: {
      "-ExecutionPolicy": "Bypass",
      "-NoProfile": true,
    },
  });
  await ps.invoke(INIT_SCRIPT);
  return { ps, busy: false, alive: true, id };
}

// ── 풀 ──
class SessionPool {
  private generalPool: Session[] = [];
  private exclusiveSession: Session | null = null;
  private roundRobinIndex = 0;
  private initialized = false;

  // exclusive 실행 중 general 차단용
  private exclusiveRunning = false;
  private exclusiveQueue: Array<{
    script: string;
    resolve: (v: string) => void;
    reject: (e: Error) => void;
  }> = [];
  private generalActiveCount = 0;
  private generalDrainResolve: (() => void) | null = null;

  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;

  async init(): Promise<void> {
    if (this.initialized) return;

    // general pool 생성 (병렬 초기화)
    const sessions = await Promise.all(
      Array.from({ length: POOL_SIZE }, (_, i) => createSession(i))
    );
    this.generalPool = sessions;

    // exclusive 전용 세션
    this.exclusiveSession = await createSession(100);

    // heartbeat 시작
    this.heartbeatTimer = setInterval(() => this.heartbeat(), HEARTBEAT_INTERVAL);

    this.initialized = true;
  }

  // ── 일반 실행 ──
  async executeGeneral(script: string): Promise<string> {
    await this.init();

    // exclusive 실행 중이면 대기
    if (this.exclusiveRunning) {
      await this.waitForExclusiveEnd();
    }

    const session = this.pickGeneral();
    return this.invokeOnSession(session, script, false);
  }

  // ── exclusive 실행 ──
  async executeExclusive(script: string): Promise<string> {
    await this.init();

    // 이미 exclusive 실행 중이면 큐에 대기
    if (this.exclusiveRunning) {
      return new Promise<string>((resolve, reject) => {
        this.exclusiveQueue.push({ script, resolve, reject });
      });
    }

    return this.runExclusive(script);
  }

  private async runExclusive(script: string): Promise<string> {
    this.exclusiveRunning = true;

    // drain: general pool의 진행 중 작업 완료 대기
    if (this.generalActiveCount > 0) {
      await new Promise<void>((resolve) => {
        this.generalDrainResolve = resolve;
      });
    }

    try {
      const result = await this.invokeOnSession(this.exclusiveSession!, script, true);
      return result;
    } finally {
      // 큐에 대기 중인 exclusive 작업 처리
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
    session.busy = true;

    const wrapped = `
      try {
        ${script}
      } catch {
        [Console]::Error.WriteLine(($_ | ConvertTo-Json -Compress))
        throw $_
      }
    `;

    try {
      const result = await this.withTimeout(
        session.ps.invoke(wrapped),
        INVOKE_TIMEOUT
      );
      return result.raw ?? "";
    } catch (err: unknown) {
      // 프로세스 사망 가능성 → 복구 시도
      if (this.isProcessDead(err)) {
        await this.recoverSession(session, isExclusive);
      }
      const msg = err instanceof Error ? err.message : String(err);
      const cleaned = msg.replace(/\r?\n/g, " ").trim();
      let errorMessage = cleaned;
      const jsonStart = cleaned.indexOf("{");
      const jsonEnd = cleaned.lastIndexOf("}");
      if (jsonStart !== -1 && jsonEnd > jsonStart) {
        try {
          const parsed = JSON.parse(cleaned.slice(jsonStart, jsonEnd + 1));
          errorMessage =
            parsed.Exception?.Message ??
            parsed.FullyQualifiedErrorId ??
            cleaned;
        } catch {
          // 원본 사용
        }
      }
      throw new Error(
        JSON.stringify({
          error: true,
          message: errorMessage,
          type: "PowerShellError",
        })
      );
    } finally {
      session.busy = false;
      if (!isExclusive) {
        this.generalActiveCount--;
        // drain 대기 해제
        if (
          this.exclusiveRunning &&
          this.generalActiveCount === 0 &&
          this.generalDrainResolve
        ) {
          this.generalDrainResolve();
          this.generalDrainResolve = null;
        }
      }
    }
  }

  // ── 라운드 로빈 ──
  private pickGeneral(): Session {
    // busy가 아닌 세션 우선
    for (let i = 0; i < POOL_SIZE; i++) {
      const idx = (this.roundRobinIndex + i) % POOL_SIZE;
      if (!this.generalPool[idx].busy && this.generalPool[idx].alive) {
        this.roundRobinIndex = (idx + 1) % POOL_SIZE;
        return this.generalPool[idx];
      }
    }
    // 모두 busy면 라운드 로빈 (node-powershell 내부 큐에 의존)
    const session = this.generalPool[this.roundRobinIndex];
    this.roundRobinIndex = (this.roundRobinIndex + 1) % POOL_SIZE;
    return session;
  }

  // ── exclusive 대기 ──
  private waitForExclusiveEnd(): Promise<void> {
    return new Promise<void>((resolve) => {
      const check = () => {
        if (!this.exclusiveRunning) {
          resolve();
        } else {
          setTimeout(check, 50);
        }
      };
      check();
    });
  }

  // ── 타임아웃 래퍼 ──
  private withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(
        () => reject(new Error(`타임아웃: ${ms}ms 초과`)),
        ms
      );
      promise.then(
        (v) => {
          clearTimeout(timer);
          resolve(v);
        },
        (e) => {
          clearTimeout(timer);
          reject(e);
        }
      );
    });
  }

  // ── 프로세스 사망 판정 ──
  private isProcessDead(err: unknown): boolean {
    const msg = err instanceof Error ? err.message : String(err);
    return (
      msg.includes("process exited") ||
      msg.includes("invoke called after") ||
      msg.includes("EPIPE") ||
      msg.includes("타임아웃")
    );
  }

  // ── 세션 복구 ──
  private async recoverSession(
    session: Session,
    isExclusive: boolean
  ): Promise<void> {
    session.alive = false;
    try {
      await session.ps.dispose();
    } catch {
      // ignore
    }

    try {
      const newSession = await createSession(session.id);
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
    const checkSession = async (
      session: Session,
      isExclusive: boolean
    ): Promise<void> => {
      if (session.busy || !session.alive) return;
      try {
        await this.withTimeout(session.ps.invoke("$excel.Version"), 5000);
      } catch {
        await this.recoverSession(session, isExclusive);
      }
    };

    for (const s of this.generalPool) {
      checkSession(s, false);
    }
    if (this.exclusiveSession) {
      checkSession(this.exclusiveSession, true);
    }
  }

  // ── 종료 ──
  async dispose(): Promise<void> {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }

    const releaseCOM = `
      if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
      }
    `;

    const disposeSession = async (s: Session) => {
      try {
        await s.ps.invoke(releaseCOM);
      } catch {
        // ignore
      }
      try {
        await s.ps.dispose();
      } catch {
        // ignore
      }
    };

    await Promise.all([
      ...this.generalPool.map(disposeSession),
      this.exclusiveSession ? disposeSession(this.exclusiveSession) : Promise.resolve(),
    ]);

    this.generalPool = [];
    this.exclusiveSession = null;
    this.initialized = false;
  }
}

// ── 싱글턴 인스턴스 ──
const pool = new SessionPool();

// ── 외부 API (기존과 동일한 시그니처 유지) ──
export interface RunPSOptions {
  exclusive?: boolean;
}

export async function runPS(
  script: string,
  options?: RunPSOptions
): Promise<string> {
  if (options?.exclusive) {
    return pool.executeExclusive(script);
  }
  return pool.executeGeneral(script);
}

export async function dispose(): Promise<void> {
  await pool.dispose();
}
