/**
 * Runtime smoke test for the attach-only resolver. ADAPTIVE: works whether or not
 * a workbook is open, so it is safe to run in either state.
 *
 * Exercises the REAL node-powershell pool end-to-end: INIT_SCRIPT + invoke-wrapper
 * Ensure-Excel injection + actual tool scripts reading unqualified `$excel`.
 *
 *   - Excel CLOSED  -> verifies attach-only null path: list "[]", Resolve-Workbook clean throw,
 *                      pool initializes without spawning any Excel instance.
 *   - Workbook OPEN -> verifies the load-bearing path: wrapper binds $global:excel to the
 *                      user's instance and tools reading `$excel` see it (list returns the
 *                      workbook; Resolve-Workbook resolves it). THIS is the actual bug fix.
 *
 * Run:  bun plugins/xlmcp/scripts/smoke-test.ts
 */
import { runPS, getPoolStatus, dispose } from "../src/services/powershell.js";

function log(...a: unknown[]) {
  console.log("[smoke]", ...a);
}

async function main() {
  let failures = 0;

  // Probe state. The wrapper runs Ensure-Excel before this too, so `$excel` already
  // reflects the resolved workbook-bearing instance (if any).
  const probe = await runPS(`if ($excel) { [int]$excel.Workbooks.Count } else { 0 }`);
  const wbOpen = parseInt(probe.trim() || "0", 10) > 0;
  log(
    wbOpen
      ? "a workbook IS open -> verifying the NON-NULL path (wrapper -> $excel -> tool)"
      : "no workbook open -> verifying the attach-only NULL path"
  );

  // 1) excel_list_open_workbooks logic
  try {
    const list = await runPS(`
      if (-not $excel) {
        "[]"
      } else {
        $result = @()
        foreach ($wb in $excel.Workbooks) { $result += $wb.Name }
        "[" + ($result -join ",") + "]"
      }
    `);
    const t = list.trim();
    log("list_open_workbooks ->", t);
    if (wbOpen) {
      if (t === "[]" || !t.startsWith("[")) {
        log("FAIL: expected a non-empty list while a workbook is open (wrapper->$excel->tool path broken)");
        failures++;
      } else {
        log("OK: wrapper bound $excel to the open instance and the tool saw it");
      }
    } else if (t !== "[]") {
      log("FAIL: expected [] with nothing open");
      failures++;
    }
  } catch (e) {
    log("FAIL list:", (e as Error).message);
    failures++;
  }

  // 2) Resolve-Workbook (the path 44 tools use)
  try {
    const name = await runPS(`$wb = Resolve-Workbook ""; $wb.Name`);
    if (wbOpen) {
      log("OK: Resolve-Workbook resolved ->", name.trim());
    } else {
      log("FAIL: Resolve-Workbook succeeded with nothing open");
      failures++;
    }
  } catch (e) {
    const m = (e as Error).message;
    if (wbOpen) {
      log("FAIL: Resolve-Workbook threw while a workbook is open:", m);
      failures++;
    } else {
      log("Resolve-Workbook threw (expected, clean):", m);
      if (!m.includes("No workbook is open")) {
        log("FAIL: error message is not the clean 'No workbook is open' message");
        failures++;
      }
    }
  }

  // 3) pool initialized sanely (independent of Excel state)
  const st = getPoolStatus();
  log("pool:", JSON.stringify({ size: st.poolCurrentSize, max: st.poolMaxSize, exclusive: !!st.exclusive }));
  if (!st.exclusive || st.poolCurrentSize < 1) {
    log("FAIL: pool did not initialize");
    failures++;
  }

  await dispose();
  log(failures === 0 ? "SMOKE PASS" : `SMOKE FAIL (${failures})`);
  process.exit(failures === 0 ? 0 : 1);
}

main().catch((e) => {
  console.error("[smoke] crash", e);
  process.exit(2);
});
