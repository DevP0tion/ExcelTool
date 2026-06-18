#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { dispose } from "./services/powershell.js";
import { registerWorkbookTools } from "./tools/workbook/index.js";
import { registerSheetTools } from "./tools/sheet/index.js";
import { registerCellTools } from "./tools/cell/index.js";
import { registerFormatTools } from "./tools/format/index.js";
import { registerDataTools } from "./tools/data/index.js";
import { registerTableTools } from "./tools/table/index.js";
import { registerChartTools } from "./tools/chart/index.js";
import { registerPivotTools } from "./tools/pivot/index.js";
import { registerValidationTools } from "./tools/validation/index.js";
import { registerViewTools } from "./tools/view/index.js";
import { registerImageTools } from "./tools/image/index.js";
import { registerVbaTools } from "./tools/vba/index.js";

const server = new McpServer({
  name: "xlmcp",
  version: "0.3.0",
});

// 도구 등록
registerWorkbookTools(server);
registerSheetTools(server);
registerCellTools(server);
registerFormatTools(server);
registerDataTools(server);
registerTableTools(server);
registerChartTools(server);
registerPivotTools(server);
registerValidationTools(server);
registerViewTools(server);
registerImageTools(server);
registerVbaTools(server);

// stdio transport
const transport = new StdioServerTransport();

// 종료 정리 (중복 dispose 방지 가드).
// SIGINT/SIGTERM 외에 SIGHUP·stdin close/end 도 처리해 비정상 종료(stdio 파이프 끊김)에서도 정리한다.
// 참고: 'exit' 이벤트는 동기만 가능해 비동기 PS dispose 를 끝낼 수 없으므로 다루지 않는다.
// attach-only 설계상 init 단계에서 인스턴스를 만들지 않으므로, 정리 실패해도 유령은 발생하지 않는다.
let shuttingDown = false;
async function shutdown(code = 0): Promise<void> {
  if (shuttingDown) return;
  shuttingDown = true;
  try {
    await dispose();
  } catch {
    /* ignore */
  }
  process.exit(code);
}

process.on("SIGINT", () => { void shutdown(0); });
process.on("SIGTERM", () => { void shutdown(0); });
process.on("SIGHUP", () => { void shutdown(0); });
process.stdin.on("close", () => { void shutdown(0); });
process.stdin.on("end", () => { void shutdown(0); });

await server.connect(transport);
