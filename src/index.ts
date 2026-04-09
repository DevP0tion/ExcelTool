#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { runPS, dispose } from "./services/powershell.js";

const server = new McpServer({
  name: "excel-mcp-server",
  version: "0.1.0",
});

// ── 예시 도구: 시트 목록 조회 ──
server.registerTool(
  "excel_list_sheets",
  {
    title: "List Sheets",
    description: "열려 있는 워크북의 시트 이름 목록을 반환합니다.",
    inputSchema: {
      filePath: z.string().describe("Excel 파일 절대 경로"),
    },
    annotations: { readOnlyHint: true, destructiveHint: false },
  },
  async ({ filePath }) => {
    const result = await runPS(`
      $wb = $excel.Workbooks.Open("${filePath.replace(/\\/g, "\\\\")}")
      $names = @()
      foreach ($ws in $wb.Worksheets) { $names += $ws.Name }
      $wb.Close($false)
      $names -join ","
    `);
    const sheets = result.trim().split(",").filter(Boolean);
    return {
      content: [{ type: "text", text: JSON.stringify({ sheets }) }],
    };
  }
);

// ── 여기에 도구 추가 ──
// import "./tools/read.js";
// import "./tools/write.js";
// ...

// ── stdio transport 시작 ──
const transport = new StdioServerTransport();

process.on("SIGINT", async () => {
  await dispose();
  process.exit(0);
});

await server.connect(transport);
