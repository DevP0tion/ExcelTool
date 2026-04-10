import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as createChart } from "./create-chart.tool.js";

export function registerChartTools(server: McpServer) {
  createChart(server);
}
