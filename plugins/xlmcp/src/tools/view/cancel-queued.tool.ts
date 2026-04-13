import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { cancelTask, cancelAllTasks } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_cancel_queued",
    {
      title: "Cancel Queued Tasks",
      description: "Cancel queued tasks. Cancels all if taskId omitted.",
      inputSchema: {
        taskId: z.number().int().optional().describe("Task ID to cancel. Cancels all if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ taskId }) => {
      if (taskId !== undefined) {
        const found = cancelTask(taskId);
        return textContent({
          success: found,
          message: found ? `Task #${taskId} cancelled` : `Task #${taskId} not found (already running or completed)`,
        });
      } else {
        const count = cancelAllTasks();
        return textContent({
          success: true,
          cancelled: count,
          message: count > 0 ? `${count} queued task(s) cancelled` : "No queued tasks to cancel",
        });
      }
    }
  );
}
