import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_manage_image",
    {
      title: "Manage Image",
      description: "Delete, move, or resize an image.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("Image (Shape) name"),
        action: z.enum(["delete", "move", "resize"]).describe("Action: delete, move, or resize"),
        cell: z.string().optional().describe("Target cell for move (e.g. A10)"),
        width: z.number().optional().describe("Width px for resize"),
        height: z.number().optional().describe("Height px for resize"),
        keepAspect: z
          .boolean()
          .default(true)
          .describe("Keep aspect ratio when only one dimension is set"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, name, action, cell, width, height, keepAspect }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (action === "delete") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Shapes.Item('${psEscape(name)}').Delete()
        `);
        return textContent({ success: true, action: "deleted", name });
      }

      if (action === "move") {
        if (!cell) throw new Error("Cell parameter is required for move.");
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $s = $ws.Shapes.Item('${psEscape(name)}')
          $pos = $ws.Range('${psEscape(cell)}')
          $s.Left = $pos.Left
          $s.Top = $pos.Top
          @{
            Name = $s.Name
            Left = [math]::Round($s.Left, 1)
            Top = [math]::Round($s.Top, 1)
          } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      }

      // resize
      let sizeScript: string;
      if (width && height) {
        sizeScript = `
          $s.LockAspectRatio = 0
          $s.Width = ${width}
          $s.Height = ${height}`;
      } else if (width) {
        sizeScript = keepAspect
          ? `$ratio = $s.Height / $s.Width; $s.Width = ${width}; $s.Height = ${width} * $ratio`
          : `$s.LockAspectRatio = 0; $s.Width = ${width}`;
      } else if (height) {
        sizeScript = keepAspect
          ? `$ratio = $s.Width / $s.Height; $s.Height = ${height}; $s.Width = ${height} * $ratio`
          : `$s.LockAspectRatio = 0; $s.Height = ${height}`;
      } else {
        throw new Error("Width or height is required for resize.");
      }

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $s = $ws.Shapes.Item('${psEscape(name)}')
        ${sizeScript}
        @{
          Name = $s.Name
          Width = [math]::Round($s.Width, 1)
          Height = [math]::Round($s.Height, 1)
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
