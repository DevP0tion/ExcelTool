import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_insert_image",
    {
      title: "Insert Image",
      description:
        "Insert image file into sheet. Embedded in workbook, persists after source deletion.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        filePath: z.string().describe("Absolute image path (PNG, JPG, BMP, GIF)"),
        cell: z.string().describe("Anchor cell (e.g. E1)"),
        width: z.number().optional().describe("Width px. Original size if omitted"),
        height: z.number().optional().describe("Height px. Original size if omitted"),
        name: z.string().optional().describe("Shape name. Auto if omitted"),
        keepAspect: z
          .boolean()
          .default(true)
          .describe("Keep aspect ratio when only width or height is set"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, filePath, cell, width, height, name, keepAspect }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // 크기 결정 PS 스크립트
      let sizeScript: string;
      if (width && height) {
        // 둘 다 지정: 그대로 사용
        sizeScript = `$pic.Width = ${width}; $pic.Height = ${height}`;
      } else if (width) {
        sizeScript = keepAspect
          ? `$ratio = $pic.Height / $pic.Width; $pic.Width = ${width}; $pic.Height = ${width} * $ratio`
          : `$pic.Width = ${width}`;
      } else if (height) {
        sizeScript = keepAspect
          ? `$ratio = $pic.Width / $pic.Height; $pic.Height = ${height}; $pic.Width = ${height} * $ratio`
          : `$pic.Height = ${height}`;
      } else {
        sizeScript = ""; // 원본 크기 유지
      }

      const nameScript = name
        ? `$pic.Name = '${psEscape(name)}'`
        : "";

      const raw = await runPS(`
        $imgPath = '${psEscape(filePath)}'
        if (-not (Test-Path $imgPath)) {
          throw "Image file not found: $imgPath"
        }
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $pos = $ws.Range('${psEscape(cell)}')
        $pic = $ws.Shapes.AddPicture(
          $imgPath,
          0,
          -1,
          $pos.Left,
          $pos.Top,
          -1,
          -1
        )
        ${sizeScript}
        ${nameScript}
        @{
          Name = $pic.Name
          Width = [math]::Round($pic.Width, 1)
          Height = [math]::Round($pic.Height, 1)
          Cell = '${psEscape(cell)}'
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
