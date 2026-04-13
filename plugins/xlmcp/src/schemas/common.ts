import { z } from "zod";

/** 워크북 이름 (생략 시 ActiveWorkbook) */
export const workbookParam = z
  .string()
  .optional()
  .describe("Workbook name. Uses active workbook if omitted");

/** 시트 이름 (생략 시 ActiveSheet) */
export const sheetParam = z
  .string()
  .optional()
  .describe("Sheet name. Uses active sheet if omitted");
