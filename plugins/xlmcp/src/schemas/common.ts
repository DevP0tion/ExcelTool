import { z } from "zod";

/** 워크북 이름 (생략 시 ActiveWorkbook) */
export const workbookParam = z
  .string()
  .optional()
  .describe("워크북 이름. 생략하면 현재 활성 워크북 사용");

/** 시트 이름 (생략 시 ActiveSheet) */
export const sheetParam = z
  .string()
  .optional()
  .describe("시트 이름. 생략하면 현재 활성 시트 사용");
