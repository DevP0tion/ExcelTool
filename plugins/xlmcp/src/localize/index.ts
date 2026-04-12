import koKr from "./ko_kr.json";
import enUs from "./en_us.json";
import zhCn from "./zh_cn.json";
import jaJp from "./ja_jp.json";

type LocaleData = Record<string, unknown>;

const locales: Record<string, LocaleData> = {
  ko_kr: koKr as LocaleData,
  en_us: enUs as LocaleData,
  zh_cn: zhCn as LocaleData,
  ja_jp: jaJp as LocaleData,
};

const DEFAULT_LOCALE = "ko_kr";
const envLocale = process.env.XLMCP_LANG?.toLowerCase();
let current: LocaleData = locales[envLocale ?? ""] ?? locales[DEFAULT_LOCALE];

/**
 * 로케일 변경.
 * @example setLocale("en_us")
 */
export function setLocale(locale: string): void {
  const data = locales[locale];
  if (!data) throw new Error(`Unknown locale: ${locale}. Available: ${Object.keys(locales).join(", ")}`);
  current = data;
}

/**
 * 로케일 등록 (런타임에 JSON import 후 추가).
 * @example registerLocale("en_us", enUs)
 */
export function registerLocale(locale: string, data: LocaleData): void {
  locales[locale] = data;
}

/** 현재 로케일 키 반환 */
export function getLocale(): string {
  for (const [key, val] of Object.entries(locales)) {
    if (val === current) return key;
  }
  return "ko_kr";
}

/**
 * 번역 함수. dot notation으로 키를 지정하고, 플레이스홀더를 치환합니다.
 *
 * @param key - dot notation 키 (예: "common.errors.timeout")
 * @param params - 플레이스홀더 치환 값 (예: { ms: 30000 })
 * @returns 번역된 문자열. 키를 찾지 못하면 키 자체를 반환.
 *
 * @example
 * t("common.errors.timeout", { ms: 30000 })
 * // → "타임아웃: 30000ms 초과"
 *
 * t("tools.workbook.create.title")
 * // → "새 워크북 생성"
 *
 * t("common.errors.taskCancelled", { id: 5 })
 * // → "작업 #5 취소됨"
 */
export function t(key: string, params?: Record<string, string | number>): string {
  const val = resolve(current, key);
  if (val === undefined) return key;
  if (typeof val !== "string") return key;
  if (!params) return val;
  return val.replace(/\{(\w+)\}/g, (_, k) => {
    const v = params[k];
    return v !== undefined ? String(v) : `{${k}}`;
  });
}

/** dot notation으로 중첩 객체 탐색 */
function resolve(obj: unknown, path: string): unknown {
  let cur = obj;
  for (const seg of path.split(".")) {
    if (cur == null || typeof cur !== "object") return undefined;
    cur = (cur as Record<string, unknown>)[seg];
  }
  return cur;
}
