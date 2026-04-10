/** PowerShell 문자열 내 특수문자 이스케이프 */
export function psEscape(str: string): string {
  return str.replace(/'/g, "''");
}

/** runPS 결과를 JSON으로 파싱. 실패 시 raw 텍스트 반환 */
export function parseJSON<T = unknown>(raw: string): T {
  const trimmed = raw.trim();
  return JSON.parse(trimmed);
}

/** MCP text content 래퍼 */
export function textContent(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
}

/** MCP error content 래퍼 */
export function errorContent(message: string) {
  return { content: [{ type: "text" as const, text: JSON.stringify({ error: message }) }], isError: true };
}
