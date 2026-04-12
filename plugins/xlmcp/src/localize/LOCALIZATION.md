# 지역화 (Localization)

## 개요

xlmcp는 도구의 title, description, 파라미터 설명, 에러/상태 메시지를 JSON 기반으로 지역화합니다.

```
src/localize/
├── index.ts       ← t(), setLocale(), registerLocale(), getLocale()
├── ko_kr.json     ← 한국어 (기본)
├── en_us.json     ← English
├── zh_cn.json     ← 简体中文
└── ja_jp.json     ← 日本語
```

## 로케일 전환

### 환경변수 (권장)

```jsonc
// Claude Desktop / .mcp.json
{
  "excel": {
    "command": "bunx",
    "args": ["xlmcp@latest"],
    "env": {
      "XLMCP_LANG": "en_us"
    }
  }
}
```

| 코드 | 언어 |
|---|---|
| `ko_kr` | 한국어 (기본) |
| `en_us` | English |
| `zh_cn` | 简体中文 |
| `ja_jp` | 日本語 |

미설정 또는 잘못된 값이면 `ko_kr`이 기본 사용됩니다.

### 코드에서 전환

```typescript
import { t, setLocale, getLocale } from "./localize/index.js";

setLocale("en_us");
t("tools.cell.readRange.title"); // "Read Range"
getLocale(); // "en_us"
```

## JSON 키 구조

```
{locale}.json
├── common
│   ├── params                ← 공통 파라미터 (workbook, sheet)
│   └── errors                ← 공통 에러 메시지
├── tools
│   └── {category}
│       └── {tool}
│           ├── title         ← 도구 제목
│           ├── description   ← 도구 설명
│           ├── params        ← 파라미터별 설명
│           ├── messages      ← 상태/결과 메시지 (선택)
│           └── errors        ← 도구별 에러 메시지 (선택)
├── ps
│   └── errors                ← PowerShell 스크립트 내 메시지
└── format
    └── noChanges             ← 서식 변경 없음 메시지
```

### 카테고리 목록

`workbook`, `sheet`, `cell`, `format`, `data`, `table`, `chart`, `pivot`, `validation`, `image`, `vba`, `view`

## 새 로케일 추가 가이드

### 1. JSON 파일 생성

`ko_kr.json`을 복사하여 `{locale}.json`을 생성합니다.

```bash
cp src/localize/ko_kr.json src/localize/fr_fr.json
```

### 2. 번역

모든 값(value)을 번역합니다. 키(key)는 변경하지 마세요.

```jsonc
// ❌ 키를 번역하지 마세요
{ "titre": "Créer un classeur" }

// ✅ 값만 번역
{ "title": "Créer un classeur" }
```

### 3. index.ts에 등록

```typescript
import frFr from "./fr_fr.json";

const locales: Record<string, LocaleData> = {
  // ...기존 로케일
  fr_fr: frFr as LocaleData,
};
```

### 4. 키 수 검증

모든 로케일은 동일한 키 수를 가져야 합니다.

```bash
python3 -c "
import json
def count(obj):
    c = 0
    for v in obj.values():
        c += count(v) if isinstance(v, dict) else 1
    return c
for f in ['ko_kr.json', 'en_us.json', 'zh_cn.json', 'ja_jp.json', 'fr_fr.json']:
    print(f'{f}: {count(json.load(open(f)))} keys')
"
```

현재 기준: **255키**

## 번역 규칙

### 플레이스홀더

`{변수명}` 형식의 플레이스홀더는 번역하지 않고 그대로 유지합니다.

```jsonc
// ko_kr
"timeout": "타임아웃: {ms}ms 초과"

// en_us
"timeout": "Timeout: exceeded {ms}ms"

// ❌ 플레이스홀더 변경 금지
"timeout": "Timeout: exceeded {milliseconds}ms"
```

| 플레이스홀더 | 용도 |
|---|---|
| `{id}` | 작업 ID |
| `{ms}` | 밀리초 |
| `{count}` | 개수 |
| `{path}` | 파일 경로 |
| `{action}` | 동작 이름 |
| `{preview}` | 미리보기 텍스트 |

### 기술 용어

다음 용어는 번역하지 않고 원문 유지합니다.

- Excel 기능명: `UsedRange`, `ListObject`, `Named Range`, `VBA`
- 파라미터 값: `values`, `formulas`, `auto`, `true`, `false`
- 파일 형식: `.xlsx`, `.xlsm`, `PNG`, `JPG`
- 함수/타입명: `RGB hex`, `Shape`, `PowerShell`
- 도구 이름: `excel_copy_paste_format`, `excel_pool_status` 등

### 예시 값

예시의 셀 주소, 범위, 경로 등은 로케일에 맞게 변경할 수 있습니다.

```jsonc
// ko_kr
"fontName": "폰트 이름 (예: '맑은 고딕')"

// ja_jp
"fontName": "フォント名（例：'游ゴシック'）"

// en_us
"fontName": "Font name (e.g. 'Arial')"
```

### 줄바꿈

`\n`으로 표현되는 줄바꿈은 유지합니다. 줄바꿈 위치는 의미에 맞게 조정 가능합니다.

### ⚠️ 경고 접두사

`⚠️`로 시작하는 경고 문구는 접두사를 유지합니다.

## t() 함수 사용법

```typescript
import { t } from "./localize/index.js";

// 단순 조회
t("tools.workbook.create.title")
// → "새 워크북 생성" (ko_kr)

// 플레이스홀더 치환
t("common.errors.timeout", { ms: 30000 })
// → "타임아웃: 30000ms 초과"

// 키 없으면 키 자체 반환 (fallback)
t("unknown.key")
// → "unknown.key"

// 런타임 로케일 추가
import { registerLocale, setLocale } from "./localize/index.js";
registerLocale("fr_fr", frFrData);
setLocale("fr_fr");
```
