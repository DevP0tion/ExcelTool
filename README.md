# excel-mcp-server

PowerShell COM 자동화 기반 Excel MCP 서버. Claude Code 플러그인으로 동작.

## 설치

```bash
bun install
```

## Claude Code 플러그인 등록

Claude Code 내에서:
```
/plugin install /path/to/excel-mcp-server
```

## 개발 모드 (단독 테스트)

```bash
bun run dev
```

## 구조

```
src/
├── index.ts              # MCP 서버 진입점 (stdio)
├── services/
│   └── powershell.ts     # PowerShell COM 래퍼
├── tools/                # 도구 구현 (도메인별 분리)
├── schemas/              # Zod 스키마
└── constants.ts          # 공유 상수
```

## 도구 추가 방법

`src/tools/` 아래에 파일 생성 후 `server.registerTool()`로 등록,
`src/index.ts`에서 import.
