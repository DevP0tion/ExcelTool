# excel-mcp

PowerShell COM 자동화 기반 Excel MCP 서버.

열려 있는 Excel 인스턴스를 자동 감지하여 연결합니다. 워크북/시트 파라미터 생략 시 ActiveWorkbook/ActiveSheet를 사용합니다.

> **Windows 전용** — PowerShell + Excel COM을 사용하므로 Windows에서만 동작합니다.

## 설치

### 권장: Claude Code 플러그인

```
/plugin install /path/to/excel-mcp
```

플러그인으로 설치하면 `bunx excel-mcp@latest`가 자동 실행됩니다.

### 비권장: 모듈 단독 설치

```bash
bun add -g excel-mcp
# 또는
npm install -g excel-mcp
```

> ⚠️ 단독 설치 시 Claude Code와의 연결은 별도로 MCP 설정을 직접 구성해야 합니다.
> 플러그인 설치를 사용하면 MCP 서버 등록이 자동으로 처리되므로, 특별한 이유가 없다면 플러그인 방식을 권장합니다.

## 도구 목록 (18개)

### Workbook (5)
| 도구 | 설명 |
|---|---|
| `excel_list_open_workbooks` | 열린 워크북 목록 (이름, 경로, 시트 수) |
| `excel_get_active_workbook` | 활성 워크북 정보 |
| `excel_create_workbook` | 새 워크북 생성 |
| `excel_save_workbook` | 저장 / 다른 이름으로 저장 |
| `excel_close_workbook` | 워크북 닫기 |

### Sheet (5)
| 도구 | 설명 |
|---|---|
| `excel_list_sheets` | 시트 목록 |
| `excel_create_sheet` | 시트 추가 |
| `excel_delete_sheet` | 시트 삭제 |
| `excel_copy_sheet` | 시트 복사 |
| `excel_rename_sheet` | 시트 이름 변경 |

### Cell / Range (4)
| 도구 | 설명 |
|---|---|
| `excel_read_cell` | 단일 셀 읽기 (값, 수식, 표시 텍스트) |
| `excel_write_cell` | 단일 셀 쓰기 (값 또는 수식) |
| `excel_read_range` | 범위 읽기 → 2D 배열 (생략 시 UsedRange) |
| `excel_write_range` | 범위 쓰기 ← 2D 배열 |

### Format (4)
| 도구 | 설명 |
|---|---|
| `excel_format_range` | 서식 적용 (폰트, 색상, 정렬, 테두리, 표시 형식 등) |
| `excel_set_column_width` | 열 너비 (숫자 또는 auto) |
| `excel_set_row_height` | 행 높이 (숫자 또는 auto) |
| `excel_merge_cells` | 셀 병합 / 해제 |

## 구조

```
src/
├── index.ts
├── services/
│   ├── powershell.ts
│   └── utils.ts
├── schemas/
│   └── common.ts
└── tools/
    ├── workbook/
    │   ├── index.ts
    │   ├── list-open-workbooks.tool.ts
    │   ├── get-active-workbook.tool.ts
    │   ├── create-workbook.tool.ts
    │   ├── save-workbook.tool.ts
    │   └── close-workbook.tool.ts
    ├── sheet/
    │   ├── index.ts
    │   ├── list-sheets.tool.ts
    │   ├── create-sheet.tool.ts
    │   ├── delete-sheet.tool.ts
    │   ├── copy-sheet.tool.ts
    │   └── rename-sheet.tool.ts
    ├── cell/
    │   ├── index.ts
    │   ├── read-cell.tool.ts
    │   ├── write-cell.tool.ts
    │   ├── read-range.tool.ts
    │   └── write-range.tool.ts
    └── format/
        ├── index.ts
        ├── format-range.tool.ts
        ├── set-column-width.tool.ts
        ├── set-row-height.tool.ts
        └── merge-cells.tool.ts
```
