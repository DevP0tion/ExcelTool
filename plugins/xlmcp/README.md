# xlmcp

Claude Code에서 Excel을 직접 제어하는 MCP 서버.

실행 중인 Excel 인스턴스에 자동 연결되며, 워크북·시트 파라미터 생략 시 현재 활성 대상을 사용합니다.

- **런타임**: Windows + PowerShell + Excel COM
- **전송**: stdio (Claude Code 플러그인)
- **npm**: `bunx xlmcp@latest`

## 도구 (18개)

### Workbook

| 도구 | 설명 |
|---|---|
| `excel_list_open_workbooks` | 열린 워크북 목록 반환 (이름, 경로, 시트 수) |
| `excel_get_active_workbook` | 활성 워크북의 이름, 경로, 시트 수, 활성 시트명 반환 |
| `excel_create_workbook` | 새 빈 워크북 생성. `savePath` 지정 시 즉시 저장 |
| `excel_save_workbook` | 워크북 저장. `savePath` 지정 시 다른 이름으로 저장 |
| `excel_close_workbook` | 워크북 닫기. `save` 옵션으로 저장 여부 지정 |

### Sheet

| 도구 | 설명 |
|---|---|
| `excel_list_sheets` | 워크북 내 모든 시트 이름 반환 |
| `excel_create_sheet` | 새 시트 추가. `after`로 삽입 위치 지정 가능 |
| `excel_delete_sheet` | 시트 삭제 |
| `excel_copy_sheet` | 시트 복사. `newName`으로 복사본 이름 지정 가능 |
| `excel_rename_sheet` | 시트 이름 변경 |

### Cell / Range

| 도구 | 설명 |
|---|---|
| `excel_read_cell` | 단일 셀의 값, 수식, 표시 텍스트, 표시 형식 반환 |
| `excel_write_cell` | 단일 셀에 값 또는 수식 입력. `=`로 시작하면 수식 처리 |
| `excel_read_range` | 범위를 2D 배열로 반환. 생략 시 UsedRange 전체 읽기 |
| `excel_write_range` | 시작 셀부터 2D 배열 데이터 입력. 수식 혼합 가능 |

### Format

| 도구 | 설명 |
|---|---|
| `excel_format_range` | 서식 적용 — 폰트(이름/크기/굵기/기울임/색상), 배경색, 정렬, 줄바꿈, 테두리, 표시 형식 |
| `excel_set_column_width` | 열 너비 설정. `auto`로 자동 맞춤 가능 |
| `excel_set_row_height` | 행 높이 설정. `auto`로 자동 맞춤 가능 |
| `excel_merge_cells` | 셀 병합 또는 병합 해제 |

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
    ├── sheet/
    ├── cell/
    └── format/
```

각 카테고리 폴더에 `{도구명}.tool.ts` + `index.ts`로 구성. 도구 추가 시 `.tool.ts` 파일 생성 후 해당 `index.ts`에 import 추가.
