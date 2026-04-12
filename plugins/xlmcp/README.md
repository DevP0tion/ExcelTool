# xlmcp

Excel을 직접 제어하는 MCP 서버.

실행 중인 Excel 인스턴스에 자동 연결되며, 워크북·시트 파라미터 생략 시 현재 활성 대상을 사용합니다.

- **런타임**: Windows + PowerShell + Excel COM
- **전송**: stdio
- **npm**: `bunx xlmcp@latest`

## 아키텍처

- **Session Pool**: PowerShell 프로세스 4개 (General) + 1개 (Exclusive). 라운드 로빈 배분.
- **Exclusive Queue**: 클립보드·Selection·구조 변경 도구는 직렬 실행. 실행 중 General Pool 일시 차단.
- **HealthMonitor**: 10초 heartbeat, 사망 감지 시 자동 재생성, 호출당 30초 타임아웃.
- **대용량 쓰기**: 500행 이상 시 자동 청크 분할 + JSON 임시 파일 + 병렬 쓰기.
- **가상 클립보드**: 값/수식 복사 시 시스템 클립보드 미사용. 서식 복사만 시스템 클립보드 사용.

## 도구 (42개)

### Workbook (6)

| 도구 | 설명 |
|---|---|
| `excel_list_open_workbooks` | 열린 워크북 목록 반환 (이름, 경로, 시트 수) |
| `excel_get_active_workbook` | 활성 워크북의 이름, 경로, 시트 수, 활성 시트명 반환 |
| `excel_create_workbook` | 새 빈 워크북 생성. `savePath` 지정 시 즉시 저장 |
| `excel_open_workbook` | 파일 경로로 워크북 열기. 이미 열려 있으면 활성화 |
| `excel_save_workbook` | 워크북 저장. `savePath` 지정 시 다른 이름으로 저장 |
| `excel_close_workbook` | 워크북 닫기. `save` 옵션으로 저장 여부 지정 |

### Sheet (5)

| 도구 | 설명 |
|---|---|
| `excel_list_sheets` | 워크북 내 모든 시트 이름 반환 |
| `excel_create_sheet` | 새 시트 추가. `after`로 삽입 위치 지정 가능 |
| `excel_delete_sheet` | 시트 삭제 |
| `excel_copy_sheet` | 시트 복사. `newName`으로 복사본 이름 지정 가능 |
| `excel_rename_sheet` | 시트 이름 변경 |

### Cell / Range (6)

| 도구 | 설명 |
|---|---|
| `excel_read_cell` | 단일 셀의 값, 수식, 표시 텍스트, 표시 형식 반환 |
| `excel_write_cell` | 단일 셀에 값 또는 수식 입력. `=`로 시작하면 수식 처리 |
| `excel_read_range` | 범위를 2D 배열로 반환. 타입 유지. 생략 시 UsedRange |
| `excel_write_range` | 시작 셀부터 2D 배열 데이터 입력. 숫자 자동 감지, 수식 혼합 가능. 대용량 시 자동 청크 분할 병렬 쓰기 |
| `excel_read_range_formulas` | 범위의 수식을 2D 배열로 반환 |
| `excel_clear_range` | 범위 삭제 (값만/서식만/전체) |

### Format (6)

| 도구 | 설명 |
|---|---|
| `excel_format_range` | 서식 적용 — 폰트, 배경색, 정렬, 테두리(전체/개별), 표시 형식 |
| `excel_set_column_width` | 열 너비 설정. `auto`로 자동 맞춤 가능 |
| `excel_set_row_height` | 행 높이 설정. `auto`로 자동 맞춤 가능 |
| `excel_merge_cells` | 셀 병합 또는 병합 해제 |
| `excel_read_cell_format` | 셀의 현재 서식 정보 반환 (폰트, 색상, 정렬, 테두리, 병합 여부) |
| `excel_write_cell_format` | 서식 데이터 기반 일괄 적용. `read_cell_format` 출력으로 서식 복제 가능 |

### Data (6)

| 도구 | 설명 |
|---|---|
| `excel_insert_delete_rows_cols` | 행 또는 열 삽입·삭제 |
| `excel_copy_paste_range` | 값 또는 수식 복사/붙여넣기. 시스템 클립보드 미사용, 병렬 안전 |
| `excel_copy_paste_format` | 서식 복사/붙여넣기. 값·수식과 함께 필요 시 `copy_paste_range`와 순차 호출 |
| `excel_find_replace` | 시트 내 찾기/바꾸기. 찾기만도 가능 |
| `excel_sort_range` | 범위를 지정 열 기준으로 정렬 |
| `excel_auto_filter` | 자동 필터 설정/해제/조건 적용 |

### Table (4)

| 도구 | 설명 |
|---|---|
| `excel_list_tables` | 시트 내 표(ListObject) 목록 반환 |
| `excel_create_table` | 범위를 Excel 표로 변환 |
| `excel_edit_table` | 표 이름, 스타일, 크기 변경, 행/열 추가, 요약 행 토글 |
| `excel_delete_table` | 표 삭제 (데이터 유지 또는 함께 삭제) |

### Chart (1)

| 도구 | 설명 |
|---|---|
| `excel_create_chart` | 차트 생성 (line/bar/column/pie/scatter/area) |

### Pivot (1)

| 도구 | 설명 |
|---|---|
| `excel_create_pivot_table` | 피벗 테이블 생성. 행/열/데이터 필드 및 집계 함수 지정 |

### Validation (2)

| 도구 | 설명 |
|---|---|
| `excel_set_data_validation` | 데이터 유효성 검사 (드롭다운, 숫자 범위, 수식 등) |
| `excel_set_conditional_format` | 조건부 서식 (셀 값, 수식, 색조, 데이터 막대) |

### View (4)

| 도구 | 설명 |
|---|---|
| `excel_freeze_panes` | 틀 고정/해제 |
| `excel_named_range` | 이름 정의 조회/생성/삭제 |
| `excel_pool_status` | 세션 풀 상태 조회 (세션 수, busy/alive, 큐 길이, 처리 통계) |
| `excel_cancel_queued` | 대기 중인 작업 취소 (단건 또는 전체). 실행 중 작업은 취소 불가 |

## 구조

```
src/
├── index.ts
├── services/
│   ├── powershell.ts       # SessionPool (General 4 + Exclusive 1)
│   └── utils.ts            # psEscape, parseJSON, textContent
├── schemas/
│   └── common.ts
└── tools/
    ├── workbook/            # 6 도구
    ├── sheet/               # 5 도구
    ├── cell/                # 6 도구
    ├── format/              # 6 도구
    ├── data/                # 6 도구
    ├── table/               # 4 도구
    ├── chart/               # 1 도구
    ├── pivot/               # 1 도구
    ├── validation/          # 2 도구
    └── view/                # 4 도구 (+ 1 exclusive)
```

각 카테고리 폴더에 `{도구명}.tool.ts` + `index.ts`로 구성. 도구 추가 시 `.tool.ts` 파일 생성 후 해당 `index.ts`에 import 추가.
