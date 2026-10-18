# Sheet Studio — Microsoft Excel 대비 종합 검증 보고서 (Overview)

- **검증 일자**: 2026-10-18
- **대상 프로젝트**: `lattice-sheet-studio` (v0.7.7)
- **검증 방법**: 정적 코드 전수 감사 + Node.js 런타임 재현 스크립트(수식 엔진, 클립보드/붙여넣기 플랜, OOXML ZIP/DOM 구조, UI 레이아웃 좌표 계산) 실증 대조

---

## 1. 프로젝트 아키텍처 요약

**Sheet Studio**는 Microsoft Excel의 한국어 리본 UI와 핵심 스프레드시트 작업 흐름을 독자적으로 구현한 Electron/웹 애플리케이션입니다.

| 계층 | 주요 파일 | 역할 및 특징 |
|---|---|---|
| **앱 셸 & 브리지** | [`electron/main.cjs`](../electron/main.cjs), [`electron/preload.cjs`](../electron/preload.cjs), [`electron/vba.cjs`](../electron/vba.cjs) | 컨텍스트 격리된 파일 입출력 다이얼로그 및 Windows Excel COM 기반 `.xlsm` VBA 매크로 실행 브리지 |
| **메인 UI & 그리드** | [`src/main.mjs`](../src/main.mjs), [`src/worksheet-view.mjs`](../src/worksheet-view.mjs), [`src/worksheet-ui.mjs`](../src/worksheet-ui.mjs) | 가상화 그리드(`A1`~`XFD1048576`), 인라인 셀 편집기, 수식 입력줄, 시트 탭, 틀 고정, Undo/Redo 스냅샷 관리 |
| **수식 계산 엔진** | [`src/engine.mjs`](../src/engine.mjs), [`src/function-runtime.mjs`](../src/function-runtime.mjs), [`src/function-catalog.mjs`](../src/function-catalog.mjs) | 자체 토크나이저·AST 파서·평가기 + `Formula.js` 4.6.1 호환 래퍼로 379개 함수(별칭 30개 포함) 지원 |
| **워크시트 기능** | [`src/cell-paste.mjs`](../src/cell-paste.mjs), [`src/conditional.mjs`](../src/conditional.mjs), [`src/filters.mjs`](../src/filters.mjs), [`src/formula-tools.mjs`](../src/formula-tools.mjs) | 선택하여 붙여넣기, 동적 조건부 서식, 자동 필터 및 정렬, 수식 참조 색상 하이라이트 및 자동 합계 |
| **차트 & 피벗** | [`src/chart-model.mjs`](../src/chart-model.mjs), [`src/chart-render.mjs`](../src/chart-render.mjs), [`src/pivot.mjs`](../src/pivot.mjs) | 10종·18개 조합 SVG 차트 및 DrawingML 입출력, 다중 필드 피벗 테이블 계산 및 네이티브 피벗 캐시 생성 |
| **파일 입출력 (I/O)** | [`src/io.mjs`](../src/io.mjs), [`src/worksheet-xlsx.mjs`](../src/worksheet-xlsx.mjs), [`src/filter-xlsx.mjs`](../src/filter-xlsx.mjs), [`src/chart-xlsx.mjs`](../src/chart-xlsx.mjs), [`src/pivot-xlsx.mjs`](../src/pivot-xlsx.mjs) | `ExcelJS` + `JSZip` + `@xmldom/xmldom` 기반 `.xlsx`, `.xlsm`, `.csv`, `.tsv` 읽기/쓰기 및 OOXML 파트 후처리 |

---

## 2. 문서 구성 안내

검증 결과는 아래 5개 세부 마크다운 문서로 분리하여 정리했습니다.

1. [**01-reference-and-lifecycle.md**](./01-reference-and-lifecycle.md) — 참조 무결성 및 시트/행/열 수명주기 검증 (14건)
2. [**02-formula-engine-and-coercion.md**](./02-formula-engine-and-coercion.md) — 수식 평가 엔진, 타입 강제 변환 및 함수 경계 조건 검증 (14건)
3. [**03-clipboard-filter-and-merge.md**](./03-clipboard-filter-and-merge.md) — 클립보드, 채우기 핸들, 필터 및 병합 셀 상호작용 검증 (8건)
4. [**04-xlsx-ooxml-compatibility.md**](./04-xlsx-ooxml-compatibility.md) — XLSX/CSV 파일 왕복(Round-Trip) 및 OOXML 호환성 검증 (6건)
5. [**05-ui-state-and-interactions.md**](./05-ui-state-and-interactions.md) — UI 상태 동기화, 틀 고정/확대 좌표계 및 비동기 인터랙션 검증 (6건)

> [!NOTE]
> 본 감사에서는 기존 문서(`README.md`, `IO-NOTES.md`, `docs/*.md`)에 이미 명시된 설계 범위 제한(예: 동적 배열 스필 미지원, 전체 열 참조 `A:A` 미지원 등)과 **실제 구현상의 논리 결함(Implementation Bug)**을 명확히 구분하여 표기했습니다.

---

## 3. 핵심 발견 사항 요약 (Critical & High 우선순위)

총 **48건**의 이슈(Critical 1건, High 16건, Medium 23건, Low 8건)를 실증 확인했습니다. 이 중 데이터 손상이나 오계산을 유발하는 주요 **Critical / High** 항목은 다음과 같습니다.

| 문서 / ID | 심각도 | 분류 | 핵심 증상 요약 | 관련 코드 위치 |
|---|---|---|---|---|
| [**Area 2 #4-1**](./02-formula-engine-and-coercion.md) | **Critical** | 수식 엔진 (`COUNTIF`/`SUMIF`) | `0`·`FALSE`·`"<200"`·`">=0"` 조건이 **빈 셀(`BLANK`)까지 모두 매칭**하며, `"="`·`"<>"`가 빈 셀과 `""`를 구분하지 못하고, `">100"`이 텍스트 `"123"`까지 매칭함 | [`src/engine.mjs:376-401`](../src/engine.mjs#L376-L401) |
| [**Area 2 #4-2**](./02-formula-engine-and-coercion.md) | **High** | 수식 엔진 (`VLOOKUP`/`MATCH`/`XLOOKUP`) | 조회 범위의 빈 셀(`BLANK`)이 검색값 `0`, `FALSE`, `""`와 일치 판정되며, `VLOOKUP(..., TRUE)` 근삿값 모드에서 데이터 하단 빈 셀이 정상 결과를 덮어씀 | [`src/engine.mjs:598-639`](../src/engine.mjs#L598-L639), [`src/function-runtime.mjs:78-90`](../src/function-runtime.mjs#L78-L90) |
| [**Area 2 #1-1, #1-2**](./02-formula-engine-and-coercion.md) | **High** | 타입 강제 변환 (`numeric`/`logical`) | 빈 문자열(`""`)을 `0` 및 `FALSE`로 강제 변환하여 `=""+1` → `1`, `=SUM("", 5)` → `5`, `=IF("", "T", "F")` → `"F"` 반환 (Excel은 모두 `#VALUE!`) | [`src/engine.mjs:33-56`](../src/engine.mjs#L33-L56) |
| [**Area 2 #3-1**](./02-formula-engine-and-coercion.md) | **High** | 날짜 포맷 (`TEXT`) | 1900-01-01~1900-02-29 구간 요일(`dddd`)이 하루씩 밀리고, `=TEXT(60, "mmmm dd, yyyy")`가 `"March 29, 1900"`을 반환함 | [`src/engine.mjs:705-741`](../src/engine.mjs#L705-L741) |
| [**Area 3 #1-1**](./03-clipboard-filter-and-merge.md) | **High** | 필터 / 숨김 행 보호 | 자동 필터가 걸린 상태에서 범위 선택 후 `Delete`, `Ctrl+Enter`, `Ctrl+D`/`R`, 붙여넣기, 채우기 핸들 실행 시 **필터로 숨겨진 행의 데이터까지 모두 덮어쓰거나 삭제함** | [`src/main.mjs:39, 135, 172, 218, 221, 311`](../src/main.mjs#L39), [`src/cell-paste.mjs:19-36`](../src/cell-paste.mjs#L19-L36) |
| [**Area 3 #1-3**](./03-clipboard-filter-and-merge.md) | **High** | `SUBTOTAL` & XLSX 왕복 | XLSX 저장 후 재열기 시 필터 범위 내의 **수동 숨김 행**이 `autoFilter.hiddenRows`로 흡수되어 `=SUBTOTAL(9, ...)` 결과가 달라지고 필터 해제 시 수동 숨김까지 풀림 | [`src/filter-xlsx.mjs:81-84`](../src/filter-xlsx.mjs#L81-L84) |
| [**Area 3 #2-1, #2-2**](./03-clipboard-filter-and-merge.md) | **High** | 병합 셀 (`merges`) | 병합 셀 위에 붙여넣기/채우기 시 비-좌상단 셀에 보이지 않는 유령 데이터가 남아 `SUM` 등을 오염시키며, 일반 정렬(`sortData`)이 병합 셀 검사를 누락함 | [`src/main.mjs:80, 217-222, 236`](../src/main.mjs#L80) |
| [**Area 3 #3-1**](./03-clipboard-filter-and-merge.md) | **High** | 전치 붙여넣기 (`transpose`) | 행/열 바꿈 붙여넣기 시 수식 오프셋을 전치하지 않고 단순 좌표 차(`dr, dc`)만 더해 `A1:B1`(`B1="=A1"`)을 `A1:A2`로 전치하면 `=#REF!`가 발생함 | [`src/cell-paste.mjs:20, 27`](../src/cell-paste.mjs#L20) |
| [**Area 1 #1-2**](./01-reference-and-lifecycle.md) | **High** | 범위 참조 `#REF!` 파싱 | 범위 참조의 시작 셀이 삭제/이동되어 `=SUM(#REF!:B5)` 또는 `=SUM(Sheet1!#REF!:B5)`가 될 때 파서가 `:B5`를 소비하지 못해 `#REF!` 대신 **`#VALUE!`**를 반환함 | [`src/engine.mjs:204, 220`](../src/engine.mjs#L204) |
| [**Area 1 #1-4**](./01-reference-and-lifecycle.md) | **High** | 조건부 서식 행/열 삭제 | 조건부 서식 범위(`B2:B4`)의 첫 행(2행) 삭제 시 기준 셀 상대 수식(`=A2>10`)이 `=#REF!>10`으로 변환되어 남은 범위 전체의 조건부 서식이 영구 파괴됨 | [`src/conditional.mjs:61-66`](../src/conditional.mjs#L61-L66) |
| [**Area 1 #1-8, #2-5**](./01-reference-and-lifecycle.md) | **High** | 피벗 테이블 & XLSX 저장 | 피벗 원본 시트가 삭제되거나 원본 범위의 데이터 행이 모두 삭제되어 1행만 남으면 `writeSpreadsheet` 시 `applyPivot`이 예외를 던져 **XLSX 파일 저장 자체가 차단됨** | [`src/io.mjs:286`](../src/io.mjs#L286), [`src/pivot.mjs:8-13`](../src/pivot.mjs#L8-L13) |
| [**Area 1 #3-1**](./01-reference-and-lifecycle.md) | **High** | 잘라내기·붙여넣기 | 셀을 잘라내어 이동(Cut & Paste)할 때 해당 셀을 참조하던 다른 셀의 수식과 병합 영역이 새 위치로 따라가지 않음 (*문서에 일부 한계 언급됨*) | [`src/main.mjs:218`](../src/main.mjs#L218) |
| [**Area 4 #A-1**](./04-xlsx-ooxml-compatibility.md) | **High** | CSV 왕복 데이터 변질 | CSV 내보내기 시 수식 인젝션 방어용으로 붙인 `'`가 재임포트 시 `value.startsWith("'")`에 의해 `''`로 이중 누적되어 화면에 선행 `'`가 노출됨 | [`src/io.mjs:131, 277`](../src/io.mjs#L131) |
| [**Area 4 #A-2**](./04-xlsx-ooxml-compatibility.md) | **High** | XLSX 날짜 (`t="d"`) | 표준 OOXML ISO 8601 날짜 셀(`t="d"`)을 가져올 때 날짜 시리얼이 아닌 일반 문자열로 저장되어 날짜 함수/서식이 동작하지 않음 | [`src/io.mjs:158, 161`](../src/io.mjs#L158) |
| [**Area 5 #B-1**](./05-ui-state-and-interactions.md) | **High** | UI 편집 & 시트 상태 | 수식 편집 중 다른 시트를 클릭해 참조(`showReferenceSheet`)하는 상태에서 문서 제목 수정 등 `commitEdit()`을 거치지 않는 저장이 일어나면 임시 시트 전환 상태가 영구 저장됨 | [`src/main.mjs:85, 145-147`](../src/main.mjs#L85) |
