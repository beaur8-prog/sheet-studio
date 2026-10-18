# 05. UI 상태 동기화, 틀 고정/확대 좌표계 및 비동기 인터랙션 검증 (UI State & Interactions)

- **검증 대상 파일**: [`src/main.mjs`](../src/main.mjs), [`src/worksheet-view.mjs`](../src/worksheet-view.mjs), [`src/dimension-interactions.mjs`](../src/dimension-interactions.mjs), [`src/chart-interactions.mjs`](../src/chart-interactions.mjs), [`src/pivot-ui.mjs`](../src/pivot-ui.mjs), [`src/styles.css`](../src/styles.css)

---

## 요약 테이블

| ID | 심각도 | 구분 | 파일 및 라인 번호 | 핵심 내용 |
|---|---|---|---|---|
| **B-1** | **High** | UI 편집 / 시트 상태 | [`src/main.mjs:85, 145-147`](../src/main.mjs#L85) | 교차 시트 수식 참조(`showReferenceSheet`)로 `book.activeSheetId`가 임시 변경된 상태에서 `#book-title` 변경 등 `commitEdit()`을 거치지 않는 `checkpoint()/changed()` 호출 시, 임시 시트 상태가 `undo` 및 `localStorage`에 그대로 영구 저장됨. |
| **B-2** | **Medium** | UI 오버레이 / 틀 고정 | [`src/main.mjs:101-107`](../src/main.mjs#L101-L107), [`src/styles.css:3`](../src/styles.css#L3) | 틀 고정(Freeze Panes) 활성 상태에서 비고정 셀 편집 중 스크롤 시 인라인 편집기 `#editor`(`z-index: 8`, `clip-path` 없음)가 고정 행/열 및 머리글(`z-index: 7`) 위로 뚫고 올라와 겹쳐 표시됨. |
| **B-3** | **Medium** | UI 오버레이 / 스크롤바 | [`src/styles.css:5`](../src/styles.css#L5), [`src/main.mjs:74`](../src/main.mjs#L74) | `#formula-ranges`의 `right: 12px; bottom: 12px; overflow: hidden` 때문에 오버레이 스크롤바 환경에서 우측/하단 끝 12px 구간의 수식 참조 점선 박스가 선택 영역(`#selection-extras`)보다 먼저 잘림. |
| **B-4** | **Medium** | UI 차트 / 틀 고정 | [`src/main.mjs:39-43`](../src/main.mjs#L39), [`src/styles.css:1`](../src/styles.css#L1) | 틀 고정(Freeze Panes) 사용 시 차트(`.chart-card`, `z-index: 4`)가 고정 행/열 셀(`z-index: 2`)보다 위에 그려져 스크롤 시 고정 영역을 덮어버림. |
| **B-5** | **Low** | UI 리본 / 편집 | [`src/main.mjs:48, 140`](../src/main.mjs#L48) | 수식(`=10+`) 입력 중 리본의 `자동 합계(autosum)` 클릭 시 미완성 수식을 먼저 커밋(`#VALUE!`)한 뒤 덮어써 불필요한 `undo` 단계가 2개 생성됨. |
| **B-6** | **Low** | UI 리사이즈 툴팁 | [`src/dimension-interactions.mjs:14`](../src/dimension-interactions.mjs#L14) | `zoom != 100%`일 때 열/행 리사이즈 가이드 라인은 화면 픽셀에 표시되나 툴팁 라벨은 비스케일 논리 크기(`px`)를 표시함. |

---

## 1. 셀 편집 중 리본/시트 조작 및 교차 시트 포인트 모드 상태 동기화

- **정상 동작 확인 항목**:
  - 일반 리본 버튼(`bold`, `currency`, `sort-asc`, `filter`, `freeze-menu` 등) 클릭 시 `action(name, event)`에서 `commitEdit()`이 먼저 호출되어 편집 내용이 정상 커밋된 후 명령이 실행됩니다.
  - 글꼴/크기/표시형식 `<select>` 변경 시에도 `applyStyle()` 첫머리에서 `commitEdit()`이 호출됩니다.

### 🚨 [B-1] 교차 시트 수식 참조(`showReferenceSheet`) 중 비-커밋 액션 발생 시 임시 시트 상태 영구 저장 (High)
- **위치**: [`src/main.mjs:85, 145-147`](../src/main.mjs#L85)
- **원인**:
  - 사용자가 `Sheet1!A1`에서 `=SUM(`을 입력하고 하단 시트 탭에서 `Sheet2`를 클릭하면 `showReferenceSheet('Sheet2')`가 호출되어 **`book.activeSheetId`가 임시로 `Sheet2`로 변경**됩니다(편집 완료/취소 시 `restoreEditView()`가 원래 시트인 `Sheet1`로 되돌림).
  - 그러나 이 상태에서 통합 문서 제목 입력창(`#book-title`, [`src/main.mjs:85`](../src/main.mjs#L85))을 수정하는 등 `commitEdit()`을 거치지 않고 `checkpoint(); ... changed();`를 호출하는 경로가 실행되면, **`editSession`의 임시 시트 전환 상태(`activeSheetId = Sheet2`)가 그대로 `undo` 스택과 `localStorage`에 영구 저장**됩니다.

---

### 🚨 [B-5] 미완성 수식 편집 중 리본의 `자동 합계(autosum)` 클릭 시 이중 Undo 생성 (Low)
- **위치**: [`src/main.mjs:48, 140`](../src/main.mjs#L48)
- **현상**: 셀에서 `=10+`를 입력하던 중 리본의 **자동 합계(`autosum`)**를 클릭하면, `action('autosum')`이 먼저 `commitEdit()`을 호출해 미완성 수식 `=10+`(`#VALUE!`)를 커밋하고 체크포인트를 만든 뒤 `startAutoSum()`으로 덮어써서, 실행 취소(`Ctrl+Z`) 시 원래 값으로 한 번에 돌아가지 않고 `#VALUE!` 단계를 거치게 됩니다.

---

## 2. 틀 고정(Freeze Panes) + 확대/축소(`zoom != 100%`) + 스크롤 좌표계 검증

- **정상 동작 확인 항목**:
  - `makeLayout()`([`src/main.mjs:69`](../src/main.mjs#L69))에서 `layout.cols`와 `layout.rows`가 `scale = zoom / 100`이 반영된 누적 픽셀 배열로 계산되며, `paneRects`([`src/worksheet-view.mjs:17-27`](../src/worksheet-view.mjs#L17-L27))가 4개 사분면별로 클리핑된 사각형을 생성하므로 셀 선택 테두리(`#selection`, `#selection-extras`), 복사 점선(`#copy-outlines`), 수식 참조 점선(`#formula-ranges`)의 좌표는 확대/축소 및 스크롤 상태에서 정확히 일치합니다.

### 🚨 [B-2] 인라인 셀 편집기(`#editor`)가 틀 고정(Freeze) 경계 및 행/열 머리글을 침범함 (Medium)
- **위치**: [`src/main.mjs:101-107`](../src/main.mjs#L101-L107), [`src/styles.css:3`](../src/styles.css#L3)
- **원인**:
  - 일반 비고정 셀(`.cell`)은 스크롤 시 고정 행/열 아래로 들어갈 때 `clip-path: inset(${clipY}px 0 0 ${clipX}px)`로 잘리고 고정 셀(`z-index: 2`) 아래에 위치합니다.
  - 반면 인라인 편집기 `#editor`(`z-index: 8`, `#grid-shell`의 직접 자식)는 `positionEditor()`에서 `cellRect(r, c)`로 위치만 설정할 뿐 **`clip-path`를 적용하지 않으며**, 고정 영역 침범 여부도 검사하지 않습니다.
  - 그 결과 틀 고정이 설정된 시트에서 비고정 셀을 편집하면서 스크롤하면 **`#editor` 입력창이 고정 행/열(Frozen Pane)과 행/열 머리글(`#headings`, `z-index: 7`) 위로 뚫고 올라와 겹쳐 표시**됩니다.

---

### 🚨 [B-3] 수식 참조 박스(`#formula-ranges`)의 우측/하단 `12px` 여백으로 인한 조기 클리핑 (Medium)
- **위치**: [`src/styles.css:5`](../src/styles.css#L5), [`src/main.mjs:74`](../src/main.mjs#L74)
- **원인**: `viewRects()`는 가시 영역 크기를 `g.clientWidth - HEADER_W`, `g.clientHeight - HEADER_H`로 계산하지만, `#formula-ranges`의 CSS는 `left: 40px; top: 25px; right: 12px; bottom: 12px; overflow: hidden;`으로 설정되어 있습니다. 오버레이 스크롤바 환경(macOS 기본)에서는 우측 끝 12px과 하단 끝 12px에 걸친 수식 참조 점선 박스가 `#selection-extras`(`inset: 0`)보다 12px 먼저 잘려 나갑니다.

---

### 🚨 [B-4] 틀 고정(Freeze Panes) 사용 시 차트(`.chart-card`)가 고정 행/열 위를 덮음 (Medium)
- **위치**: [`src/styles.css:1`](../src/styles.css#L1), [`src/main.mjs:79`](../src/main.mjs#L79)
- **원인**: `#charts`는 `#surface` 내부에 있어 고정 행/열 구분 없이 전체가 스크롤되며, `.chart-card`의 `z-index: 4`가 고정 행/열 셀의 `z-index: 2`보다 높아서 스크롤 시 차트가 고정된 행/열 위를 덮어버립니다(참고: [`docs/WORKSHEET.md:21`](../docs/WORKSHEET.md#L21)에 고정 경계를 가로지르는 차트의 전체 표시 규칙 미지원이 언급되어 있음).

---

### 🚨 [B-6] `zoom != 100%` 상태에서 열/행 크기 조절 툴팁의 픽셀 표기 불일치 (Low)
- **위치**: [`src/dimension-interactions.mjs:14`](../src/dimension-interactions.mjs#L14)
- **현상**: 확대/축소 배율이 100%가 아닐 때 열/행 경계 드래그 가이드 라인(`position`)은 화면 스케일이 적용된 위치에 표시되지만, 툴팁 라벨(`g.value + 'px'`)은 100% 기준 논리 픽셀을 표시합니다.
