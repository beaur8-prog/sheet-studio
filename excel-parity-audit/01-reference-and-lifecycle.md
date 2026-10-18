# 01. 참조 무결성 및 시트/행/열 수명주기 검증 (Reference & Lifecycle)

- **검증 대상 파일**: [`src/main.mjs`](../src/main.mjs), [`src/engine.mjs`](../src/engine.mjs), [`src/conditional.mjs`](../src/conditional.mjs), [`src/filters.mjs`](../src/filters.mjs), [`src/chart-model.mjs`](../src/chart-model.mjs), [`src/pivot.mjs`](../src/pivot.mjs), [`src/formula-tools.mjs`](../src/formula-tools.mjs), [`src/io.mjs`](../src/io.mjs)

---

## 요약 테이블

| ID | 구분 | 요약 | 파일 및 라인 | 심각도 | 유형 |
|---|---|---|---|---|---|
| **Bug 1-2** | 행/열 삽입·삭제 | 범위 참조의 시작 셀이 `#REF!`가 될 때(`#REF!:B5`, `'Sheet1'!#REF!:B5`) 파서가 `#REF!` 대신 `#VALUE!` 반환 | [`src/engine.mjs:204, 220, 790-794`](../src/engine.mjs#L204) | **High** | 구현 결함 |
| **Bug 1-4** | 조건부 서식 | 조건부 서식 범위의 첫 행/열 삭제 시 기준 셀 상대 수식(`=A2>10`)이 `=#REF!>10`으로 파괴됨 | [`src/conditional.mjs:61-66`](../src/conditional.mjs#L61-L66) | **High** | 구현 결함 |
| **Bug 1-8** | 피벗 테이블 | 행 삭제로 피벗 `sourceRange`가 1행(`A1:B1`)으로 축소되면 XLSX 저장 시 전체 저장이 실패함 | [`src/pivot.mjs:13, 96-102`](../src/pivot.mjs#L13), [`src/io.mjs:286`](../src/io.mjs#L286) | **High** | 구현 결함 |
| **Bug 2-5** | 시트 삭제 / 피벗 | 피벗 원본 시트 삭제 후 통합 문서 XLSX 저장 시 `applyPivot` 예외로 저장이 완전히 차단됨 | [`src/main.mjs:293`](../src/main.mjs#L293), [`src/io.mjs:286`](../src/io.mjs#L286) | **High** | 구현 결함 |
| **Bug 3-1** | 잘라내기·붙여넣기 | 셀을 잘라내어 이동(Cut & Paste)해도 해당 셀을 참조하는 다른 셀의 수식 및 병합 영역이 새 위치로 갱신되지 않음 | [`src/main.mjs:218`](../src/main.mjs#L218), [`src/cell-paste.mjs:27`](../src/cell-paste.mjs#L27) | **High** | 설계 한계 (`WORKSHEET.md` 명시) |
| **Bug 1-1** | 행/열 삽입·삭제 | 여러 행/열을 선택하고 삽입·삭제해도 항상 `selection.r` / `selection.c` 1개만 처리됨 | [`src/main.mjs:238`](../src/main.mjs#L238) | **Medium** | Excel 동작 차이 |
| **Bug 1-3** | 병합 셀 | 2셀 병합(`A1:A2`)에서 1행 삭제 시 1×1 병합(`A1:A1`)이 해제되지 않고 `merges`에 잔류 | [`src/main.mjs:238`](../src/main.mjs#L238) | **Medium** | 구현 결함 |
| **Bug 1-5** | 조건부 서식 | 다른 시트에서 행/열 삽입·삭제 시 교차 시트 참조 조건부 서식 수식이 업데이트되지 않음 | [`src/main.mjs:238`](../src/main.mjs#L238), [`src/conditional.mjs:64`](../src/conditional.mjs#L64) | **Medium** | 구현 결함 |
| **Bug 1-7** | 차트 | 차트 범위가 1행(`A1:B1`)으로 축소될 때 `catRef`/`valRef`가 `A2:A1`(역전 범위)이 되어 잘못된 값 표시 | [`src/chart-model.mjs:38`](../src/chart-model.mjs#L38) | **Medium** | 구현 결함 |
| **Bug 2-1** | 시트 복제 | 시트 복제 시 자기 시트 이름을 명시한 수식(`=Sales!B2`)이 복제 시트로 전환되지 않고 원본 시트를 가리킴 | [`src/main.mjs:292`](../src/main.mjs#L292) | **Medium** | 구현 결함 |
| **Bug 2-2** | 시트 생성·복제 | 대소문자만 다른 시트 이름 존재 시 `addSheet`/`duplicate-sheet`가 중복을 감지하지 못해 검증 에러 유발 | [`src/main.mjs:204, 292`](../src/main.mjs#L204), [`src/engine.mjs:941`](../src/engine.mjs#L941) | **Medium** | 구현 결함 |
| **Bug 2-3** | 특수 시트 이름 | 숫자로 시작하는 시트 이름을 따옴표 없이 `=2026년!A1`로 쓰면 `#VALUE!` 및 참조 인식 오류 발생 | [`src/engine.mjs:117-119`](../src/engine.mjs#L117-L119), [`src/formula-tools.mjs:10`](../src/formula-tools.mjs#L10) | **Medium** | 구현 결함 |
| **Bug 2-4** | 시트 이름 변경 | 시트 이름 변경 시 `workbook.chartNames` 수식이 갱신되지 않아 이후 행/열 삽입·삭제가 무시됨 | [`src/engine.mjs:855-864`](../src/engine.mjs#L855-L864), [`src/chart-model.mjs:59`](../src/chart-model.mjs#L59) | **Medium** | 구현 결함 |
| **Bug 1-6** | 자동 필터 | 필터의 모든 데이터 행이 삭제되어 헤더 1행(`r1 === r2`)만 남아도 빈 `autoFilter`가 유지됨 | [`src/filters.mjs:128-132`](../src/filters.mjs#L128-L132) | **Low** | 구현 결함 |

---

## 1. 행/열 삽입 및 삭제 (`insertRow`, `deleteRow`, `insertColumn`, `deleteColumn`)

행/열 구조 변경은 [`structural(axis, delta)`](../src/main.mjs#L238) 및 [`transformFormulaReferences`](../src/engine.mjs#L771-L826)에서 처리됩니다.

### 정상 동작 확인 항목
- 모든 시트(`for (const owner of book.sheets)`)를 순회하며 수정 대상 시트(`s.name`)를 가리키는 셀 수식을 업데이트합니다.
- 단일 셀 `=A3`에서 3행(`index=2, delta=-1`) 삭제 시 `=#REF!`가 되며, 범위 `=SUM(A1:A5)`에서 1행 또는 5행 삭제 시 `=SUM(A1:A4)`로 정상 축소됩니다.
- 1행짜리 범위 `=SUM(A1:C1)`에서 1행 삭제 시 전체 범위 토큰이 `#REF!`로 치환되어 `=SUM(#REF!)`가 됩니다.

---

### 🚨 [Bug 1-1] 다중 행/열 선택 시에도 항상 1개 행/열만 삽입·삭제됨 (Medium)
- **위치**: [`src/main.mjs:238`](../src/main.mjs#L238)
- **코드**:
  ```js
  const index = axis === 'r' ? selection.r : selection.c;
  ```
- **현상 vs Excel**: 사용자가 3개 행(`r=2..4`)을 드래그하여 선택하고 '행 삽입' 또는 '행 삭제'를 실행해도 `bounds()`(`r1..r2`)가 아닌 `selection.r` 단일 행만 삽입/삭제됩니다. Excel은 선택된 행/열 개수만큼 일괄 삽입/삭제합니다.

---

### 🚨 [Bug 1-2] 범위 수식의 한쪽 끝이 `#REF!`가 될 때 파서가 `#REF!` 대신 `#VALUE!` 반환 (High)
- **위치**: [`src/engine.mjs:204, 220`](../src/engine.mjs#L204), [`src/engine.mjs:790-794`](../src/engine.mjs#L790-L794)
- **원인**:
  1. `Sheet2`에 `=SUM(A1048575:A1048576)` 입력 후 1행 삽입(`index=0, delta=1`) 시 `=SUM(A1048576:#REF!)`가 생성됩니다.
  2. 또한 수식 이동(`shiftFormula`) 등으로 범위의 첫 셀이 `#REF!`가 되어 `=SUM(#REF!:B5)` 또는 `=SUM('Sheet1'!#REF!:B5)`가 될 때, [`Parser.primary()`](../src/engine.mjs#L204)에서 첫 에러 토큰을 만나자마자 뒤따르는 `:B5`를 소비하지 않고 즉시 반환합니다.
  3. 그 결과 `Parser.parse()` 또는 함수 인자 루프에서 잔여 토큰 `:B5` 때문에 `#REF!`가 아닌 **`#VALUE!`**를 반환합니다.

---

### 🚨 [Bug 1-3] 2셀 병합(`A1:A2`)에서 1개 행 삭제 시 단일 셀 병합(`A1:A1`)이 제거되지 않고 잔류 (Medium)
- **위치**: [`src/main.mjs:238`](../src/main.mjs#L238)
- **코드**:
  ```js
  s.merges = (s.merges || []).map(m => transformRefs(m, axis, index, delta, s.name, s.name)).filter(m => !m.includes('#REF!'));
  ```
- **재현**: `A1:A2`를 병합한 뒤 1행(`index=0, delta=-1`) 또는 2행(`index=1, delta=-1`) 삭제.
- **실증 결과**: `transformRefs("A1:A2", 'r', 0, -1)` → `"A1:A1"`. `!m.includes('#REF!')` 필터를 통과하여 `s.merges`에 `['A1:A1']`로 남습니다.
- **영향**: 단일 셀이 병합 상태로 취급되며 XLSX 내보내기 시 `<mergeCell ref="A1:A1"/>`이 기록되어 Excel에서 파일 복구 경고를 유발할 수 있습니다. `a !== b` (시작과 끝 주소가 다른 경우만 유지) 필터가 필요합니다.

---

### 🚨 [Bug 1-4] 조건부 서식 적용 범위의 첫 번째 행/열 삭제 시 상대 참조 수식이 `=#REF!`로 파괴됨 (High)
- **위치**: [`src/conditional.mjs:58-67`](../src/conditional.mjs#L58-L67)
- **코드**:
  ```js
  sheet.conditionalRules = (sheet.conditionalRules || []).flatMap(rule => {
    const range = transformFormulaReferences(rule.range, { axis, index, delta, targetSheet: sheet.name, ownerSheet: sheet.name });
    if (range.includes('#REF!')) return [];
    const formulae = rule.formulae.map(f => transformFormulaReferences(f, { axis, index, delta, targetSheet: sheet.name, ownerSheet: sheet.name }));
    return [{ ...rule, range, formulae }];
  });
  ```
- **원인**: 조건부 서식의 수식(`rule.formulae`)은 적용 범위(`rule.range`)의 좌상단 셀(`b.r1, b.c1`)을 기준으로 매 평가 시마다 [`shiftFormula('='+f, r-b.r1, c-b.c1)`](../src/conditional.mjs#L42)로 오프셋 이동되어 평가됩니다.
  그런데 `B2:B4` 범위에 `=A2>10` 규칙이 있을 때 **2행(`index=1, delta=-1`)을 삭제**하면:
  - `rule.range`는 `B2:B4` → `B2:B3`으로 축소됩니다(기존 3행이 새 2행 `B2`가 됨).
  - 반면 `rule.formulae[0]`(`=A2>10`)에는 단일 셀 참조 `A2`(`r=1`)가 들어 있으므로 `transformFormulaReferences`가 삭제된 행(`index=1`)과 일치한다고 판단해 **`=#REF!>10`**으로 만들어 버립니다.
- **실증 결과**: 2행 삭제 후 남은 `B2:B3` 범위 전체에서 조건부 서식이 영구적으로 작동 불능이 됩니다.

---

### 🚨 [Bug 1-5] 다른 시트에서 행/열 삽입·삭제 시 교차 시트를 참조하는 `conditionalRules.formulae` 미갱신 (Medium)
- **위치**: [`src/main.mjs:238`](../src/main.mjs#L238)
- **원인**: `structural()`에서 셀 수식은 모든 시트를 순회하지만 `shiftWorksheetFeatures(s, axis, index, delta)`는 현재 시트 `s`에 대해서만 호출됩니다.
- **재현**: `Sheet1`의 조건부 서식 수식이 `=Sheet2!$A$5>100`일 때 `Sheet2`에서 1행을 삽입해도 `Sheet1`의 조건부 서식 수식은 `=Sheet2!$A$5>100`에 그대로 머무릅니다(Excel은 `=Sheet2!$A$6>100`으로 이동).

---

### 🚨 [Bug 1-6] 필터 범위의 모든 데이터 행을 삭제해도 헤더만 남은 `r1 === r2` 상태의 빈 `autoFilter`가 유지됨 (Low)
- **위치**: [`src/filters.mjs:127-140`](../src/filters.mjs#L127-L140)
- **재현**: `r1: 0, r2: 1`(1행 헤더, 2행 데이터)인 필터에서 2행을 삭제하면 `b.r2`가 `0`이 되어 `r1: 0, r2: 0`이 됩니다. `b[hi] < b[lo]`만 검사하므로 데이터 행이 0개인 상태로 남습니다.

---

### 🚨 [Bug 1-7] 차트 범위(`chart.range`)에서 행 삭제 시 1행(`A1:B1`)으로 축소되면 역전된 범위(`A2:A1`)를 생성함 (Medium)
- **위치**: [`src/chart-model.mjs:37-43`](../src/chart-model.mjs#L37-L43)
- **재현**: `A1:B2` 범위를 가지는 차트에서 2행(`index=1, delta=-1`)을 삭제하면 `chart.range`는 `'A1:B1'`(`a.r=0, b.r=0`)이 됩니다. 이후 `chartSeries()` 호출 시 `r1 = a.r + 1 = 1`, `r2 = b.r = 0`이 되어 `catRef.range`가 **`'A2:A1'`**, `valRef.range`가 **`'B2:B1'`**로 역전되며, `refValues()`가 `Math.min`~`Math.max`로 순회하여 헤더 행과 2행을 데이터 포인트로 잘못 렌더링합니다.

---

### 🚨 [Bug 1-8] 피벗 원본 범위(`pivot.sourceRange`)의 데이터 행이 모두 삭제되어 1행(`A1:B1`)만 남으면 XLSX 저장이 실패함 (High)
- **위치**: [`src/pivot.mjs:13, 96-102`](../src/pivot.mjs#L13), [`src/io.mjs:286`](../src/io.mjs#L286)
- **재현**: `Sheet1`의 `A1:B2`를 원본으로 하는 피벗 테이블이 있을 때 `Sheet1`에서 2행을 삭제하면 `sourceRange`가 `'A1:B1'`로 축소됩니다. 이후 통합 문서를 XLSX로 저장(`writeSpreadsheet`)하려고 하면 [`src/io.mjs:286`](../src/io.mjs#L286)에서 `applyPivot`이 자동 실행되고, `sourceData()`의 `if (a.r >= b.r) fail(...)`에서 예외가 발생하여 **통합 문서의 XLSX 저장 자체가 차단**됩니다.

---

## 2. 시트 수명주기 (이름 변경, 삭제, 복제)

### 🚨 [Bug 2-1] 시트 복제 시 자기 자신의 시트 이름을 명시한 셀/조건부 서식 수식(`=Sales!A1`)이 원본 시트에 고정됨 (Medium)
- **위치**: [`src/main.mjs:292`](../src/main.mjs#L292)
- **현상**: `Sales` 시트에 `=Sales!B2*2`가 있을 때 시트를 복제하여 `Sales (2)`를 만들면, 차트의 `sheetId`는 새 시트로 바뀌지만 셀 수식과 조건부 서식 수식은 `='Sales (2)'!B2*2`로 바뀌지 않고 원본 `=Sales!B2*2`를 그대로 가리킵니다.

---

### 🚨 [Bug 2-2] 대소문자만 다른 시트 이름 존재 시 `addSheet` / `duplicate-sheet`가 검증 에러 유발 (Medium)
- **위치**: [`src/main.mjs:204, 292`](../src/main.mjs#L204), [`src/engine.mjs:941`](../src/engine.mjs#L941)
- **원인**: `validateWorkbook()`과 `rename-sheet`는 시트 이름 중복을 `toLowerCase()`로 대소문자 구분 없이 검사하지만, `addSheet()`와 `duplicate-sheet`는 `===`로만 중복을 검사합니다. `'sheet1'`(소문자)이 있을 때 새 시트를 추가해 `'Sheet1'`이 생성되면 `validateWorkbook()`에서 예외가 발생합니다.

---

### 🚨 [Bug 2-3] 숫자로 시작하는 따옴표 없는 시트 참조(`=2026년!A1`) 파싱 실패 (`#VALUE!`) (Medium)
- **위치**: [`src/engine.mjs:117-119`](../src/engine.mjs#L117-L119), [`src/formula-tools.mjs:10`](../src/formula-tools.mjs#L10)
- **원인**: 토크나이저가 숫자 리터럴(`number`)을 식별자(`word`)보다 먼저 매칭합니다. 시트 이름이 `2026년`일 때 수식에 `=2026년!A1`을 입력하면 `2026`(number), `년`(word), `!`(op), `A1`(word)으로 쪼개져 `#VALUE!`가 발생하며, `formulaReferences("=2026년!A1")`은 시트 이름을 `'년'`으로 잘못 인식합니다.

---

### 🚨 [Bug 2-4] 시트 이름 변경 시 `workbook.chartNames`의 수식이 갱신되지 않음 (Medium)
- **위치**: [`src/engine.mjs:855-864`](../src/engine.mjs#L855-L864), [`src/chart-model.mjs:59`](../src/chart-model.mjs#L59)
- **원인**: `rewriteWorkbookSheetReferences()`가 `sheet.cells`와 `sheet.conditionalRules`만 갱신하고 `workbook.chartNames`는 갱신하지 않습니다. 시트 이름을 변경한 뒤 행/열을 삽입·삭제하면 `updateChartReferences()`가 새 시트 이름으로 매칭을 시도하므로 `chartNames`의 참조 이동이 누락됩니다.

---

### 🚨 [Bug 2-5] 피벗 테이블의 원본 시트(`sourceSheetId`)를 삭제하면 XLSX 내보내기가 영구 실패함 (High)
- **위치**: [`src/main.mjs:293`](../src/main.mjs#L293), [`src/io.mjs:286`](../src/io.mjs#L286), [`src/pivot.mjs:8-9`](../src/pivot.mjs#L8-L9)
- **재현**: `Sheet1`을 원본으로 하는 피벗 시트 생성 후 `Sheet1`을 삭제합니다. 이후 `writeSpreadsheet(book, 'xlsx')`를 호출하면 `applyPivot` → `sourceData()`에서 `'원본 시트가 없습니다. 원본을 다시 선택하세요.'` 에러가 발생하여 피벗 시트를 지우거나 일반 범위로 변환하기 전까지 XLSX 저장이 불가능해집니다.

---

## 3. 잘라내기 및 붙여넣기 (Cut & Paste)

### 🚨 [Bug 3-1] 셀을 잘라내어 이동(Cut & Paste)할 때 해당 셀을 참조하는 다른 셀의 수식이 업데이트되지 않음 (High — 설계 한계)
- **위치**: [`src/main.mjs:218`](../src/main.mjs#L218), [`src/cell-paste.mjs:27`](../src/cell-paste.mjs#L27)
- **현상**:
  - 수식 셀 자체를 잘라내어 붙여넣을 때는 `!clip.cut` 조건에 의해 상대 참조가 고정된 채로 이동하여 Excel과 동일하게 동작합니다.
  - 그러나 `A1=10`, `B1="=A1+5"` 상태에서 `A1`을 잘라내어 `E5`에 붙여넣으면, `B1`의 수식이 `=E5+5`로 갱신되지 않고 `=A1+5`로 남아 결과가 `5`로 바뀝니다. 또한 잘라낸 영역 내의 병합 셀(`merges`)도 이동하지 않습니다.
  - 참고: [`docs/WORKSHEET.md:32`](../docs/WORKSHEET.md#L32)에 *"Excel의 모든 이동 참조 재작성, 조건부 규칙·병합·외부 앱의 HTML 클립보드 복제는 포함하지 않습니다"*로 한계가 명시되어 있으나, 사용자가 Excel과 가장 크게 차이를 체감하는 부분입니다.
