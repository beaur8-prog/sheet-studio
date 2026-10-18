# 04. XLSX/CSV 파일 왕복(Round-Trip) 및 OOXML 호환성 검증 (File I/O & OOXML)

- **검증 대상 파일**: [`src/io.mjs`](../src/io.mjs), [`src/worksheet-xlsx.mjs`](../src/worksheet-xlsx.mjs), [`src/filter-xlsx.mjs`](../src/filter-xlsx.mjs), [`src/chart-xlsx.mjs`](../src/chart-xlsx.mjs), [`src/pivot-xlsx.mjs`](../src/pivot-xlsx.mjs), [`src/conditional.mjs`](../src/conditional.mjs)

---

## 요약 테이블

| ID | 심각도 | 구분 | 파일 및 라인 번호 | 핵심 내용 |
|---|---|---|---|---|
| **A-1** | **High** | CSV 왕복 | [`src/io.mjs:131, 277`](../src/io.mjs#L131), [`src/engine.mjs:258`](../src/engine.mjs#L258) | CSV 내보내기 시 추가된 수식 방어용 `'`가 재임포트 시 `value.startsWith("'")`에 걸려 `''`로 이중 누적됨. 재임포트 후 원래 문자열 앞에 `'`가 노출되어 데이터 변질 발생. 또한 CSV 임포트(`131행`)에서 `-`, `\t`, `\r` 시작 문자열 방어 누락. |
| **A-2** | **High** | XLSX 날짜 | [`src/io.mjs:158, 161`](../src/io.mjs#L158) | 1) OOXML 표준 `t="d"`(ISO 8601 날짜 셀) 임포트 시 `ExcelJS`가 문자열을 반환하여 날짜가 아닌 일반 텍스트로 처리됨.<br>2) `value instanceof Date` 시리얼 변환 시 `Math.round` 누락 및 1900-01-01~1900-02-28 구간 윤년 오프셋 미보정 잠재 위험. |
| **A-3** | **Medium** | XLSX 수식 | [`src/io.mjs:157, 310`](../src/io.mjs#L157), [`src/engine.mjs:684`](../src/engine.mjs#L684) | XLSX 내보내기 시 부착한 `_xlfn.` 접두사(예: `_xlfn.XLOOKUP`)를 `readSpreadsheet`에서 제거하지 않아 재임포트 후 수식 입력줄에 `=_xlfn.XLOOKUP(...)`이 그대로 노출됨. |
| **A-4** | **Medium** | 조건부 서식 | [`src/conditional.mjs:42, 71-79, 88`](../src/conditional.mjs#L42) | 다중 범위 `sqref`(`A1:A5 C1:C5`) 임포트 시 두 번째 범위(`C1:C5`)의 상대 수식 기준점이 전체 `sqref` 좌상단(`A1`)이 아닌 개별 범위 좌상단(`C1`)으로 잘못 계산되어 수식 평가가 틀어짐. (규칙의 수식 변환 저장은 [`WORKSHEET.md:50`](../docs/WORKSHEET.md#L50)에 명시됨) |
| **A-5** | **Medium** | 피벗 XLSX | [`src/io.mjs:323-328`](../src/io.mjs#L323-L328) | 피벗 시트에 페이지 필터(`pageFilters`)가 존재할 때 셀과 행 높이는 `offset`만큼 시프트하지만, 숨김 행(`height===0`)과 `merges`에는 `offset`을 적용하지 않아 엉뚱한 행이 숨겨짐. |
| **A-6** | **Low** | OOXML 순서 | [`src/chart-xlsx.mjs:166`](../src/chart-xlsx.mjs#L166) | `<drawing>` 삽입 시 `legacyDrawing`과 `extLst`만 기준점으로 삼아, `tableParts` 등이 있는 외부 시트에서는 순서 역전 가능성 있음. |

---

## 1. `<worksheet>` 및 `<workbook>` 자식 요소 순서 검증 (ECMA-376 Part 1, §18.3.1.99)

- **기본 내보내기 파이프라인 검증 결과 (정상)**:
  - `ExcelJS.writeBuffer` → `addPivotParts` → `addChartParts` → `addFilterParts` → `addWorksheetParts` 실행 후 생성된 `xl/worksheets/sheet1.xml`의 자식 요소 순서를 검증한 결과:
    `['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'autoFilter', 'mergeCells', 'conditionalFormatting', 'pageMargins', 'pageSetup', 'drawing', 'legacyDrawing']`
    로 ECMA-376 시퀀스 인덱스가 단조 증가(`[0, 1, 2, 3, 4, 5, 10, 14, 16, 20, 21, 29, 30]`)하여 기본 케이스에서는 Excel 복구 경고가 발생하지 않습니다.
- **[A-6] [`src/chart-xlsx.mjs:166`](../src/chart-xlsx.mjs#L166) — `<drawing>` 삽입 위치 잠재 위험 (Low)**:
  ```js
  const before = child(doc.documentElement, 'legacyDrawing') || child(doc.documentElement, 'extLst');
  doc.documentElement.insertBefore(node, before || null);
  ```
  - ECMA-376에서 `<drawing>`(index 29) 뒤에는 `legacyDrawing`, `legacyDrawingHF`, `picture`, `oleObjects`, `controls`, `webPublishItems`, `tableParts`, `extLst`가 올 수 있습니다. `tableParts` 등이 포함된 시트에 차트를 추가하면 `<drawing>`이 `<tableParts>` 뒤에 배치될 수 있습니다.

---

## 2. 🚨 [A-1] CSV 내보내기 후 재가져오기 시 아포스트로피(`'`) 이중 누적으로 인한 데이터 변질 (High)

- **위치**:
  - [`src/io.mjs:277`](../src/io.mjs#L277) (`writeSpreadsheet` CSV 내보내기)
  - [`src/io.mjs:131`](../src/io.mjs#L131) (`readDelimited` CSV 가져오기)
  - [`src/engine.mjs:258`](../src/engine.mjs#L258) (`literal` 평가)
- **원인**:
  1. 셀에 문자열 리터럴 `'+123`, `'=SUM(A1)`, `'-hello`, `'@user`가 저장되어 있을 때, `writeSpreadsheet(..., 'csv')`의 273행 `toScalar(cell?.raw)`가 선행 `'` 1개를 벗겨내어 `value = "+123"` 등을 만듭니다.
  2. 이어서 277행:
     ```js
     if (typeof value === 'string' && /^[=+@\-\t\r]/.test(text)) text = "'" + text;
     ```
     에 의해 CSV 파일에 `'+123`, `'=SUM(A1)`, `'-hello` 형태로 기록됩니다.
  3. 이 CSV 파일을 다시 `readSpreadsheet`(`readDelimited`)로 불러올 때 131행:
     ```js
     const raw = /^[=+@]/.test(value) || /^0\d+/.test(value) || value.startsWith("'") ? "'" + value : value;
     ```
     에서 `value`가 이미 `'`로 시작하므로(`value.startsWith("'")`), 앞에 `'`를 **한 번 더 붙여** `raw = "''+123"`, `raw = "''=SUM(A1)"`, `raw = "''-hello"`로 저장합니다.
  4. 그 결과 `evaluateCell` / `literal(raw)`가 첫 번째 `'`만 제거하므로 화면에 **`'+123`**, **`'=SUM(A1)`**, **`'-hello`**처럼 불필요한 아포스트로피가 노출되며, CSV로 저장하고 열 때마다 `'`가 계속 증식합니다.
  5. 또한 `readDelimited` 131행은 `/^[=+@]/.test(value)`만 검사하고 `-`로 시작하는 수식 형태 문자열(`-1+1`)은 이스케이프하지 않는 불일치가 있습니다.

---

## 3. 🚨 [A-2] OOXML `t="d"` (ISO 8601 날짜 셀) 가져오기 미지원 및 1900 윤년 오프셋 (High)

- **위치**: [`src/io.mjs:158, 161`](../src/io.mjs#L158)
- **원인**:
  1. **ISO 8601 날짜(`t="d"`) 처리 누락**: ECMA-376 §18.18.11 표준인 `t="d"` 셀(예: `<c r="B1" t="d"><v>2024-01-15T00:00:00.000Z</v></c>`)을 `ExcelJS`로 읽으면 `cell.value`가 `Date` 인스턴스가 아닌 문자열 `"2024-01-15T00:00:00.000Z"`로 반환됩니다. `readSpreadsheet`는 이를 일반 문자열로 저장하고 `style.format = 'date'`도 설정하지 않아 날짜 계산이나 날짜 서식이 적용되지 않습니다.
  2. **`Date` → 시리얼 변환 시 부동소수점 반올림 누락**: [`src/io.mjs:158`](../src/io.mjs#L158)의 `String(value.getTime() / DAY + 25569)`는 `ExcelJS`가 만든 `Date`의 밀리초 반올림 오차로 인해 정수가 아닌 `.999999999` 같은 소수점 시리얼을 만들 수 있습니다.

---

## 4. 🚨 [A-3] 최신 함수 `_xlfn.` 접두사 XLSX 왕복 시 수식 입력줄 노출 (Medium)

- **위치**: [`src/io.mjs:157, 310`](../src/io.mjs#L157), [`src/engine.mjs:684-693`](../src/engine.mjs#L684-L693)
- **현상**: `=XLOOKUP("Alpha", A2:A3, B2:B3)`를 XLSX로 내보낼 때 `formulaForXlsx`가 `<f>_xlfn.XLOOKUP(...)</f>`로 변환하여 저장하지만, 다시 `readSpreadsheet`로 읽어올 때는 `_xlfn.`을 제거하지 않고 `raw = '=' + cell.formula`로 저장합니다. 계산은 정상 수행되지만 사용자가 셀이나 수식 입력줄을 편집할 때 `=_xlfn.XLOOKUP(...)`이 그대로 노출됩니다.

---

## 5. 🚨 [A-4] 다중 범위(`sqref="A1:A5 C1:C5"`) 조건부 서식 가져오기 시 상대 수식 기준점 오류 (Medium)

- **위치**: [`src/conditional.mjs:42, 71-79`](../src/conditional.mjs#L42)
- **원인**:
  - Excel에서 `<conditionalFormatting sqref="A1:A5 C1:C5">`의 `<formula>A1=1</formula>`는 전체 `sqref`의 첫 번째 범위 좌상단(`A1`)을 기준으로 작성된 상대 참조이므로, `C1` 셀을 평가할 때는 `C1=1`로 시프트되어야 합니다.
  - 그러나 `importConditional`(`src/conditional.mjs:71`)이 `group.ref.split(/\s+/)`로 분리하면서 두 번째 규칙을 `range: "C1:C5"`, `formulae: ["A1=1"]`로 그대로 복제합니다.
  - 이후 `createConditionalEvaluator`([`src/conditional.mjs:42`](../src/conditional.mjs#L42))는 `C1:C5`의 좌상단인 `C1`(`b.r1=0, b.c1=2`)을 기준으로 `shiftFormula('=A1=1', r - 0, c - 2)`를 수행하여, **`C1` 셀에서 `C1=1`이 아니라 `A1=1`을 평가하는 오동작**이 발생합니다.

---

## 6. 🚨 [A-5] 피벗 테이블 페이지 필터(`pageFilters`) 내보내기 시 숨김 행·병합 셀 오프셋 누락 (Medium)

- **위치**: [`src/io.mjs:299-328`](../src/io.mjs#L299-L328)
- **원인**: 피벗 테이블에 `pageFilters`가 있을 때 `writeSpreadsheet`는 셀 데이터, 행 높이(`320행`), 조건부 서식(`300행`)을 `offset = pageFilters.length + 1`만큼 아래로 밀어서 기록하지만, **323~328행의 숨김 행(`height === 0`), 자동 필터 숨김 행, 병합 셀(`sheet.merges`)에는 `offset`을 더하지 않습니다**:
  ```js
  for (const [r, height] of Object.entries(sheet.rowHeights || {})) if (height === 0) target.getRow(+r + 1).hidden = true; // offset 누락!
  ```
  이로 인해 XLSX 저장 시 엉뚱한 상단 행이 숨겨지거나 병합 위치가 어긋납니다.
