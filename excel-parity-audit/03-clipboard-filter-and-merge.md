# 03. 클립보드, 채우기 핸들, 필터 및 병합 셀 상호작용 검증 (Clipboard, Filter & Merge)

- **검증 대상 파일**: [`src/cell-paste.mjs`](../src/cell-paste.mjs), [`src/main.mjs`](../src/main.mjs), [`src/filters.mjs`](../src/filters.mjs), [`src/filter-xlsx.mjs`](../src/filter-xlsx.mjs), [`src/worksheet-ui.mjs`](../src/worksheet-ui.mjs), [`src/formula-tools.mjs`](../src/formula-tools.mjs), [`src/engine.mjs`](../src/engine.mjs)

---

## 요약 테이블

| 번호 | 구분 | 핵심 결함 요약 | 위치 | 심각도 |
|---|---|---|---|---|
| **1-1** | 필터 / 숨김 행·열 | `Delete`/`Backspace`, `Ctrl+Enter`, `Ctrl+D`/`R`, 채우기 핸들 드래그, 붙여넣기(`Ctrl+V`), 서식 변경이 **필터로 숨겨진 행(`filteredRows`) 및 수동 숨김 행·열을 모두 덮어씀** (데이터 손상 위험) | [`src/main.mjs:39, 135, 172, 218, 221, 311`](../src/main.mjs#L39), [`src/cell-paste.mjs:19-36`](../src/cell-paste.mjs#L19-L36) | **High** |
| **1-3** | `SUBTOTAL` & XLSX 왕복 | 메모리 내에서는 `SUBTOTAL(1~11)`과 `SUBTOTAL(101~111)`이 정상 구분되나, **XLSX 저장 후 다시 열 때(`restoreFilterParts`) 자동 필터 범위 안의 수동 숨김 행이 `autoFilter.hiddenRows`로 편입되고 `rowHeights[r]=0`이 삭제되어 `SUBTOTAL(1~11)`에서도 제외됨** | [`src/engine.mjs:487-503`](../src/engine.mjs#L487-L503), [`src/filter-xlsx.mjs:81-84`](../src/filter-xlsx.mjs#L81-L84) | **High** |
| **2-1** | 병합 셀 복사/붙여넣기/편집 | 병합 셀 복사 시 `merges` 정보가 복사되지 않으며, 기존 병합 셀의 일부 또는 전체 위에 붙여넣기/채우기/`Ctrl+Enter` 실행 시 **비-좌상단(하위) 셀에 화면에 보이지 않는 유령 데이터(phantom cell)가 기록되어 `SUM` 등 수식 결과를 오염시킴** | [`src/main.mjs:80, 91, 217-222`](../src/main.mjs#L80), [`src/cell-paste.mjs:19-36`](../src/cell-paste.mjs#L19-L36) | **High** |
| **2-2** | 병합 셀 정렬 / 필터 | 자동 필터 정렬(`sortFilterRange`)은 병합 셀을 차단하지만, **일반 정렬(`sortData`)은 병합 셀 검사를 전혀 하지 않아 병합 범위와 행 데이터를 뒤섞음**. 또한 수직 병합 셀의 좌상단 행이 필터로 숨겨져도 하위 가시 행에 좌상단 셀이 렌더링됨 | [`src/main.mjs:79-80, 236`](../src/main.mjs#L79-L80), [`src/filters.mjs:112`](../src/filters.mjs#L112) | **High** |
| **3-1** | 행/열 바꿈(`transpose`) 수식 보정 | `transpose: true` 붙여넣기 시 수식의 상대 오프셋(`dr`, `dc`)을 전치하거나 혼합 참조(`$A1` ↔ `A$1`)를 변환하지 않고 단순 위치 차이(`dest - orig`)만 `shiftFormula`에 적용하여 **엉뚱한 셀을 참조하거나 `#REF!`가 발생** | [`src/cell-paste.mjs:20, 27`](../src/cell-paste.mjs#L20), [`src/engine.mjs:866-886`](../src/engine.mjs#L866-L886) | **High** |
| **1-2** | 복사 (`Ctrl+C`) | 자동 필터로 숨겨진 행(`filteredRows`)은 복사에서 제외하나, **수동으로 숨긴 행(`rowHeights[r]===0`)과 숨긴 열(`columnWidths[c]===0`)은 복사에서 제외하지 않음** | [`src/main.mjs:217`](../src/main.mjs#L217) | **Medium** |
| **3-2** | 범위 참조의 경계 밖 이동 | 범위 참조(`=SUM(A1:B2)`)를 위/왼쪽으로 복사해 시작점이 경계 밖으로 나가면 `=SUM(#REF!:A1)`이 되고 파서가 `#REF!` 대신 `#VALUE!`를 반환 | [`src/engine.mjs:204, 220, 866-886`](../src/engine.mjs#L204) | **Medium** |
| **3-3** | 다중 셀 반복 붙여넣기 & 채우기 핸들 | 역방향(아래→위, 오른쪽→왼쪽) 다중 선택 후 반복 붙여넣기 시 시작점이 `bounds().r1, c1`이 아닌 `selection.r, c`로 설정됨. 채우기 핸들은 다중 셀 선택 시에도 단일 시작 셀만 복제함 | [`src/main.mjs:218, 309-311`](../src/main.mjs#L218), [`src/cell-paste.mjs:16-22`](../src/cell-paste.mjs#L16-L22) | **Medium** |

---

## 1. 필터링된 행 / 수동 숨김 행·열 상호작용

### 🚨 [1-1] 숨김/필터 행을 가로지르는 범위 편집 시 데이터 덮어쓰기 (High)
Excel에서는 자동 필터로 숨겨진 행이 포함된 범위를 선택하고 `Delete`(내용 지우기), `Ctrl+Enter`(선택 범위 채우기), `Ctrl+D`/`Ctrl+R`(아래/오른쪽 채우기), 채우기 핸들 드래그를 수행하면 **화면에 보이는(가시) 셀만 수정**되고 필터로 숨겨진 행은 보호됩니다. 반면 Sheet Studio는 모든 범위 편집 루프에서 `filteredRows` 및 수동 숨김 여부를 확인하지 않고 숨겨진 행까지 덮어씁니다.

1. **`Delete` / `Backspace` / 서식 지정 (`forSelection`)** — [`src/main.mjs:39, 172`](../src/main.mjs#L39):
   ```js
   function forSelection(fn){const b=bounds();...;for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)fn(address(r,c),r,c);return true;}
   ```
   - 필터링 상태에서 범위를 선택하고 `Delete`를 누르면 **필터로 숨겨진 행의 데이터까지 모두 삭제**됩니다. (반면 하단 상태 표시줄 통계 계산인 `updateStats` [`src/main.mjs:94`](../src/main.mjs#L94)에서는 `!filteredRows?.has(r)&&sheet().rowHeights?.[r]!==0`로 숨김 행을 제외하고 있어 UI 표시와 실제 삭제 동작이 불일치합니다.)
2. **`Ctrl+Enter` 선택 범위 일괄 입력 (`commitEdit(null, true)`)** — [`src/main.mjs:135`](../src/main.mjs#L135):
   - 숨겨진 행(`filteredRows`)을 건너뛰지 않고 `b.r1..b.r2` 전체에 값을 덮어씁니다.
3. **`Ctrl+D` / `Ctrl+R` (`fillDown`) 및 채우기 핸들 (`fillStart`)** — [`src/main.mjs:221, 311`](../src/main.mjs#L221):
   - 필터로 숨겨진 행을 건너뛰지 않고 모두 덮어씁니다.
4. **붙여넣기 (`Ctrl+V` / `pasteText` / `pastePlan`)** — [`src/main.mjs:218`](../src/main.mjs#L218), [`src/cell-paste.mjs:19-36`](../src/cell-paste.mjs#L19-L36):
   - 가시 행 2개를 복사한 뒤 필터가 걸린 영역에 붙여넣으면, 두 번째 행이 다음 가시 행이 아니라 바로 아래의 **숨겨진 행(`start.r + 1`)에 덮어써지는 데이터 손상**이 발생합니다.

---

### 🚨 [1-2] 수동으로 숨긴 행·열이 복사(`Ctrl+C`)에서 제외되지 않음 (Medium)
- **위치**: [`src/main.mjs:217`](../src/main.mjs#L217)
- **현상**: `if (!cut && filteredRows?.has(r)) continue;`로 자동 필터 숨김 행은 복사에서 제외하지만, 수동으로 숨긴 행(`rowHeights[r] === 0`)과 수동으로 숨긴 열(`columnWidths[c] === 0`)은 그대로 클립보드에 포함되어 복사됩니다.

---

### 🚨 [1-3] `SUBTOTAL`의 XLSX 왕복 시 수동 숨김 행이 필터 숨김 행으로 변질됨 (High)
- **위치**: [`src/engine.mjs:487-503`](../src/engine.mjs#L487-L503), [`src/filter-xlsx.mjs:81-84`](../src/filter-xlsx.mjs#L81-L84)
- **코드**:
  ```js
  // src/filter-xlsx.mjs:81-84
  const hiddenRows = all(child(root, 'sheetData'), 'row').filter(n => attr(n, 'hidden') === '1').map(n => Number(attr(n, 'r')) - 1).filter(r => r > range.r1 && r <= range.r2);
  s.autoFilter = { range, columns, hiddenRows, ...(unsupported.length ? { unsupported } : {}) };
  const rowNodes = new Map(all(child(root, 'sheetData'), 'row').map(n => [Number(attr(n, 'r')) - 1, n]));
  for (const r of hiddenRows) { const height = Number(attr(rowNodes.get(r), 'ht')); if (height) s.rowHeights[r] = height * 4 / 3; else delete s.rowHeights[r]; }
  ```
- **원인 및 실증 결과**:
  - 메모리 내에서는 `SUBTOTAL(9, ...)`가 필터 숨김 행만 제외하고 수동 숨김 행(`rowHeights[r] === 0`)은 포함하며, `SUBTOTAL(109, ...)`는 둘 다 제외하여 Excel과 일치합니다.
  - 그러나 XLSX로 저장했다가 다시 열면, `restoreFilterParts`가 자동 필터 범위 내의 모든 `<row hidden="1">` 행을 무조건 `autoFilter.hiddenRows`에 넣고 `s.rowHeights[r] = 0`을 복구(삭제)해 버립니다.
  - **실증**: `A2=10`(표시), `A3=20`(필터 숨김), `A4=30`(필터 통과 + 수동 숨김)일 때 저장 전 `=SUBTOTAL(9, A2:A4)`는 **`40`**이지만, XLSX 저장 후 다시 열면 **`10`**으로 변하며 필터를 지우면 수동으로 숨겼던 4행까지 강제로 나타납니다.

---

## 2. 병합된 셀 (`merges`) 상호작용

### 🚨 [2-1] 기존 병합 셀 위에 붙여넣기/채우기 시 비-좌상단 셀에 유령 데이터(Phantom Cell) 생성 (High)
- **위치**: [`src/main.mjs:80, 91, 217-222`](../src/main.mjs#L80), [`src/cell-paste.mjs:19-36`](../src/cell-paste.mjs#L19-L36)
- **현상**:
  1. `A1:B1`이 병합된 상태에서 `B1`을 포함하는 범위에 붙여넣기(`Ctrl+V`), 채우기(`Ctrl+D`/`R`), `Ctrl+Enter` 등을 수행하면 병합을 해제하거나 차단하지 않고 **`sheet().cells['B1']`에 값을 직접 기록**합니다.
  2. `renderGrid()`는 병합 범위에서 좌상단 `A1`만 렌더링하므로 `B1`의 값은 화면에 보이지 않으며, `A1:B1`을 클릭하고 `Delete`를 눌러도 `selection`이 `{r:0, c:0, r2:0, c2:0}`이라 `A1`만 지워지고 **`B1`의 유령 데이터는 영구히 남아 `=SUM(A1:B1)`이나 `=B1` 수식 결과를 오염**시킵니다.

---

### 🚨 [2-2] 일반 데이터 정렬(`sortData`) 시 병합 셀 검사 누락 및 수직 병합 셀 필터 렌더링 오류 (High)
- **위치**: [`src/main.mjs:79-80, 236`](../src/main.mjs#L79-L80), [`src/filters.mjs:112`](../src/filters.mjs#L112)
- **현상**:
  1. **일반 정렬(`sortData`) 병합 셀 미차단**: 자동 필터 정렬(`sortFilterRange`)은 범위 내 병합 셀이 있으면 예외를 던져 차단하지만, 일반 정렬(`sortData`, [`src/main.mjs:236`](../src/main.mjs#L236))은 `sheet().merges`를 전혀 검사하지 않습니다. 기본 예제 문서(`demo()`)처럼 병합 행이 있는 시트에서 정렬을 누르면 병합 위치는 고정된 채 셀 데이터만 섞여 시트가 깨집니다.
  2. **수직 병합 셀의 좌상단 행이 필터로 숨겨질 때 렌더링 오류**: `A2:A3`이 병합된 상태에서 필터에 의해 2행(`A2`)이 숨겨지고 3행(`A3`)만 표시될 때, `renderGrid()`([`src/main.mjs:79-80`](../src/main.mjs#L79-L80))가 가시 행 `r=2`를 그리면서 `drawCell(1, 0, m)`을 호출하여 **숨겨진 2행(`A2`)의 값과 `data-r="1"` 속성이 가시 행인 3행 위치에 그대로 렌더링**됩니다.

---

## 3. 붙여넣기 / 행·열 바꿈(Transpose) / 채우기 시 수식 참조 이동

### 🚨 [3-1] 행/열 바꿈 (`transpose: true`) 붙여넣기 시 수식 참조 보정 결함 (High)
- **위치**: [`src/cell-paste.mjs:20, 27`](../src/cell-paste.mjs#L20)
- **코드**:
  ```js
  const sr = transpose ? c % w : r % h, sc = transpose ? r % h : c % w, source = clip.data[sr]?.[sc] || { raw: '' };
  ...
  if (mode !== 'values' && !clip.cut && raw.startsWith('=')) raw = shiftFormula(raw, start.r + r - (clip.sourceRows?.[sr] ?? clip.r + sr), start.c + c - (clip.c + sc));
  ```
- **실증 결과**:
  1. `A1=10`, `B1="=A1"`(왼쪽 셀 참조)인 `1×2` 가로 범위를 복사하여 `A1:A2`(`2×1` 세로)에 `transpose: true`로 붙여넣으면, `B1`(`r=0, c=1`)이 `A2`(`r=1, c=0`)로 가면서 `dr=+1, dc=-1`이 적용되어 `A1`(`0,0`)이 `(1,-1)`로 이동하며 **`A2`가 `=#REF!`가 됩니다!**
  2. `A1=10, B1=20, C1="=SUM(A1:B1)"`을 복사하여 `E1:E3`에 `transpose: true`로 붙여넣으면 `E3`의 수식이 세로 합계 `=SUM(E1:E2)`가 되지 않고 가로 범위 **`=SUM(C3:D3)`**이 됩니다.

---

### 🚨 [3-2] 범위 참조(`=SUM(A1:B2)`)를 시트 경계 밖(위/왼쪽)으로 복사 시 `#REF!` 대신 `#VALUE!` 발생 (Medium)
- **위치**: [`src/engine.mjs:204, 220, 866-886`](../src/engine.mjs#L204)
- **실증 결과**: `shiftFormula('=SUM(A1:B2)', -1, -1)`은 `'=SUM(#REF!:A1)'`을 생성하며, 이를 `evaluateCell`로 평가하면 `Parser.primary()`가 `:A1`을 소비하지 못해 `#REF!`가 아닌 **`#VALUE!`**를 반환합니다.

---

### 🚨 [3-3] 역방향 드래그 선택 후 붙여넣기 시 시작 좌표 오작동 및 다중 셀 채우기 핸들 한계 (Medium)
- **위치**: [`src/main.mjs:218, 309-311`](../src/main.mjs#L218)
- **현상**:
  1. `pasteText`([`src/main.mjs:218`](../src/main.mjs#L218))에서 `selection: bounds()`를 전달하면서 시작 좌표 `start`는 정규화되지 않은 `{ r: selection.r, c: selection.c }`를 전달합니다. 사용자가 `B4`에서 `B1`로 위로 드래그하여 선택한 뒤 붙여넣으면 `B1:B4`가 아닌 **`B4:B7`에 붙여넣어집니다**.
  2. 채우기 핸들(`fillStart`, [`src/main.mjs:311`](../src/main.mjs#L311))은 여러 셀(`A1:A3`)을 선택하고 끌어도 항상 단일 셀(`start.r, start.c`)만 복제합니다.
