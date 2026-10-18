# 02. 수식 평가 엔진, 타입 강제 변환 및 함수 경계 조건 검증 (Formula Engine & Coercion)

- **검증 대상 파일**: [`src/engine.mjs`](../src/engine.mjs), [`src/function-runtime.mjs`](../src/function-runtime.mjs), [`src/function-catalog.mjs`](../src/function-catalog.mjs), [`src/formula-tools.mjs`](../src/formula-tools.mjs)

---

## 요약 테이블

| ID | 심각도 | 구분 | 핵심 내용 | 파일 및 라인 |
|---|---|---|---|---|
| **4-1** | **Critical** | 조건부 함수 (`COUNTIF`/`SUMIF`/`*IFS`) | `0`/`FALSE` 조건이 빈 셀과 매칭됨; `"="`·`"<>"`가 빈 셀과 `""`를 구분하지 못함; 관계 연산자(`">100"`, `"<200"`, `"<z"`, `">a"`)가 숫자 문자열·빈 셀·타 타입을 잘못 매칭함 | [`src/engine.mjs:376-401`](../src/engine.mjs#L376-L401) |
| **4-2** | **High** | 조회 함수 (`VLOOKUP`/`MATCH`/`XLOOKUP` 등) | 조회 범위 내 빈 셀(`BLANK`)이 `0`, `FALSE`, `""`(`XLOOKUP`/`XMATCH`) 및 근삿값 조회(`VLOOKUP(..., TRUE)`)와 매칭되어 오답 반환 | [`src/engine.mjs:598-639`](../src/engine.mjs#L598-L639), [`src/function-runtime.mjs:78-90`](../src/function-runtime.mjs#L78-L90) |
| **1-1** | **High** | 타입 강제 변환 (`numeric`) | 빈 문자열(`""`)이 `0`으로 변환되어 `=""+1` → `1`, `=""*5` → `0`, `=SUM("", 5)` → `5` 반환 (Excel은 `#VALUE!`) | [`src/engine.mjs:33-40`](../src/engine.mjs#L33-L40) |
| **1-2** | **High** | 타입 강제 변환 (`logical`) | 빈 문자열(`""`)이 `false`로 변환되어 `=IF("", "T", "F")` → `"F"`, `=NOT("")` → `true` 반환 (Excel은 `#VALUE!`) | [`src/engine.mjs:48-56`](../src/engine.mjs#L48-L56) |
| **3-1** | **High** | 날짜 포맷 (`TEXT` / `formatDateValue`) | 1900-01-01 ~ 1900-02-29 구간에서 요일(`dddd`)이 하루씩 밀리고, `=TEXT(60, "mmmm dd, yyyy")`가 `"March 29, 1900"`을 반환 | [`src/engine.mjs:705-741`](../src/engine.mjs#L705-L741) |
| **3-2** | **Medium** | 날짜 계산 (`DATE`) | 1900년 2월 전후 일수 오버플로우/언더플로우 시 `DATE(1900,2,30)` → `62`(Excel `61`), `DATE(1900,3,0)` → `59`(Excel `60`) | [`src/engine.mjs:403-411, 643-650`](../src/engine.mjs#L403-L411) |
| **1-3** | **Medium** | 단항 연산자 (`+`) | 단항 `+`가 문자열·논리값에 대해 `numeric()`을 강제 호출하여 `=+"abc"` → `#VALUE!`, `=+TRUE` → `1` 반환 | [`src/engine.mjs:320-324`](../src/engine.mjs#L320-L324) |
| **1-4** | **Medium** | 생략된 함수 인자 (`BLANK`) | `=COUNTA(1, , 2)` → `2`(Excel `3`), `=AVERAGE(10, , 20)` → `15`(Excel `10`), `=IF(TRUE, , 5)&"x"` → `"x"`(Excel `"0x"`) | [`src/engine.mjs:210, 359, 532`](../src/engine.mjs#L210) |
| **2-1** | **Medium** | `COUNT` / `COUNTA` 오류 인자 | 직접 인자에 오류가 있을 때(`=COUNT(1/0, 5)`, `=COUNTA(1/0, 5)`) 무시/카운트하지 못하고 `#DIV/0!` 전파 | [`src/engine.mjs:526`](../src/engine.mjs#L526) |
| **4-3** | **Medium** | `XLOOKUP` / `SUMIF` 인자 처리 | `XLOOKUP`의 4번째 인자 생략(`=XLOOKUP(99,{1,2},{10,20},,0)`) 시 `0` 반환(Excel `#N/A`); `SUMIF`의 `sum_range` 크기 불일치 시 `#VALUE!` | [`src/engine.mjs:589, 636`](../src/engine.mjs#L589) |
| **5-2** | **Medium** | 참조 연산자 (`:`, 공백, `,`) | `INDEX(...) : B3` 범위 구성, 공백(교집합), 괄호 `,`(합집합) 미지원 (`#VALUE!` / `#REF!`) | [`src/engine.mjs:99, 206-241`](../src/engine.mjs#L99) |
| **5-3** | **Medium** | 배열/범위 인자 (`IFERROR`) | `=IFERROR({1, 2}, 1/0)`에서 `{1, 2}`를 `scalar()`가 `#VALUE!`로 간주하여 `value_if_error` 분기를 실행함 | [`src/engine.mjs:25, 516`](../src/engine.mjs#L25) |
| **2-2** | **Low** | `AVERAGEA`/`MINA`/`MAXA` | 배열 상수 `{1, "abc"}` 전달 시 문자열을 무시하지 않고 `0`으로 처리 | [`src/function-runtime.mjs:37-38`](../src/function-runtime.mjs#L37-L38) |
| **5-4** | **Low** | 수식 도구 (`formulaReferences`) | `Sheet1!A1:sheet1!B2` 범위에서 시트 이름 대소문자 불일치 시 범위 인식 실패 | [`src/formula-tools.mjs:24`](../src/formula-tools.mjs#L24) |

---

## 1. 빈 셀(`BLANK`) vs 빈 문자열(`""`) vs `0` 처리

### 🚨 [1-1] `numeric()`이 빈 문자열(`""`)을 `0`으로 강제 변환 (High)
- **위치**: [`src/engine.mjs:33-40`](../src/engine.mjs#L33-L40) (`if (value === BLANK || value === '') return 0;`)
- **원인**: Excel에서 진짜 빈 셀(`BLANK`)은 산술 연산 시 `0`으로 취급되지만, 빈 문자열 리터럴(`""`)이나 `=""`를 반환하는 수식 셀은 **텍스트(Text)**이므로 산술 연산 및 숫자 강제 변환 시 `#VALUE!` 오류가 발생해야 합니다. 그러나 `numeric()`이 `value === ''`까지 `0`으로 변환합니다.
- **실증 비교 결과**:
  | 수식 / 조건 | Sheet Studio 출력 | Microsoft Excel 출력 | 일치 여부 |
  |---|---|---|---|
  | `=A1+1`, `=A1*5` (`A1`이 빈 셀) | `1`, `0` | `1`, `0` | 일치 |
  | `=TRUE+1`, `=FALSE+1`, `=TRUE*5` | `2`, `1`, `5` | `2`, `1`, `5` | 일치 |
  | `=""+1` | **`1`** | **`#VALUE!`** | **불일치** |
  | `=""*5` | **`0`** | **`#VALUE!`** | **불일치** |
  | `=-""`, `=""%` | **`0`**, **`0`** | **`#VALUE!`**, **`#VALUE!`** | **불일치** |
  | `=A1+1` (`A1`에 `=""` 수식) | **`1`** | **`#VALUE!`** | **불일치** |
  | `=SUM("", 5)` (직접 스칼라 인자 `""`) | **`5`** | **`#VALUE!`** | **불일치** |
  | `=AVERAGE("", 10)` (직접 스칼라 인자 `""`) | **`5`** (`(0+10)/2`) | **`#VALUE!`** | **불일치** |

---

### 🚨 [1-2] `logical()`이 빈 문자열(`""`)을 `FALSE`로 강제 변환 (High)
- **위치**: [`src/engine.mjs:48-56`](../src/engine.mjs#L48-L56) (`if (value === BLANK || value === '') return false;`)
- **원인**: Excel에서 빈 셀(`BLANK`)은 논리 문맥(`=IF(A1, ...)` 등)에서 `FALSE`로 평가되지만, 빈 문자열(`""`)은 `"TRUE"`/`"FALSE"`가 아닌 일반 문자열이므로 `#VALUE!`를 발생시켜야 합니다.
- **실증 비교 결과**:
  | 수식 / 조건 | Sheet Studio 출력 | Microsoft Excel 출력 | 일치 여부 |
  |---|---|---|---|
  | `=IF(A1, "T", "F")` (`A1`이 빈 셀) | `"F"` | `"F"` | 일치 |
  | `=IF("", "T", "F")` | **`"F"`** | **`#VALUE!`** | **불일치** |
  | `=IF(A1, "T", "F")` (`A1`에 `=""` 수식) | **`"F"`** | **`#VALUE!`** | **불일치** |
  | `=NOT("")` | **`true`** | **`#VALUE!`** | **불일치** |

---

### 🚨 [1-3] 단항 플러스(`+`) 연산자가 문자열·논리값에 대해 항등 연산이 아닌 `numeric()` 변환 수행 (Medium)
- **위치**: [`src/engine.mjs:320-324`](../src/engine.mjs#L320-L324)
- **현상**: Excel에서 단항 `+` 연산자는 피연산자의 타입을 바꾸지 않는 항등(identity) 연산자입니다(`=+"abc"` → `"abc"`, `=+""` → `""`, `=+TRUE` → `TRUE`). 그러나 `engine.mjs`는 `node.op === '+'`에도 `numeric(v)`를 호출하여 `=+"abc"` → `#VALUE!`, `=+TRUE` → `1`, `=+""` → `0`을 반환합니다.

---

### 🚨 [1-4] 생략된 함수 인자(Omitted argument, 예: `=AVERAGE(10, , 20)`)가 `BLANK`로 파싱되어 무시됨 (Medium)
- **위치**: [`src/engine.mjs:210, 359, 532`](../src/engine.mjs#L210)
- **실증 비교 결과**:
  - `=COUNTA(1, , 2)`: Sheet Studio **`2`** vs Excel **`3`** (생략된 인자 슬롯도 카운트됨)
  - `=AVERAGE(10, , 20)`: Sheet Studio **`15`** vs Excel **`10`** (생략된 직접 인자는 `0`으로 취급되어 `(10+0+20)/3`)
  - `=MIN(10, , 20)`: Sheet Studio **`10`** vs Excel **`0`**
  - `=IF(TRUE, , 5) & "x"`: Sheet Studio **`"x"`** vs Excel **`"0x"`** (생략된 분기가 `0` 대신 `BLANK`를 반환)
- 참고: 일반 비교 연산자(`=A1=0` → `true`, `=A1=""` → `true`, `=0=""` → `false`, `="9">10` → `true`, `=TRUE>"Z"` → `true`, `="a"="A"` → `true`)와 범위 내 `ISBLANK`/`COUNTA`/`COUNTBLANK`/`AVERAGE`는 Excel과 정확히 일치합니다.

---

## 2. 직접 스칼라 인자 vs 범위(Range) 참조 인자

- **정상 동작 확인**: `=SUM("5", TRUE)` → `6` vs `=SUM(A1:A2)` → `0` (`A1="5", A2=TRUE`) 등 직접 인자와 범위 인자의 구분은 `item.reference` 플래그를 통해 Excel과 동일하게 동작합니다.

### 🚨 [2-1] `COUNT` 및 `COUNTA`에 오류값이 직접 인자로 전달될 때 즉시 오류가 전파됨 (Medium)
- **위치**: [`src/engine.mjs:526`](../src/engine.mjs#L526)
- **원인**: 범위 참조 내의 오류(`=COUNT(A1:A2)`, `A1`이 `=1/0`)는 `COUNT`가 무시(`1`)하고 `COUNTA`가 포함(`2`)하여 Excel과 일치하지만, 직접 인자(`=COUNT(1/0, 5)`, `=COUNTA(1/0, 5)`)는 `nodes.map(evaluate)` 단계에서 예외가 던져져 함수 전체가 `#DIV/0!`로 실패합니다 (Excel은 각각 `1`, `2` 반환).

### 🚨 [2-2] `AVERAGEA`, `MINA`, `MAXA`에 배열 상수(`{1, "abc"}`) 전달 시 문자열을 `0`으로 처리 (Low)
- **위치**: [`src/function-runtime.mjs:33-40`](../src/function-runtime.mjs#L33-L40)
- **현상**: `=MINA({1, "abc"})`가 Excel에서는 `1`이지만 Sheet Studio에서는 `0`을 반환합니다.

---

## 3. 날짜 일련번호(Date Serial Numbers) 및 1900년 윤년 버그

### 🚨 [3-1] `TEXT` 및 셀 날짜 서식(`formatDateValue`)에서 `1900-01-01` ~ `1900-02-29` 구간 요일 및 월 이름 오계산 (High)
- **위치**: [`src/engine.mjs:705-741`](../src/engine.mjs#L705-L741) (`const date = new Date(Date.UTC(year, month - 1, Math.max(1, day)));`)
- **원인**:
  1. `serial = 60` (`1900-02-29`)일 때 `serialParts(60)`은 `{ year: 1900, month: 2, day: 29 }`를 반환하지만, `Date.UTC(1900, 1, 29)`는 그레고리력 평년 기준 **1900년 3월 1일(목요일)**로 오버플로우됩니다.
  2. `serial = 1` ~ `59` 구간에서도 `WEEKDAY` 함수는 Excel과 일치(`WEEKDAY(1)=1` 일요일)하지만, `formatDateValue`는 JS `Date.UTC`의 요일을 그대로 사용하여 요일이 하루씩 밀립니다.
- **실증 비교 결과**:
  | 수식 | Sheet Studio 출력 | Microsoft Excel 출력 | 일치 여부 |
  |---|---|---|---|
  | `=TEXT(1, "dddd")` (`1900-01-01`, `WEEKDAY=1`) | **`"Monday"`** | **`"Sunday"`** | **불일치** |
  | `=TEXT(59, "dddd")` (`1900-02-28`, `WEEKDAY=3`) | **`"Wednesday"`** | **`"Tuesday"`** | **불일치** |
  | `=TEXT(60, "dddd")` (`1900-02-29`, `WEEKDAY=4`) | **`"Thursday"`** | **`"Wednesday"`** | **불일치** |
  | `=TEXT(60, "mmmm dd, yyyy")` (`1900-02-29`) | **`"March 29, 1900"`** | **`"February 29, 1900"`** | **불일치** |

---

### 🚨 [3-2] `DATE` 함수에서 1900년 2월 말일 전후의 일수 오버플로우/언더플로우 오계산 (Medium)
- **위치**: [`src/engine.mjs:403-411, 643-650`](../src/engine.mjs#L403-L411)
- **원인**: `year === 1900 && month === 2 && day === 29`만 하드코딩(`return 60`)하고 나머지는 JS `Date.setUTCFullYear`에 위임하여 1900년 2월이 28일로 계산됩니다.
  - `=DATE(1900, 2, 30)`: Sheet Studio **`62`** vs Excel **`61`** (`1900-03-01`)
  - `=DATE(1900, 3, 0)`: Sheet Studio **`59`** vs Excel **`60`** (`1900-02-29`)
  - `=DATE(1900, 1, 60)`: Sheet Studio **`61`** vs Excel **`60`** (`1900-02-29`)

---

## 4. 조건부 집계 함수(`COUNTIF`, `SUMIF`, `*IFS`) 및 조회 함수(`MATCH`, `VLOOKUP`, `XLOOKUP`)

### 🚨 [4-1] `COUNTIF`/`SUMIF`/`*IFS`의 `criterion()`이 빈 셀, 빈 조건, 숫자 텍스트 및 관계 연산자를 잘못 평가 (Critical)
- **위치**: [`src/engine.mjs:376-401`](../src/engine.mjs#L376-L401)
- **원인**:
  1. `criterion()`이 내부적으로 `compare(value, target)`을 호출하는데, `compare`는 `BLANK`를 상대 타입에 맞춰 `0`, `""`, `false`로 자동 변환합니다. 그 결과 `0`, `FALSE`, `"<200"`, `">=0"` 조건이 **범위 내의 모든 빈 셀(`BLANK`)을 매칭**합니다.
  2. `"="`는 `op='=', target=''`로 파싱되어 `""`와 동일하게 `=""` 수식 셀까지 매칭하며(Excel에서 `"="`는 진짜 빈 셀만 매칭), `"<>"`는 `=""` 수식 셀까지 제외해 버립니다.
  3. 394~396행에서 `typeof target === 'number' && typeof value === 'string'`일 때 `op`가 `=`/`<>`인지 확인하지 않고 숫자로 변환하여, `">100"` 조건이 **문자열 `"123"`까지 매칭**합니다.
  4. 문자열 관계 연산(`"<z"`, `">a"`)에서도 `compare`의 타입 서열(`number < string < boolean`) 때문에 `"<z"`가 숫자와 빈 셀을 모두 매칭하고 `">a"`가 `TRUE`/`FALSE`를 매칭합니다.
- **실증 비교 결과** (`A1=123`, `A2="123"`, `A3="abc"`, `A4=TRUE`, `A5=""`, `A6` 빈 셀):
  | 수식 | Sheet Studio 출력 | Microsoft Excel 출력 | 일치 여부 |
  |---|---|---|---|
  | `=COUNTIF(A1:A6, "=")` | **`2`** (`A5`, `A6`) | **`1`** (`A6` 빈 셀만) | **불일치** |
  | `=COUNTIF(A1:A6, "<>")` | **`4`** (`A5` 제외됨) | **`5`** (`A6`만 제외) | **불일치** |
  | `=COUNTIF(A1:A6, 0)` | **`1`** (`A6` 빈 셀 매칭!) | **`0`** | **불일치** |
  | `=COUNTIF(A1:A6, FALSE)` | **`1`** (`A6` 빈 셀 매칭!) | **`0`** | **불일치** |
  | `=COUNTIF(A1:A2, ">100")` | **`2`** (`"123"` 매칭!) | **`1`** (`123`만 매칭) | **불일치** |
  | `=COUNTIF(A1:A6, "<200")` | **`3`** (`123`, `"123"`, 빈 셀!) | **`1`** (`123`만 매칭) | **불일치** |
  | `=COUNTIF(A1:A6, "<z")` | **`5`** (숫자·빈 셀 포함!) | **`2`** (문자열만 비교) | **불일치** |
  | `=COUNTIF(A1:A6, ">a")` | **`2`** (`"abc"`, `TRUE`!) | **`1`** (`"abc"`만) | **불일치** |

---

### 🚨 [4-2] `VLOOKUP`, `HLOOKUP`, `LOOKUP`, `MATCH`, `XLOOKUP`, `XMATCH`에서 빈 셀(`BLANK`) 오매칭 (High)
- **위치**: [`src/engine.mjs:598-639`](../src/engine.mjs#L598-L639), [`src/function-runtime.mjs:78-90`](../src/function-runtime.mjs#L78-L90)
- **원인**: 조회 범위의 셀이 `BLANK`일 때 건너뛰지 않고 `compare(value, lookup)`을 그대로 호출합니다.
- **실증 비교 결과**:
  | 수식 / 조건 | Sheet Studio 출력 | Microsoft Excel 출력 | 일치 여부 |
  |---|---|---|---|
  | `=MATCH(0, A6:A7, 0)` (`A6=""`, `A7` 빈 셀) | **`2`** (`A7` 매칭!) | **`#N/A`** | **불일치** |
  | `=VLOOKUP(0, A6:B7, 2, FALSE)` (`A6=""`, `A7` 빈 셀, `B7=70`) | **`70`** | **`#N/A`** | **불일치** |
  | `=XLOOKUP(0, A6:A7, B6:B7)` (`A6=""`, `A7` 빈 셀, `B7=70`) | **`70`** | **`#N/A`** | **불일치** |
  | `=XLOOKUP("", A6:A7, B6:B7)` (`A6` 빈 셀, `A7=""`, `B6=32, B7=64`) | **`32`** (`A6` 빈 셀 매칭!) | **`64`** (`A7=""` 매칭) | **불일치** |
  | `=XMATCH("", A6:A7)` (`A6` 빈 셀, `A7=""`) | **`1`** (`A6` 빈 셀 매칭!) | **`2`** (`A7=""` 매칭) | **불일치** |
  | `=VLOOKUP(5, A1:B3, 2, TRUE)` (`A1=1, A2=3, A3=빈셀, B1=10, B2=30, B3=99`) | **`99`** (`A3` 빈 셀 선택!) | **`30`** (`A2` 선택) | **불일치** |
  | `=VLOOKUP("abc", {100,"num";"def","str"}, 2, TRUE)` | **`"num"`** (숫자 `100 < "abc"`) | **`#N/A`** | **불일치** |

---

### 🚨 [4-3] `XLOOKUP`의 4번째 인자 생략 시 `0` 반환 및 `SUMIF`의 `sum_range` 크기 자동 확장 미지원 (Medium)
- **위치**: [`src/engine.mjs:589, 636`](../src/engine.mjs#L589)
- **현상**:
  - `=XLOOKUP(99, {1,2}, {10,20}, , 0)`: `nodes[3]`이 생략되어 `BLANK`일 때 `args.length > 3` 조건만 검사하여 `#N/A` 대신 **`0`**을 반환합니다.
  - `=SUMIF(A1:A2, ">0", B1)`: Excel에서는 3번째 인자 `B1`을 `range` 크기(`B1:B2`)로 자동 확장하지만, Sheet Studio는 크기가 다르면 `#VALUE!`를 던집니다.

---

## 5. 연산자 우선순위, 참조 연산자 및 단락 평가

- **연산자 우선순위 및 단락 평가 정상 확인**:
  - `=-2^2` → `4`, `=2^3^2` → `64`, `=2^-2^2` → `0.0625`, `=50%+10%` → `0.6` 모두 **Excel과 일치**합니다.
  - `IF`, `IFS`, `IFERROR`, `IFNA`, `CHOOSE`, `SWITCH`에서 선택되지 않은 분기의 `#DIV/0!`나 순환 참조는 지연 평가(lazy evaluation)로 단락되어 **Excel과 일치**합니다.

### 🚨 [5-2] 함수 반환 참조에 대한 범위 연산자(`:`), 교집합(공백), 합집합(`,`) 미지원 (Medium)
- **위치**: [`src/engine.mjs:99, 206-241`](../src/engine.mjs#L99)
- **현상**:
  - `=SUM(INDEX(A1:B2,1,2):B3)` → `#VALUE!` (Excel은 `B1:B3` 합계)
  - `=SUM(A1:INDEX(A1:B3,3,2))` → `#REF!` (Excel은 `A1:B3` 합계)
  - `=SUM(A1:B2 A2:C3)` (공백 교집합) → `#VALUE!`, `=SUM((A1:A2, B1:B2))` (괄호 합집합) → `#VALUE!`

### 🚨 [5-3] `IFERROR`에 배열 상수 전달 시 오류 분기 오실행 (Medium)
- **위치**: [`src/engine.mjs:25, 516`](../src/engine.mjs#L25)
- **현상**: `=IFERROR({1, 2}, "err")`에서 `checked({1, 2})`가 `scalar()`를 호출해 `#VALUE!`를 던지고, `IFERROR`가 이를 잡아서 `"err"` 분기를 실행합니다.

### 🚨 [5-4] `formulaReferences`의 시트 이름 대소문자 비교 불일치 (Low)
- **위치**: [`src/formula-tools.mjs:24`](../src/formula-tools.mjs#L24)
- **현상**: `engine.mjs`는 `Sheet1!A1:sheet1!B2`를 대소문자 구분 없이 같은 시트 범위로 인정하지만, `formulaReferences`는 `b.sheet !== a.sheet`로 엄격 비교하여 하이라이트에서 누락합니다.
