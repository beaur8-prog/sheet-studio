# 함수 라이브러리 — 0.5.0

37개에서 **379개 함수 이름**으로 확장했습니다. 이 중 30개는 이전 Excel 이름의 별칭입니다. 함수 이름을 지원한다는 의미이며 모든 인수 조합과 Excel 내부 계산 규칙이 완전히 동일하다는 의미는 아닙니다.

## 사용법

- **수식 → 함수 삽입** 또는 수식 입력줄의 **ƒx**를 눌러 이름·한국어 설명으로 검색하고 범주를 선택합니다. 구문과 예시를 확인한 뒤 **수식에 삽입**하면 편집 상태에서 인수를 입력할 수 있습니다.
- 셀 또는 수식 입력줄에 `=SUMI`처럼 입력하면 후보가 나옵니다. **↑↓**로 선택하고 **Tab**으로 삽입합니다. **Esc**는 후보를 닫고, 다시 누르면 편집을 취소합니다. **Enter**는 수식을 확정합니다.
- 인수 안내는 중첩 함수, 따옴표 안의 쉼표, 배열 상수의 구분자를 구분하여 현재 인수를 강조합니다. 한국어 IME 조합 중에는 완성 목록이 입력을 가로채지 않습니다.
- `=SUMPRODUCT((A1:A5>2)*A1:A5)`, `=SUM({1,2;3,4})`, `=SUM(INDEX(A1:C5,0,2))`처럼 배열을 내부 계산에 사용할 수 있습니다. 셀 여러 개로 결과를 펼치는 스필과 모든 함수의 배열 자동 확장(예: 배열 조건 IF)은 아직 구현하지 않았습니다.

## 검증 범위

Windows Excel 16.0, build 19127에서 **371개 함수 예시와 111개 추가 사례**의 결과를 기록해 플랫폼 공통 회귀 검사에 사용합니다. 수치 비교 허용 오차는 `1e-7 × max(1, |Excel 값|)`이며 복소수는 실수부·허수부를 비교합니다. 확률 역함수와 반복 재무 계산은 Excel과 마지막 유효 숫자가 다를 수 있습니다. 실제 Excel에서 앱이 만든 XLSX를 읽고 375개 결과를 다시 계산해 직접 입력 결과와도 대조합니다.

이 로컬 Excel은 XLOOKUP·XMATCH·TEXTBEFORE·TEXTAFTER를 인식하지 못하므로 이 네 함수는 Microsoft 문서에 따른 별도 기대값과 경계 검사로 확인합니다. NOW·TODAY·RAND·RANDBETWEEN은 고정 Excel 결과 대신 반환 형식과 범위를 검사합니다. 전수 경계값 검증이나 macOS 실제 실행 검증은 아닙니다.

## 호환 범위와 남은 제한

- 기본 연산, 숫자/텍스트/논리값/빈 셀 구분, 오류 전파, 순환 참조 감지, 상대·절대 A1 참조와 다른 시트 참조를 지원합니다. IF·IFERROR·IFNA·IFS·SWITCH·CHOOSE의 선택되지 않은 결과 식은 평가하지 않습니다.
- ROW/COLUMN은 단일 번호, INDEX는 셀이나 내부 범위를 반환합니다. 다중 영역 INDEX, 3차원 참조, 이름 관리자, 외부 통합 문서 참조, 전체 행/열 참조, 표 구조화 참조, @ 암시적 교차, # 스필 참조는 미지원입니다. ARRAY의 최종 결과는 한 셀이어야 합니다.
- FILTER·SORT·UNIQUE·SEQUENCE 같은 스필 함수, LET·LAMBDA·MAP·REDUCE, 데이터베이스 함수, 큐브·웹·추가 기능 함수, 채권 가격·수익률 일부와 다변량 회귀/행렬 함수는 아직 없습니다. 없는 함수는 #NAME?을 반환합니다.
- SUBTOTAL은 직접 중첩된 SUBTOTAL/AGGREGATE를 제외합니다. 0.6.0부터 자동 필터로 숨긴 행을 모든 집계 코드에서 제외하며, 101~111 코드는 수동 숨김(행 높이 0)도 제외합니다. 일반 SUM은 숨긴 행을 포함합니다.
- TEXT는 앱의 숫자·날짜 표시 형식을 사용합니다. Excel의 모든 다중 구간/조건부/분수/회계 표시 코드를 구현하지 않았습니다. DOLLAR는 원화를 사용하며 CHAR/CODE의 128 이상 문자는 OS별 Excel 코드 페이지와 다를 수 있습니다. 날짜 문자열은 YYYY-MM-DD/구분자 변형 및 미국식 M/D/YYYY, 시간 문자열은 HH:MM[:SS]와 AM/PM을 지원합니다.
- XLOOKUP/XMATCH의 이진 검색 모드는 정렬된 입력에서의 값 검색을 지원하지만 선형 탐색을 사용합니다. 동일 값이 반복될 때 Excel 이진 탐색이 고르는 위치와 다를 수 있습니다.
- 함수당 범위·배열은 최대 100,000셀입니다. 긴 문자열·반복 계산·계승·이항/조합 및 근무일 계산에도 제한을 둡니다. XIRR은 최대 10,000개 현금 흐름과 제한된 반복 횟수를 사용하며 찾지 못한 근은 #NUM!을 반환합니다.
- XLSX 저장 시 MS-XLSX의 최신 함수에 `_xlfn.` 접두사를 붙이고 읽기/계산 시 이를 인식합니다. 알려지지 않은 접두사 함수는 보존합니다. 해당 함수를 제공하지 않는 구버전 Excel까지 실행 가능하게 만드는 기능은 아닙니다. XLSM/VBA 사용자 정의 함수는 기존 Windows Excel 연결 경로를 사용합니다.

계산 확장에는 [Formula.js](https://github.com/formulajs/formulajs) 4.6.1을 사용합니다. 날짜 번호, 조건 집계, 참조, 오류 검사, 배열/참조 인수 구분, 일부 분포와 재무 계산 등은 별도로 보정했습니다. [MIT 고지](../public/THIRD-PARTY-NOTICES.txt)는 배포 파일에도 포함합니다.

참고: [Microsoft 함수 분류](https://support.microsoft.com/en-us/excel/excel-functions-by-category), [XLSX 미래 함수 규격](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf), [XLOOKUP](https://support.microsoft.com/en-gb/excel/functions/xlookup-function), [TEXTBEFORE](https://support.microsoft.com/en-us/excel/functions/textbefore-function).

## 함수 목록

### 수학·삼각 (72)

| 함수 | 구문 | 용도 |
|---|---|---|
| ABS | `ABS(number)` | 절댓값 |
| ACOS | `ACOS(number)` | 역코사인 |
| ACOSH | `ACOSH(number)` | 역쌍곡 코사인 |
| ACOT | `ACOT(number)` | 역코탄젠트 |
| ACOTH | `ACOTH(number)` | 역쌍곡 코탄젠트 |
| ASIN | `ASIN(number)` | 역사인 |
| ASINH | `ASINH(number)` | 역쌍곡 사인 |
| ATAN | `ATAN(number)` | 역탄젠트 |
| ATAN2 | `ATAN2(x_num, y_num)` | 좌표의 역탄젠트 |
| ATANH | `ATANH(number)` | 역쌍곡 탄젠트 |
| BASE | `BASE(number, radix, [min_length])` | 숫자를 다른 진법의 문자열로 변환 |
| CEILING | `CEILING(number, significance)` | 배수로 올림 |
| CEILING.MATH | `CEILING.MATH(number, [significance], [mode])` | 방향을 지정하여 배수로 올림 |
| CEILING.PRECISE | `CEILING.PRECISE(number, [significance])` | 배수로 올림 |
| COMBIN | `COMBIN(number, number_chosen)` | 조합 수 |
| COMBINA | `COMBINA(number, number_chosen)` | 중복 조합 수 |
| COS | `COS(number)` | 코사인 |
| COSH | `COSH(number)` | 쌍곡 코사인 |
| COT | `COT(number)` | 코탄젠트 |
| COTH | `COTH(number)` | 쌍곡 코탄젠트 |
| CSC | `CSC(number)` | 코시컨트 |
| CSCH | `CSCH(number)` | 쌍곡 코시컨트 |
| DECIMAL | `DECIMAL(text, radix)` | 진법 문자열을 십진수로 변환 |
| DEGREES | `DEGREES(angle)` | 라디안을 도로 변환 |
| EVEN | `EVEN(number)` | 가장 가까운 짝수로 절댓값 올림 |
| EXP | `EXP(number)` | 자연상수의 거듭제곱 |
| FACT | `FACT(number)` | 계승 |
| FACTDOUBLE | `FACTDOUBLE(number)` | 이중 계승 |
| FLOOR | `FLOOR(number, significance)` | 배수로 내림 |
| FLOOR.MATH | `FLOOR.MATH(number, [significance], [mode])` | 방향을 지정하여 배수로 내림 |
| FLOOR.PRECISE | `FLOOR.PRECISE(number, [significance])` | 배수로 내림 |
| GCD | `GCD(number1, [number2], …)` | 최대 공약수 |
| INT | `INT(number)` | 정수로 내림 |
| ISO.CEILING | `ISO.CEILING(number, [significance])` | 배수로 올림 |
| LCM | `LCM(number1, [number2], …)` | 최소 공배수 |
| LN | `LN(number)` | 자연로그 |
| LOG | `LOG(number, [base])` | 지정한 밑의 로그 |
| LOG10 | `LOG10(number)` | 상용로그 |
| MOD | `MOD(number, divisor)` | 나눗셈의 나머지 |
| MROUND | `MROUND(number, multiple)` | 배수로 반올림 |
| MULTINOMIAL | `MULTINOMIAL(number1, [number2], …)` | 다항 계수 |
| ODD | `ODD(number)` | 가장 가까운 홀수로 절댓값 올림 |
| PI | `PI()` | 원주율 |
| POWER | `POWER(number, power)` | 거듭제곱 |
| PRODUCT | `PRODUCT(number1, [number2], …)` | 곱 |
| QUOTIENT | `QUOTIENT(numerator, denominator)` | 나눗셈의 정수 부분 |
| RADIANS | `RADIANS(angle)` | 도를 라디안으로 변환 |
| RAND | `RAND()` | 0 이상 1 미만의 난수 |
| RANDBETWEEN | `RANDBETWEEN(bottom, top)` | 지정한 구간의 정수 난수 |
| ROUND | `ROUND(number, num_digits)` | 반올림 |
| ROUNDDOWN | `ROUNDDOWN(number, num_digits)` | 절댓값 내림 |
| ROUNDUP | `ROUNDUP(number, num_digits)` | 절댓값 올림 |
| SEC | `SEC(number)` | 시컨트 |
| SECH | `SECH(number)` | 쌍곡 시컨트 |
| SERIESSUM | `SERIESSUM(x, n, m, coefficients)` | 거듭제곱 급수의 합 |
| SIGN | `SIGN(number)` | 숫자의 부호 |
| SIN | `SIN(number)` | 사인 |
| SINH | `SINH(number)` | 쌍곡 사인 |
| SQRT | `SQRT(number)` | 제곱근 |
| SQRTPI | `SQRTPI(number)` | 원주율을 곱한 수의 제곱근 |
| SUBTOTAL | `SUBTOTAL(function_num, ref1, [ref2], …)` | 중첩 부분합을 제외한 집계 |
| SUM | `SUM(number1, [number2], …)` | 합계 |
| SUMIF | `SUMIF(range, criteria, [sum_range])` | 한 조건을 만족하는 값의 합계 |
| SUMIFS | `SUMIFS(sum_range, criteria_range1, criteria1, …)` | 여러 조건을 모두 만족하는 값의 합계 |
| SUMPRODUCT | `SUMPRODUCT(array1, [array2], …)` | 같은 위치의 값들을 곱한 뒤 합계 |
| SUMSQ | `SUMSQ(number1, [number2], …)` | 제곱의 합 |
| SUMX2MY2 | `SUMX2MY2(array_x, array_y)` | 제곱 차의 합 |
| SUMX2PY2 | `SUMX2PY2(array_x, array_y)` | 제곱 합의 합 |
| SUMXMY2 | `SUMXMY2(array_x, array_y)` | 차를 제곱한 값의 합 |
| TAN | `TAN(number)` | 탄젠트 |
| TANH | `TANH(number)` | 쌍곡 탄젠트 |
| TRUNC | `TRUNC(number, [num_digits])` | 지정한 자릿수에서 버림 |

### 찾기·참조 (16)

| 함수 | 구문 | 용도 |
|---|---|---|
| ADDRESS | `ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])` | 셀 주소 문자열 생성 |
| CHOOSE | `CHOOSE(index_num, value1, [value2], …)` | 순번으로 값 선택 |
| COLUMN | `COLUMN([reference])` | 열 번호 |
| COLUMNS | `COLUMNS(array)` | 열 개수 |
| FORMULATEXT | `FORMULATEXT(reference)` | 셀의 수식을 텍스트로 반환 |
| HLOOKUP | `HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])` | 첫 행에서 가로로 찾아 반환 |
| INDEX | `INDEX(array, row_num, [column_num])` | 범위의 지정한 위치 값 또는 행·열 참조 |
| INDIRECT | `INDIRECT(ref_text, [a1])` | 문자열을 셀 참조로 변환 |
| LOOKUP | `LOOKUP(lookup_value, lookup_vector, [result_vector])` | 정렬된 범위에서 근삿값 찾기 |
| MATCH | `MATCH(lookup_value, lookup_array, [match_type])` | 일치하는 값의 위치 |
| OFFSET | `OFFSET(reference, rows, cols, [height], [width])` | 셀에서 행·열만큼 이동한 참조 |
| ROW | `ROW([reference])` | 행 번호 |
| ROWS | `ROWS(array)` | 행 개수 |
| VLOOKUP | `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])` | 첫 열에서 세로로 찾아 반환 |
| XLOOKUP | `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | 방향·일치 방식을 지정하여 찾기 |
| XMATCH | `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])` | 방향·일치 방식을 지정하여 위치 찾기 |

### 논리 (11)

| 함수 | 구문 | 용도 |
|---|---|---|
| AND | `AND(logical1, [logical2], …)` | 모든 조건이 참인지 검사 |
| FALSE | `FALSE()` | 논리값 FALSE |
| IF | `IF(logical_test, value_if_true, [value_if_false])` | 조건에 따라 값 선택 |
| IFERROR | `IFERROR(value, value_if_error)` | 오류일 때 대체 값 |
| IFNA | `IFNA(value, value_if_na)` | N/A 오류일 때 대체 값 |
| IFS | `IFS(logical_test1, value_if_true1, …)` | 첫 번째 참인 조건의 값 |
| NOT | `NOT(logical)` | 논리값 반전 |
| OR | `OR(logical1, [logical2], …)` | 하나 이상의 조건이 참인지 검사 |
| SWITCH | `SWITCH(expression, value1, result1, …)` | 값에 따라 결과 선택 (마지막 기본값 선택 가능) |
| TRUE | `TRUE()` | 논리값 TRUE |
| XOR | `XOR(logical1, [logical2], …)` | 참인 조건의 수가 홀수인지 검사 |

### 통계 (57)

| 함수 | 구문 | 용도 |
|---|---|---|
| AVEDEV | `AVEDEV(number1, [number2], …)` | 평균 절대 편차 |
| AVERAGE | `AVERAGE(number1, [number2], …)` | 산술 평균 |
| AVERAGEA | `AVERAGEA(value1, [value2], …)` | 논리값과 텍스트를 포함하는 평균 |
| AVERAGEIF | `AVERAGEIF(range, criteria, [average_range])` | 한 조건을 만족하는 값의 평균 |
| AVERAGEIFS | `AVERAGEIFS(average_range, criteria_range1, criteria1, …)` | 여러 조건을 모두 만족하는 값의 평균 |
| CORREL | `CORREL(array1, array2)` | 상관 계수 |
| COUNT | `COUNT(value1, [value2], …)` | 숫자가 있는 셀 개수 |
| COUNTA | `COUNTA(value1, [value2], …)` | 비어 있지 않은 셀 개수 |
| COUNTBLANK | `COUNTBLANK(range)` | 빈 셀과 빈 문자열 개수 |
| COUNTIF | `COUNTIF(range, criteria)` | 한 조건을 만족하는 셀 개수 |
| COUNTIFS | `COUNTIFS(criteria_range1, criteria1, …)` | 여러 조건을 모두 만족하는 셀 개수 |
| COVARIANCE.P | `COVARIANCE.P(array1, array2)` | 모집단 공분산 |
| COVARIANCE.S | `COVARIANCE.S(array1, array2)` | 표본 공분산 |
| DEVSQ | `DEVSQ(number1, [number2], …)` | 평균과의 편차 제곱 합 |
| FISHER | `FISHER(x)` | 피셔 변환 |
| FISHERINV | `FISHERINV(y)` | 피셔 역변환 |
| FORECAST.LINEAR | `FORECAST.LINEAR(x, known_ys, known_xs)` | 선형 회귀 예측 |
| GEOMEAN | `GEOMEAN(number1, [number2], …)` | 기하 평균 |
| HARMEAN | `HARMEAN(number1, [number2], …)` | 조화 평균 |
| INTERCEPT | `INTERCEPT(known_ys, known_xs)` | 선형 회귀 절편 |
| KURT | `KURT(number1, [number2], …)` | 첨도 |
| LARGE | `LARGE(array, k)` | k번째로 큰 값 |
| MAX | `MAX(number1, [number2], …)` | 최댓값 |
| MAXA | `MAXA(value1, [value2], …)` | 논리값과 텍스트를 포함하는 최댓값 |
| MAXIFS | `MAXIFS(max_range, criteria_range1, criteria1, …)` | 여러 조건을 만족하는 최댓값 |
| MEDIAN | `MEDIAN(number1, [number2], …)` | 중앙값 |
| MIN | `MIN(number1, [number2], …)` | 최솟값 |
| MINA | `MINA(value1, [value2], …)` | 논리값과 텍스트를 포함하는 최솟값 |
| MINIFS | `MINIFS(min_range, criteria_range1, criteria1, …)` | 여러 조건을 만족하는 최솟값 |
| MODE.SNGL | `MODE.SNGL(number1, [number2], …)` | 최빈값 |
| PEARSON | `PEARSON(array1, array2)` | 피어슨 상관 계수 |
| PERCENTILE.EXC | `PERCENTILE.EXC(array, k)` | 양 끝을 제외한 백분위수 |
| PERCENTILE.INC | `PERCENTILE.INC(array, k)` | 양 끝을 포함한 백분위수 |
| PERCENTRANK.EXC | `PERCENTRANK.EXC(array, x, [significance])` | 양 끝을 제외한 백분율 순위 |
| PERCENTRANK.INC | `PERCENTRANK.INC(array, x, [significance])` | 양 끝을 포함한 백분율 순위 |
| PERMUT | `PERMUT(number, number_chosen)` | 순열 수 |
| PERMUTATIONA | `PERMUTATIONA(number, number_chosen)` | 중복 순열 수 |
| QUARTILE.EXC | `QUARTILE.EXC(array, quart)` | 양 끝을 제외한 사분위수 |
| QUARTILE.INC | `QUARTILE.INC(array, quart)` | 양 끝을 포함한 사분위수 |
| RANK.AVG | `RANK.AVG(number, ref, [order])` | 동률의 평균 순위 |
| RANK.EQ | `RANK.EQ(number, ref, [order])` | 동률의 최고 순위 |
| RSQ | `RSQ(known_ys, known_xs)` | 상관 계수의 제곱 |
| SKEW | `SKEW(number1, [number2], …)` | 표본 왜도 |
| SKEW.P | `SKEW.P(number1, [number2], …)` | 모집단 왜도 |
| SLOPE | `SLOPE(known_ys, known_xs)` | 선형 회귀 기울기 |
| SMALL | `SMALL(array, k)` | k번째로 작은 값 |
| STANDARDIZE | `STANDARDIZE(x, mean, standard_dev)` | 표준화 값 |
| STDEV.P | `STDEV.P(number1, [number2], …)` | 모집단 표준 편차 |
| STDEV.S | `STDEV.S(number1, [number2], …)` | 표본 표준 편차 |
| STDEVA | `STDEVA(value1, [value2], …)` | 논리값·텍스트를 포함하는 표본 표준 편차 |
| STDEVPA | `STDEVPA(value1, [value2], …)` | 논리값·텍스트를 포함하는 모집단 표준 편차 |
| STEYX | `STEYX(known_ys, known_xs)` | 회귀 예측의 표준 오차 |
| TRIMMEAN | `TRIMMEAN(array, percent)` | 양 끝의 일부를 제외한 평균 |
| VAR.P | `VAR.P(number1, [number2], …)` | 모집단 분산 |
| VAR.S | `VAR.S(number1, [number2], …)` | 표본 분산 |
| VARA | `VARA(value1, [value2], …)` | 논리값·텍스트를 포함하는 표본 분산 |
| VARPA | `VARPA(value1, [value2], …)` | 논리값·텍스트를 포함하는 모집단 분산 |

### 공학 (54)

| 함수 | 구문 | 용도 |
|---|---|---|
| BESSELI | `BESSELI(x, n)` | 수정 베셀 함수 I |
| BESSELJ | `BESSELJ(x, n)` | 베셀 함수 J |
| BESSELK | `BESSELK(x, n)` | 수정 베셀 함수 K |
| BESSELY | `BESSELY(x, n)` | 베셀 함수 Y |
| BIN2DEC | `BIN2DEC(number)` | 이진수를 십진수로 변환 |
| BIN2HEX | `BIN2HEX(number, [places])` | 이진수를 십육진수로 변환 |
| BIN2OCT | `BIN2OCT(number, [places])` | 이진수를 팔진수로 변환 |
| BITAND | `BITAND(number1, number2)` | 비트 AND |
| BITLSHIFT | `BITLSHIFT(number, shift_amount)` | 비트 왼쪽 이동 |
| BITOR | `BITOR(number1, number2)` | 비트 OR |
| BITRSHIFT | `BITRSHIFT(number, shift_amount)` | 비트 오른쪽 이동 |
| BITXOR | `BITXOR(number1, number2)` | 비트 XOR |
| COMPLEX | `COMPLEX(real_num, i_num, [suffix])` | 복소수 문자열 생성 |
| CONVERT | `CONVERT(number, from_unit, to_unit)` | 단위 변환 |
| DEC2BIN | `DEC2BIN(number, [places])` | 십진수를 이진수로 변환 |
| DEC2HEX | `DEC2HEX(number, [places])` | 십진수를 십육진수로 변환 |
| DEC2OCT | `DEC2OCT(number, [places])` | 십진수를 팔진수로 변환 |
| DELTA | `DELTA(number1, [number2])` | 두 수가 같으면 1 |
| ERF | `ERF(lower_limit, [upper_limit])` | 오차 함수 |
| ERF.PRECISE | `ERF.PRECISE(x)` | 오차 함수 |
| ERFC | `ERFC(x)` | 상보 오차 함수 |
| ERFC.PRECISE | `ERFC.PRECISE(x)` | 상보 오차 함수 |
| GESTEP | `GESTEP(number, [step])` | 기준값 이상이면 1 |
| HEX2BIN | `HEX2BIN(number, [places])` | 십육진수를 이진수로 변환 |
| HEX2DEC | `HEX2DEC(number)` | 십육진수를 십진수로 변환 |
| HEX2OCT | `HEX2OCT(number, [places])` | 십육진수를 팔진수로 변환 |
| IMABS | `IMABS(inumber)` | 복소수 절댓값 |
| IMAGINARY | `IMAGINARY(inumber)` | 복소수 허수부 |
| IMARGUMENT | `IMARGUMENT(inumber)` | 복소수 편각 |
| IMCONJUGATE | `IMCONJUGATE(inumber)` | 켤레 복소수 |
| IMCOS | `IMCOS(inumber)` | 복소수 코사인 |
| IMCOSH | `IMCOSH(inumber)` | 복소수 쌍곡 코사인 |
| IMCOT | `IMCOT(inumber)` | 복소수 코탄젠트 |
| IMCSC | `IMCSC(inumber)` | 복소수 코시컨트 |
| IMCSCH | `IMCSCH(inumber)` | 복소수 쌍곡 코시컨트 |
| IMDIV | `IMDIV(inumber1, inumber2)` | 복소수 나눗셈 |
| IMEXP | `IMEXP(inumber)` | 복소수 지수 |
| IMLN | `IMLN(inumber)` | 복소수 자연로그 |
| IMLOG10 | `IMLOG10(inumber)` | 복소수 상용로그 |
| IMLOG2 | `IMLOG2(inumber)` | 복소수 밑 2 로그 |
| IMPOWER | `IMPOWER(inumber, number)` | 복소수 거듭제곱 |
| IMPRODUCT | `IMPRODUCT(inumber1, [inumber2], …)` | 복소수 곱 |
| IMREAL | `IMREAL(inumber)` | 복소수 실수부 |
| IMSEC | `IMSEC(inumber)` | 복소수 시컨트 |
| IMSECH | `IMSECH(inumber)` | 복소수 쌍곡 시컨트 |
| IMSIN | `IMSIN(inumber)` | 복소수 사인 |
| IMSINH | `IMSINH(inumber)` | 복소수 쌍곡 사인 |
| IMSQRT | `IMSQRT(inumber)` | 복소수 제곱근 |
| IMSUB | `IMSUB(inumber1, inumber2)` | 복소수 뺄셈 |
| IMSUM | `IMSUM(inumber1, [inumber2], …)` | 복소수 합 |
| IMTAN | `IMTAN(inumber)` | 복소수 탄젠트 |
| OCT2BIN | `OCT2BIN(number, [places])` | 팔진수를 이진수로 변환 |
| OCT2DEC | `OCT2DEC(number)` | 팔진수를 십진수로 변환 |
| OCT2HEX | `OCT2HEX(number, [places])` | 팔진수를 십육진수로 변환 |

### 확률·분포 (40)

| 함수 | 구문 | 용도 |
|---|---|---|
| BETA.DIST | `BETA.DIST(x, alpha, beta, cumulative, [A], [B])` | 베타 분포 |
| BETA.INV | `BETA.INV(probability, alpha, beta, [A], [B])` | 베타 역누적 분포 |
| BINOM.DIST | `BINOM.DIST(number_s, trials, probability_s, cumulative)` | 이항 분포 |
| BINOM.DIST.RANGE | `BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])` | 이항 분포 구간 확률 |
| BINOM.INV | `BINOM.INV(trials, probability_s, alpha)` | 이항 분포 역함수 |
| CHISQ.DIST | `CHISQ.DIST(x, deg_freedom, cumulative)` | 카이 제곱 분포 |
| CHISQ.DIST.RT | `CHISQ.DIST.RT(x, deg_freedom)` | 카이 제곱 분포 오른쪽 꼬리 |
| CHISQ.INV | `CHISQ.INV(probability, deg_freedom)` | 카이 제곱 역누적 분포 |
| CHISQ.INV.RT | `CHISQ.INV.RT(probability, deg_freedom)` | 카이 제곱 오른쪽 꼬리 역함수 |
| CONFIDENCE.NORM | `CONFIDENCE.NORM(alpha, standard_dev, size)` | 정규 분포 신뢰 구간의 반폭 |
| CONFIDENCE.T | `CONFIDENCE.T(alpha, standard_dev, size)` | t 분포 신뢰 구간의 반폭 |
| EXPON.DIST | `EXPON.DIST(x, lambda, cumulative)` | 지수 분포 |
| F.DIST | `F.DIST(x, deg_freedom1, deg_freedom2, cumulative)` | F 분포 |
| F.DIST.RT | `F.DIST.RT(x, deg_freedom1, deg_freedom2)` | F 분포 오른쪽 꼬리 |
| F.INV | `F.INV(probability, deg_freedom1, deg_freedom2)` | F 역누적 분포 |
| F.INV.RT | `F.INV.RT(probability, deg_freedom1, deg_freedom2)` | F 분포 오른쪽 꼬리 역함수 |
| GAMMA | `GAMMA(number)` | 감마 함수 |
| GAMMA.DIST | `GAMMA.DIST(x, alpha, beta, cumulative)` | 감마 분포 |
| GAMMA.INV | `GAMMA.INV(probability, alpha, beta)` | 감마 역누적 분포 |
| GAMMALN | `GAMMALN(x)` | 감마 함수의 자연로그 |
| GAMMALN.PRECISE | `GAMMALN.PRECISE(x)` | 감마 함수의 자연로그 |
| GAUSS | `GAUSS(z)` | 표준 정규 누적 분포에서 0.5를 뺀 값 |
| HYPGEOM.DIST | `HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)` | 초기하 분포 |
| LOGNORM.DIST | `LOGNORM.DIST(x, mean, standard_dev, cumulative)` | 로그 정규 분포 |
| LOGNORM.INV | `LOGNORM.INV(probability, mean, standard_dev)` | 로그 정규 역누적 분포 |
| NEGBINOM.DIST | `NEGBINOM.DIST(number_f, number_s, probability_s, cumulative)` | 음이항 분포 |
| NORM.DIST | `NORM.DIST(x, mean, standard_dev, cumulative)` | 정규 분포 |
| NORM.INV | `NORM.INV(probability, mean, standard_dev)` | 정규 역누적 분포 |
| NORM.S.DIST | `NORM.S.DIST(z, cumulative)` | 표준 정규 분포 |
| NORM.S.INV | `NORM.S.INV(probability)` | 표준 정규 역누적 분포 |
| PHI | `PHI(x)` | 표준 정규 밀도 |
| POISSON.DIST | `POISSON.DIST(x, mean, cumulative)` | 포아송 분포 |
| PROB | `PROB(x_range, prob_range, lower_limit, [upper_limit])` | 지정한 구간의 확률 |
| T.DIST | `T.DIST(x, deg_freedom, cumulative)` | 스튜던트 t 분포 |
| T.DIST.2T | `T.DIST.2T(x, deg_freedom)` | t 분포 양쪽 꼬리 |
| T.DIST.RT | `T.DIST.RT(x, deg_freedom)` | t 분포 오른쪽 꼬리 |
| T.INV | `T.INV(probability, deg_freedom)` | t 역누적 분포 |
| T.INV.2T | `T.INV.2T(probability, deg_freedom)` | t 분포 양쪽 꼬리 역함수 |
| WEIBULL.DIST | `WEIBULL.DIST(x, alpha, beta, cumulative)` | 와이블 분포 |
| Z.TEST | `Z.TEST(array, x, [sigma])` | 단측 z 검정 |

### 이전 버전 호환 (30)

| 함수 | 구문 | 용도 |
|---|---|---|
| BETAINV | `BETAINV(probability, alpha, beta, [A], [B])` | BETA.INV의 이전 버전 이름 · 베타 역누적 분포 |
| BINOMDIST | `BINOMDIST(number_s, trials, probability_s, cumulative)` | BINOM.DIST의 이전 버전 이름 · 이항 분포 |
| CHIDIST | `CHIDIST(x, deg_freedom)` | CHISQ.DIST.RT의 이전 버전 이름 · 카이 제곱 분포 오른쪽 꼬리 |
| CHIINV | `CHIINV(probability, deg_freedom)` | CHISQ.INV.RT의 이전 버전 이름 · 카이 제곱 오른쪽 꼬리 역함수 |
| CONFIDENCE | `CONFIDENCE(alpha, standard_dev, size)` | CONFIDENCE.NORM의 이전 버전 이름 · 정규 분포 신뢰 구간의 반폭 |
| COVAR | `COVAR(array1, array2)` | COVARIANCE.P의 이전 버전 이름 · 모집단 공분산 |
| CRITBINOM | `CRITBINOM(trials, probability_s, alpha)` | BINOM.INV의 이전 버전 이름 · 이항 분포 역함수 |
| EXPONDIST | `EXPONDIST(x, lambda, cumulative)` | EXPON.DIST의 이전 버전 이름 · 지수 분포 |
| FDIST | `FDIST(x, deg_freedom1, deg_freedom2)` | F.DIST.RT의 이전 버전 이름 · F 분포 오른쪽 꼬리 |
| FINV | `FINV(probability, deg_freedom1, deg_freedom2)` | F.INV.RT의 이전 버전 이름 · F 분포 오른쪽 꼬리 역함수 |
| FORECAST | `FORECAST(x, known_ys, known_xs)` | FORECAST.LINEAR의 이전 버전 이름 · 선형 회귀 예측 |
| GAMMADIST | `GAMMADIST(x, alpha, beta, cumulative)` | GAMMA.DIST의 이전 버전 이름 · 감마 분포 |
| GAMMAINV | `GAMMAINV(probability, alpha, beta)` | GAMMA.INV의 이전 버전 이름 · 감마 역누적 분포 |
| LOGINV | `LOGINV(probability, mean, standard_dev)` | LOGNORM.INV의 이전 버전 이름 · 로그 정규 역누적 분포 |
| MODE | `MODE(number1, [number2], …)` | MODE.SNGL의 이전 버전 이름 · 최빈값 |
| NORMDIST | `NORMDIST(x, mean, standard_dev, cumulative)` | NORM.DIST의 이전 버전 이름 · 정규 분포 |
| NORMINV | `NORMINV(probability, mean, standard_dev)` | NORM.INV의 이전 버전 이름 · 정규 역누적 분포 |
| NORMSINV | `NORMSINV(probability)` | NORM.S.INV의 이전 버전 이름 · 표준 정규 역누적 분포 |
| PERCENTILE | `PERCENTILE(array, k)` | PERCENTILE.INC의 이전 버전 이름 · 양 끝을 포함한 백분위수 |
| PERCENTRANK | `PERCENTRANK(array, x, [significance])` | PERCENTRANK.INC의 이전 버전 이름 · 양 끝을 포함한 백분율 순위 |
| POISSON | `POISSON(x, mean, cumulative)` | POISSON.DIST의 이전 버전 이름 · 포아송 분포 |
| QUARTILE | `QUARTILE(array, quart)` | QUARTILE.INC의 이전 버전 이름 · 양 끝을 포함한 사분위수 |
| RANK | `RANK(number, ref, [order])` | RANK.EQ의 이전 버전 이름 · 동률의 최고 순위 |
| STDEV | `STDEV(number1, [number2], …)` | STDEV.S의 이전 버전 이름 · 표본 표준 편차 |
| STDEVP | `STDEVP(number1, [number2], …)` | STDEV.P의 이전 버전 이름 · 모집단 표준 편차 |
| TINV | `TINV(probability, deg_freedom)` | T.INV.2T의 이전 버전 이름 · t 분포 양쪽 꼬리 역함수 |
| VAR | `VAR(number1, [number2], …)` | VAR.S의 이전 버전 이름 · 표본 분산 |
| VARP | `VARP(number1, [number2], …)` | VAR.P의 이전 버전 이름 · 모집단 분산 |
| WEIBULL | `WEIBULL(x, alpha, beta, cumulative)` | WEIBULL.DIST의 이전 버전 이름 · 와이블 분포 |
| ZTEST | `ZTEST(array, x, [sigma])` | Z.TEST의 이전 버전 이름 · 단측 z 검정 |

### 텍스트 (30)

| 함수 | 구문 | 용도 |
|---|---|---|
| CHAR | `CHAR(number)` | 문자 코드로 문자 생성 |
| CLEAN | `CLEAN(text)` | 제어 문자 제거 |
| CODE | `CODE(text)` | 첫 문자의 코드 |
| CONCAT | `CONCAT(text1, [text2], …)` | 텍스트와 범위 연결 |
| CONCATENATE | `CONCATENATE(text1, [text2], …)` | 텍스트 연결 |
| DOLLAR | `DOLLAR(number, [decimals])` | 통화 형식의 텍스트 (원화) |
| EXACT | `EXACT(text1, text2)` | 대소문자를 구분하여 비교 |
| FIND | `FIND(find_text, within_text, [start_num])` | 대소문자를 구분하여 위치 찾기 |
| FIXED | `FIXED(number, [decimals], [no_commas])` | 고정 소수점 텍스트 |
| LEFT | `LEFT(text, [num_chars])` | 왼쪽 문자 추출 |
| LEN | `LEN(text)` | 문자 수 |
| LOWER | `LOWER(text)` | 소문자로 변환 |
| MID | `MID(text, start_num, num_chars)` | 중간 문자 추출 |
| NUMBERVALUE | `NUMBERVALUE(text, [decimal_separator], [group_separator])` | 구분자를 지정하여 텍스트를 숫자로 변환 |
| PROPER | `PROPER(text)` | 단어 첫 글자를 대문자로 변환 |
| REPLACE | `REPLACE(old_text, start_num, num_chars, new_text)` | 지정한 위치의 문자 바꾸기 |
| REPT | `REPT(text, number_times)` | 문자열 반복 |
| RIGHT | `RIGHT(text, [num_chars])` | 오른쪽 문자 추출 |
| SEARCH | `SEARCH(find_text, within_text, [start_num])` | 와일드카드로 대소문자 구분 없이 찾기 |
| SUBSTITUTE | `SUBSTITUTE(text, old_text, new_text, [instance_num])` | 일치하는 문자열 바꾸기 |
| T | `T(value)` | 텍스트만 반환 |
| TEXT | `TEXT(value, format_text)` | 숫자를 표시 형식에 맞는 텍스트로 변환 |
| TEXTAFTER | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | 구분자 뒤의 텍스트 |
| TEXTBEFORE | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | 구분자 앞의 텍스트 |
| TEXTJOIN | `TEXTJOIN(delimiter, ignore_empty, text1, [text2], …)` | 구분자를 넣어 텍스트 연결 |
| TRIM | `TRIM(text)` | 불필요한 공백 제거 |
| UNICHAR | `UNICHAR(number)` | 유니코드 코드 포인트로 문자 생성 |
| UNICODE | `UNICODE(text)` | 첫 문자의 유니코드 코드 포인트 |
| UPPER | `UPPER(text)` | 대문자로 변환 |
| VALUE | `VALUE(text)` | 텍스트를 숫자로 변환 |

### 재무 (26)

| 함수 | 구문 | 용도 |
|---|---|---|
| CUMIPMT | `CUMIPMT(rate, nper, pv, start_period, end_period, type)` | 기간 동안 납부한 이자 합계 |
| CUMPRINC | `CUMPRINC(rate, nper, pv, start_period, end_period, type)` | 기간 동안 상환한 원금 합계 |
| DB | `DB(cost, salvage, life, period, [month])` | 정률법 감가상각 |
| DDB | `DDB(cost, salvage, life, period, [factor])` | 이중 체감법 감가상각 |
| DOLLARDE | `DOLLARDE(fractional_dollar, fraction)` | 분수식 가격을 십진수로 변환 |
| DOLLARFR | `DOLLARFR(decimal_dollar, fraction)` | 십진수 가격을 분수식으로 변환 |
| EFFECT | `EFFECT(nominal_rate, npery)` | 실효 연이율 |
| FV | `FV(rate, nper, pmt, [pv], [type])` | 미래 가치 |
| FVSCHEDULE | `FVSCHEDULE(principal, schedule)` | 기간별 이율에 따른 미래 가치 |
| IPMT | `IPMT(rate, per, nper, pv, [fv], [type])` | 지정한 회차의 이자 |
| IRR | `IRR(values, [guess])` | 내부 수익률 |
| ISPMT | `ISPMT(rate, per, nper, pv)` | 균등 원금 상환의 이자 |
| MIRR | `MIRR(values, finance_rate, reinvest_rate)` | 수정 내부 수익률 |
| NOMINAL | `NOMINAL(effect_rate, npery)` | 명목 연이율 |
| NPER | `NPER(rate, pmt, pv, [fv], [type])` | 상환 또는 투자 기간 |
| NPV | `NPV(rate, value1, [value2], …)` | 순현재 가치 |
| PDURATION | `PDURATION(rate, pv, fv)` | 목표 가치까지 필요한 기간 |
| PMT | `PMT(rate, nper, pv, [fv], [type])` | 정기 상환 금액 |
| PPMT | `PPMT(rate, per, nper, pv, [fv], [type])` | 지정한 회차의 원금 |
| PV | `PV(rate, nper, pmt, [fv], [type])` | 현재 가치 |
| RATE | `RATE(nper, pmt, pv, [fv], [type], [guess])` | 기간당 이율 |
| RRI | `RRI(nper, pv, fv)` | 등가 성장률 |
| SLN | `SLN(cost, salvage, life)` | 정액법 감가상각 |
| SYD | `SYD(cost, salvage, life, per)` | 연수 합계법 감가상각 |
| XIRR | `XIRR(values, dates, [guess])` | 불규칙 현금 흐름의 내부 수익률 |
| XNPV | `XNPV(rate, values, dates)` | 불규칙 현금 흐름의 순현재 가치 |

### 날짜·시간 (25)

| 함수 | 구문 | 용도 |
|---|---|---|
| DATE | `DATE(year, month, day)` | 연·월·일을 Excel 날짜 번호로 변환 |
| DATEDIF | `DATEDIF(start_date, end_date, unit)` | 두 날짜의 연·월·일 차이 |
| DATEVALUE | `DATEVALUE(date_text)` | 날짜 문자열을 날짜 번호로 변환 |
| DAY | `DAY(serial_number)` | 날짜의 일 |
| DAYS | `DAYS(end_date, start_date)` | 두 날짜의 일수 차이 |
| DAYS360 | `DAYS360(start_date, end_date, [method])` | 1년을 360일로 계산한 날짜 차이 |
| EDATE | `EDATE(start_date, months)` | 지정한 개월 뒤의 날짜 |
| EOMONTH | `EOMONTH(start_date, months)` | 지정한 개월 뒤의 월말 |
| HOUR | `HOUR(serial_number)` | 시간의 시 |
| ISOWEEKNUM | `ISOWEEKNUM(date)` | ISO 기준 주 번호 |
| MINUTE | `MINUTE(serial_number)` | 시간의 분 |
| MONTH | `MONTH(serial_number)` | 날짜의 월 |
| NETWORKDAYS | `NETWORKDAYS(start_date, end_date, [holidays])` | 주말·공휴일을 제외한 근무일 수 |
| NETWORKDAYS.INTL | `NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])` | 주말을 지정한 근무일 수 |
| NOW | `NOW()` | 현재 날짜와 시간 |
| SECOND | `SECOND(serial_number)` | 시간의 초 |
| TIME | `TIME(hour, minute, second)` | 시·분·초를 하루의 비율로 변환 |
| TIMEVALUE | `TIMEVALUE(time_text)` | 시간 문자열을 하루의 비율로 변환 |
| TODAY | `TODAY()` | 오늘 날짜 |
| WEEKDAY | `WEEKDAY(serial_number, [return_type])` | 요일 번호 |
| WEEKNUM | `WEEKNUM(serial_number, [return_type])` | 연중 주 번호 |
| WORKDAY | `WORKDAY(start_date, days, [holidays])` | 근무일 기준으로 이동한 날짜 |
| WORKDAY.INTL | `WORKDAY.INTL(start_date, days, [weekend], [holidays])` | 주말을 지정해 근무일 기준으로 이동 |
| YEAR | `YEAR(serial_number)` | 날짜의 연도 |
| YEARFRAC | `YEARFRAC(start_date, end_date, [basis])` | 두 날짜 사이의 연도 비율 |

### 정보 (18)

| 함수 | 구문 | 용도 |
|---|---|---|
| ERROR.TYPE | `ERROR.TYPE(error_val)` | 오류의 종류 번호 |
| ISBLANK | `ISBLANK(value)` | 실제로 비어 있는 셀인지 검사 |
| ISERR | `ISERR(value)` | N/A 이외의 오류인지 검사 |
| ISERROR | `ISERROR(value)` | 오류인지 검사 |
| ISEVEN | `ISEVEN(number)` | 짝수인지 검사 |
| ISFORMULA | `ISFORMULA(reference)` | 수식이 들어 있는 셀인지 검사 |
| ISLOGICAL | `ISLOGICAL(value)` | 논리값인지 검사 |
| ISNA | `ISNA(value)` | N/A 오류인지 검사 |
| ISNONTEXT | `ISNONTEXT(value)` | 텍스트 이외의 값인지 검사 |
| ISNUMBER | `ISNUMBER(value)` | 숫자인지 검사 |
| ISODD | `ISODD(number)` | 홀수인지 검사 |
| ISREF | `ISREF(value)` | 셀 참조인지 검사 |
| ISTEXT | `ISTEXT(value)` | 텍스트인지 검사 |
| N | `N(value)` | 값을 숫자로 변환 |
| NA | `NA()` | N/A 오류 생성 |
| SHEET | `SHEET([value])` | 시트 순번 |
| SHEETS | `SHEETS([reference])` | 시트 개수 |
| TYPE | `TYPE(value)` | 값의 자료형 번호 |

