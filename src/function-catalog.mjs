/** Explicit, reviewed Excel names. Never expose arbitrary library properties to formulas.
 * Each example is also an executable compatibility fixture (see tests/functions.test.mjs).
 * Square brackets mark optional arguments; an ellipsis marks a repeated argument group.
 */
const groups = {
  '수학·삼각': `
ABS|number|-3|절댓값
ACOS|number|0.5|역코사인
ACOSH|number|2|역쌍곡 코사인
ACOT|number|2|역코탄젠트
ACOTH|number|2|역쌍곡 코탄젠트
ASIN|number|0.5|역사인
ASINH|number|2|역쌍곡 사인
ATAN|number|1|역탄젠트
ATAN2|x_num, y_num|1,1|좌표의 역탄젠트
ATANH|number|0.5|역쌍곡 탄젠트
BASE|number, radix, [min_length]|31,16,4|숫자를 다른 진법의 문자열로 변환
CEILING|number, significance|4.3,2|배수로 올림
CEILING.MATH|number, [significance], [mode]|-4.3,2,1|방향을 지정하여 배수로 올림
CEILING.PRECISE|number, [significance]|-4.3,2|배수로 올림
COMBIN|number, number_chosen|10,3|조합 수
COMBINA|number, number_chosen|4,3|중복 조합 수
COS|number|1|코사인
COSH|number|1|쌍곡 코사인
COT|number|1|코탄젠트
COTH|number|1|쌍곡 코탄젠트
CSC|number|1|코시컨트
CSCH|number|1|쌍곡 코시컨트
DECIMAL|text, radix|"FF",16|진법 문자열을 십진수로 변환
DEGREES|angle|PI()|라디안을 도로 변환
EVEN|number|-3|가장 가까운 짝수로 절댓값 올림
EXP|number|1|자연상수의 거듭제곱
FACT|number|6|계승
FACTDOUBLE|number|7|이중 계승
FLOOR|number, significance|4.3,2|배수로 내림
FLOOR.MATH|number, [significance], [mode]|-4.3,2,1|방향을 지정하여 배수로 내림
FLOOR.PRECISE|number, [significance]|-4.3,2|배수로 내림
GCD|number1, [number2], …|24,36|최대 공약수
INT|number|-2.1|정수로 내림
ISO.CEILING|number, [significance]|-4.3,2|배수로 올림
LCM|number1, [number2], …|6,8|최소 공배수
LN|number|EXP(2)|자연로그
LOG|number, [base]|8,2|지정한 밑의 로그
LOG10|number|1000|상용로그
MOD|number, divisor|-3,2|나눗셈의 나머지
MROUND|number, multiple|10,3|배수로 반올림
MULTINOMIAL|number1, [number2], …|2,3,4|다항 계수
ODD|number|-2|가장 가까운 홀수로 절댓값 올림
PI|||원주율
POWER|number, power|2,3|거듭제곱
PRODUCT|number1, [number2], …|2,3,4|곱
QUOTIENT|numerator, denominator|-7,3|나눗셈의 정수 부분
RADIANS|angle|180|도를 라디안으로 변환
RAND|||0 이상 1 미만의 난수
RANDBETWEEN|bottom, top|1,10|지정한 구간의 정수 난수
ROUND|number, num_digits|-1.25,1|반올림
ROUNDDOWN|number, num_digits|-1.29,1|절댓값 내림
ROUNDUP|number, num_digits|-1.21,1|절댓값 올림
SEC|number|1|시컨트
SECH|number|1|쌍곡 시컨트
SERIESSUM|x, n, m, coefficients|2,0,1,{1,2,3}|거듭제곱 급수의 합
SIGN|number|-4|숫자의 부호
SIN|number|1|사인
SINH|number|1|쌍곡 사인
SQRT|number|9|제곱근
SQRTPI|number|2|원주율을 곱한 수의 제곱근
SUBTOTAL|function_num, ref1, [ref2], …|9,A1:A5|중첩 부분합을 제외한 집계
SUM|number1, [number2], …|1,2,3|합계
SUMIF|range, criteria, [sum_range]|A1:A5,">2"|한 조건을 만족하는 값의 합계
SUMIFS|sum_range, criteria_range1, criteria1, …|A1:A5,A1:A5,">2",A1:A5,"<5"|여러 조건을 모두 만족하는 값의 합계
SUMPRODUCT|array1, [array2], …|{1,2,3},{4,5,6}|같은 위치의 값들을 곱한 뒤 합계
SUMSQ|number1, [number2], …|1,2,3|제곱의 합
SUMX2MY2|array_x, array_y|{1,2,3},{4,5,6}|제곱 차의 합
SUMX2PY2|array_x, array_y|{1,2,3},{4,5,6}|제곱 합의 합
SUMXMY2|array_x, array_y|{1,2,3},{4,5,6}|차를 제곱한 값의 합
TAN|number|1|탄젠트
TANH|number|1|쌍곡 탄젠트
TRUNC|number, [num_digits]|-12.345,2|지정한 자릿수에서 버림`,
  '통계': `
AVEDEV|number1, [number2], …|1,2,3,4,5|평균 절대 편차
AVERAGE|number1, [number2], …|1,2,3|산술 평균
AVERAGEA|value1, [value2], …|1,TRUE,3|논리값과 텍스트를 포함하는 평균
AVERAGEIF|range, criteria, [average_range]|A1:A5,">2"|한 조건을 만족하는 값의 평균
AVERAGEIFS|average_range, criteria_range1, criteria1, …|A1:A5,A1:A5,">2",A1:A5,"<5"|여러 조건을 모두 만족하는 값의 평균
CORREL|array1, array2|{1,2,3,4},{2,3,5,8}|상관 계수
COUNT|value1, [value2], …|1,2,"a"|숫자가 있는 셀 개수
COUNTA|value1, [value2], …|1,"a",TRUE|비어 있지 않은 셀 개수
COUNTBLANK|range|A1:A7|빈 셀과 빈 문자열 개수
COUNTIF|range, criteria|A1:A5,">2"|한 조건을 만족하는 셀 개수
COUNTIFS|criteria_range1, criteria1, …|A1:A5,">2",A1:A5,"<5"|여러 조건을 모두 만족하는 셀 개수
COVARIANCE.P|array1, array2|{1,2,3,4},{2,3,5,8}|모집단 공분산
COVARIANCE.S|array1, array2|{1,2,3,4},{2,3,5,8}|표본 공분산
DEVSQ|number1, [number2], …|1,2,3,4,5|평균과의 편차 제곱 합
FISHER|x|0.5|피셔 변환
FISHERINV|y|0.5|피셔 역변환
FORECAST.LINEAR|x, known_ys, known_xs|5,{2,4,6,8},{1,2,3,4}|선형 회귀 예측
GEOMEAN|number1, [number2], …|2,4,8|기하 평균
HARMEAN|number1, [number2], …|2,4,8|조화 평균
INTERCEPT|known_ys, known_xs|{2,3,5,8},{1,2,3,4}|선형 회귀 절편
KURT|number1, [number2], …|1,2,3,4,8|첨도
LARGE|array, k|{1,4,2,5,3},2|k번째로 큰 값
MAX|number1, [number2], …|1,2,3|최댓값
MAXA|value1, [value2], …|0,TRUE,0.5|논리값과 텍스트를 포함하는 최댓값
MAXIFS|max_range, criteria_range1, criteria1, …|A1:A5,A1:A5,"<4"|여러 조건을 만족하는 최댓값
MEDIAN|number1, [number2], …|1,2,5,8|중앙값
MIN|number1, [number2], …|1,2,3|최솟값
MINA|value1, [value2], …|2,FALSE,3|논리값과 텍스트를 포함하는 최솟값
MINIFS|min_range, criteria_range1, criteria1, …|A1:A5,A1:A5,">2"|여러 조건을 만족하는 최솟값
MODE.SNGL|number1, [number2], …|1,2,2,3|최빈값
PEARSON|array1, array2|{1,2,3,4},{2,3,5,8}|피어슨 상관 계수
PERCENTILE.EXC|array, k|{1,2,3,4,5},0.5|양 끝을 제외한 백분위수
PERCENTILE.INC|array, k|{1,2,3,4,5},0.25|양 끝을 포함한 백분위수
PERCENTRANK.EXC|array, x, [significance]|{1,2,3,4,5},3|양 끝을 제외한 백분율 순위
PERCENTRANK.INC|array, x, [significance]|{1,2,3,4,5},3|양 끝을 포함한 백분율 순위
PERMUT|number, number_chosen|10,3|순열 수
PERMUTATIONA|number, number_chosen|3,2|중복 순열 수
QUARTILE.EXC|array, quart|{1,2,3,4,5,6,7},1|양 끝을 제외한 사분위수
QUARTILE.INC|array, quart|{1,2,3,4,5},1|양 끝을 포함한 사분위수
RANK.AVG|number, ref, [order]|2,A1:A5,1|동률의 평균 순위
RANK.EQ|number, ref, [order]|2,A1:A5,1|동률의 최고 순위
RSQ|known_ys, known_xs|{2,3,5,8},{1,2,3,4}|상관 계수의 제곱
SKEW|number1, [number2], …|1,2,3,4,8|표본 왜도
SKEW.P|number1, [number2], …|1,2,3,4,8|모집단 왜도
SLOPE|known_ys, known_xs|{2,3,5,8},{1,2,3,4}|선형 회귀 기울기
SMALL|array, k|{1,4,2,5,3},2|k번째로 작은 값
STANDARDIZE|x, mean, standard_dev|4,2,2|표준화 값
STDEV.P|number1, [number2], …|1,2,3,4,5|모집단 표준 편차
STDEV.S|number1, [number2], …|1,2,3,4,5|표본 표준 편차
STDEVA|value1, [value2], …|1,TRUE,3|논리값·텍스트를 포함하는 표본 표준 편차
STDEVPA|value1, [value2], …|1,TRUE,3|논리값·텍스트를 포함하는 모집단 표준 편차
STEYX|known_ys, known_xs|{2,3,5,8},{1,2,3,4}|회귀 예측의 표준 오차
TRIMMEAN|array, percent|{1,2,3,4,100},0.4|양 끝의 일부를 제외한 평균
VAR.P|number1, [number2], …|1,2,3,4,5|모집단 분산
VAR.S|number1, [number2], …|1,2,3,4,5|표본 분산
VARA|value1, [value2], …|1,TRUE,3|논리값·텍스트를 포함하는 표본 분산
VARPA|value1, [value2], …|1,TRUE,3|논리값·텍스트를 포함하는 모집단 분산`,
  '확률·분포': `
BETA.DIST|x, alpha, beta, cumulative, [A], [B]|0.5,2,3,TRUE|베타 분포
BETA.INV|probability, alpha, beta, [A], [B]|0.5,2,3|베타 역누적 분포
BINOM.DIST|number_s, trials, probability_s, cumulative|2,10,0.3,FALSE|이항 분포
BINOM.DIST.RANGE|trials, probability_s, number_s, [number_s2]|10,0.3,2,4|이항 분포 구간 확률
BINOM.INV|trials, probability_s, alpha|10,0.3,0.5|이항 분포 역함수
CHISQ.DIST|x, deg_freedom, cumulative|3,4,TRUE|카이 제곱 분포
CHISQ.DIST.RT|x, deg_freedom|3,4|카이 제곱 분포 오른쪽 꼬리
CHISQ.INV|probability, deg_freedom|0.5,4|카이 제곱 역누적 분포
CHISQ.INV.RT|probability, deg_freedom|0.5,4|카이 제곱 오른쪽 꼬리 역함수
CONFIDENCE.NORM|alpha, standard_dev, size|0.05,2,100|정규 분포 신뢰 구간의 반폭
CONFIDENCE.T|alpha, standard_dev, size|0.05,2,100|t 분포 신뢰 구간의 반폭
EXPON.DIST|x, lambda, cumulative|2,0.5,TRUE|지수 분포
F.DIST|x, deg_freedom1, deg_freedom2, cumulative|2,5,6,TRUE|F 분포
F.DIST.RT|x, deg_freedom1, deg_freedom2|2,5,6|F 분포 오른쪽 꼬리
F.INV|probability, deg_freedom1, deg_freedom2|0.5,5,6|F 역누적 분포
F.INV.RT|probability, deg_freedom1, deg_freedom2|0.5,5,6|F 분포 오른쪽 꼬리 역함수
GAMMA|number|4|감마 함수
GAMMA.DIST|x, alpha, beta, cumulative|2,3,2,TRUE|감마 분포
GAMMA.INV|probability, alpha, beta|0.5,3,2|감마 역누적 분포
GAMMALN|x|4|감마 함수의 자연로그
GAMMALN.PRECISE|x|4|감마 함수의 자연로그
GAUSS|z|1|표준 정규 누적 분포에서 0.5를 뺀 값
HYPGEOM.DIST|sample_s, number_sample, population_s, number_pop, cumulative|1,4,8,20,FALSE|초기하 분포
LOGNORM.DIST|x, mean, standard_dev, cumulative|2,0,1,TRUE|로그 정규 분포
LOGNORM.INV|probability, mean, standard_dev|0.5,0,1|로그 정규 역누적 분포
NEGBINOM.DIST|number_f, number_s, probability_s, cumulative|3,2,0.5,FALSE|음이항 분포
NORM.DIST|x, mean, standard_dev, cumulative|1,0,1,TRUE|정규 분포
NORM.INV|probability, mean, standard_dev|0.8,0,1|정규 역누적 분포
NORM.S.DIST|z, cumulative|1,TRUE|표준 정규 분포
NORM.S.INV|probability|0.8|표준 정규 역누적 분포
PHI|x|1|표준 정규 밀도
POISSON.DIST|x, mean, cumulative|3,2,TRUE|포아송 분포
PROB|x_range, prob_range, lower_limit, [upper_limit]|{1,2,3},{0.2,0.3,0.5},1,2|지정한 구간의 확률
T.DIST|x, deg_freedom, cumulative|1,10,TRUE|스튜던트 t 분포
T.DIST.2T|x, deg_freedom|1,10|t 분포 양쪽 꼬리
T.DIST.RT|x, deg_freedom|1,10|t 분포 오른쪽 꼬리
T.INV|probability, deg_freedom|0.8,10|t 역누적 분포
T.INV.2T|probability, deg_freedom|0.2,10|t 분포 양쪽 꼬리 역함수
WEIBULL.DIST|x, alpha, beta, cumulative|2,3,4,TRUE|와이블 분포
Z.TEST|array, x, [sigma]|{1,2,3,4,5},2,1|단측 z 검정`,
  '재무': `
CUMIPMT|rate, nper, pv, start_period, end_period, type|0.01,12,1000,1,6,0|기간 동안 납부한 이자 합계
CUMPRINC|rate, nper, pv, start_period, end_period, type|0.01,12,1000,1,6,0|기간 동안 상환한 원금 합계
DB|cost, salvage, life, period, [month]|10000,1000,5,1|정률법 감가상각
DDB|cost, salvage, life, period, [factor]|10000,1000,5,1|이중 체감법 감가상각
DOLLARDE|fractional_dollar, fraction|1.02,16|분수식 가격을 십진수로 변환
DOLLARFR|decimal_dollar, fraction|1.125,16|십진수 가격을 분수식으로 변환
EFFECT|nominal_rate, npery|0.12,12|실효 연이율
FV|rate, nper, pmt, [pv], [type]|0.01,12,-100,-1000|미래 가치
FVSCHEDULE|principal, schedule|1000,{0.1,0.2}|기간별 이율에 따른 미래 가치
IPMT|rate, per, nper, pv, [fv], [type]|0.01,1,12,1000|지정한 회차의 이자
IRR|values, [guess]|{-1000,300,400,500}|내부 수익률
ISPMT|rate, per, nper, pv|0.1,1,3,1000|균등 원금 상환의 이자
MIRR|values, finance_rate, reinvest_rate|{-1000,300,400,500},0.1,0.12|수정 내부 수익률
NOMINAL|effect_rate, npery|0.12,12|명목 연이율
NPER|rate, pmt, pv, [fv], [type]|0.01,-100,1000|상환 또는 투자 기간
NPV|rate, value1, [value2], …|0.1,-1000,300,400,500|순현재 가치
PDURATION|rate, pv, fv|0.1,1000,2000|목표 가치까지 필요한 기간
PMT|rate, nper, pv, [fv], [type]|0.01,12,1000|정기 상환 금액
PPMT|rate, per, nper, pv, [fv], [type]|0.01,1,12,1000|지정한 회차의 원금
PV|rate, nper, pmt, [fv], [type]|0.01,12,-100|현재 가치
RATE|nper, pmt, pv, [fv], [type], [guess]|12,-100,1000|기간당 이율
RRI|nper, pv, fv|10,1000,2000|등가 성장률
SLN|cost, salvage, life|10000,1000,5|정액법 감가상각
SYD|cost, salvage, life, per|10000,1000,5,1|연수 합계법 감가상각
XIRR|values, dates, [guess]|{-1000,1100},{43831,44197}|불규칙 현금 흐름의 내부 수익률
XNPV|rate, values, dates|0.1,{-1000,1100},{43831,44197}|불규칙 현금 흐름의 순현재 가치`,
  '날짜·시간': `
DATE|year, month, day|2024,2,29|연·월·일을 Excel 날짜 번호로 변환
DATEDIF|start_date, end_date, unit|DATE(2020,1,15),DATE(2024,3,20),"M"|두 날짜의 연·월·일 차이
DATEVALUE|date_text|"2024-02-29"|날짜 문자열을 날짜 번호로 변환
DAY|serial_number|DATE(2024,2,29)|날짜의 일
DAYS|end_date, start_date|DATE(2024,3,1),DATE(2024,2,28)|두 날짜의 일수 차이
DAYS360|start_date, end_date, [method]|DATE(2024,1,1),DATE(2024,3,1),FALSE|1년을 360일로 계산한 날짜 차이
EDATE|start_date, months|DATE(2024,1,31),1|지정한 개월 뒤의 날짜
EOMONTH|start_date, months|DATE(2024,1,10),1|지정한 개월 뒤의 월말
HOUR|serial_number|TIME(13,24,35)|시간의 시
ISOWEEKNUM|date|DATE(2021,1,1)|ISO 기준 주 번호
MINUTE|serial_number|TIME(13,24,35)|시간의 분
MONTH|serial_number|DATE(2024,2,29)|날짜의 월
NETWORKDAYS|start_date, end_date, [holidays]|DATE(2024,1,1),DATE(2024,1,7)|주말·공휴일을 제외한 근무일 수
NETWORKDAYS.INTL|start_date, end_date, [weekend], [holidays]|DATE(2024,1,1),DATE(2024,1,7),11|주말을 지정한 근무일 수
NOW|||현재 날짜와 시간
SECOND|serial_number|TIME(13,24,35)|시간의 초
TIME|hour, minute, second|13,24,35|시·분·초를 하루의 비율로 변환
TIMEVALUE|time_text|"13:24:35"|시간 문자열을 하루의 비율로 변환
TODAY|||오늘 날짜
WEEKDAY|serial_number, [return_type]|DATE(2024,1,1),2|요일 번호
WEEKNUM|serial_number, [return_type]|DATE(2024,1,1),2|연중 주 번호
WORKDAY|start_date, days, [holidays]|DATE(2024,1,5),2|근무일 기준으로 이동한 날짜
WORKDAY.INTL|start_date, days, [weekend], [holidays]|DATE(2024,1,5),2,11|주말을 지정해 근무일 기준으로 이동
YEAR|serial_number|DATE(2024,2,29)|날짜의 연도
YEARFRAC|start_date, end_date, [basis]|DATE(2024,1,1),DATE(2024,7,1),3|두 날짜 사이의 연도 비율`,
  '텍스트': `
CHAR|number|65|문자 코드로 문자 생성
CLEAN|text|"a"&CHAR(10)&"b"|제어 문자 제거
CODE|text|"A"|첫 문자의 코드
CONCAT|text1, [text2], …|"안녕"," ","Excel"|텍스트와 범위 연결
CONCATENATE|text1, [text2], …|"안녕"," ","Excel"|텍스트 연결
DOLLAR|number, [decimals]|1234.567,2|통화 형식의 텍스트 (원화)
EXACT|text1, text2|"Abc","abc"|대소문자를 구분하여 비교
FIND|find_text, within_text, [start_num]|"b","abcabc",3|대소문자를 구분하여 위치 찾기
FIXED|number, [decimals], [no_commas]|1234.567,2,FALSE|고정 소수점 텍스트
LEFT|text, [num_chars]|"abcdef",2|왼쪽 문자 추출
LEN|text|"안녕하세요"|문자 수
LOWER|text|"AbC"|소문자로 변환
MID|text, start_num, num_chars|"abcdef",2,3|중간 문자 추출
NUMBERVALUE|text, [decimal_separator], [group_separator]|"1.234,5",",","."|구분자를 지정하여 텍스트를 숫자로 변환
PROPER|text|"hello WORLD"|단어 첫 글자를 대문자로 변환
REPLACE|old_text, start_num, num_chars, new_text|"abcdef",2,3,"X"|지정한 위치의 문자 바꾸기
REPT|text, number_times|"ab",3|문자열 반복
RIGHT|text, [num_chars]|"abcdef",2|오른쪽 문자 추출
SEARCH|find_text, within_text, [start_num]|"B?","abcabc",1|와일드카드로 대소문자 구분 없이 찾기
SUBSTITUTE|text, old_text, new_text, [instance_num]|"a-a-a","a","b",2|일치하는 문자열 바꾸기
T|value|"text"|텍스트만 반환
TEXT|value, format_text|1234.567,"#,##0.00"|숫자를 표시 형식에 맞는 텍스트로 변환
TEXTAFTER|text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]|"a-b-c","-",-1|구분자 뒤의 텍스트
TEXTBEFORE|text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]|"a-b-c","-",-1|구분자 앞의 텍스트
TEXTJOIN|delimiter, ignore_empty, text1, [text2], …|", ",TRUE,"a","","b"|구분자를 넣어 텍스트 연결
TRIM|text|"  one   two  "|불필요한 공백 제거
UNICHAR|number|9731|유니코드 코드 포인트로 문자 생성
UNICODE|text|"☃"|첫 문자의 유니코드 코드 포인트
UPPER|text|"AbC"|대문자로 변환
VALUE|text|"1,234.5"|텍스트를 숫자로 변환`,
  '논리': `
AND|logical1, [logical2], …|TRUE,3>2|모든 조건이 참인지 검사
FALSE|||논리값 FALSE
IF|logical_test, value_if_true, [value_if_false]|TRUE,42,1/0|조건에 따라 값 선택
IFERROR|value, value_if_error|1/0,"오류"|오류일 때 대체 값
IFNA|value, value_if_na|NA(),"없음"|N/A 오류일 때 대체 값
IFS|logical_test1, value_if_true1, …|FALSE,1,TRUE,2|첫 번째 참인 조건의 값
LET|name1, name_value1, calculation, …|x,5,SUM(x,1)|수식 안에서 이름을 붙인 계산을 재사용
NOT|logical|FALSE|논리값 반전
OR|logical1, [logical2], …|FALSE,TRUE|하나 이상의 조건이 참인지 검사
SWITCH|expression, value1, result1, …|2,1,"a",2,"b","other"|값에 따라 결과 선택 (마지막 기본값 선택 가능)
TRUE|||논리값 TRUE
XOR|logical1, [logical2], …|TRUE,FALSE,TRUE|참인 조건의 수가 홀수인지 검사`,
  '정보': `
ERROR.TYPE|error_val|1/0|오류의 종류 번호
ISBLANK|value|A7|실제로 비어 있는 셀인지 검사
ISERR|value|1/0|N/A 이외의 오류인지 검사
ISERROR|value|1/0|오류인지 검사
ISEVEN|number|4|짝수인지 검사
ISFORMULA|reference|B1|수식이 들어 있는 셀인지 검사
ISLOGICAL|value|TRUE|논리값인지 검사
ISNA|value|NA()|N/A 오류인지 검사
ISNONTEXT|value|42|텍스트 이외의 값인지 검사
ISNUMBER|value|42|숫자인지 검사
ISODD|number|3|홀수인지 검사
ISREF|value|A1|셀 참조인지 검사
ISTEXT|value|"42"|텍스트인지 검사
N|value|TRUE|값을 숫자로 변환
NA|||N/A 오류 생성
SHEET|[value]|"Sheet1"|시트 순번
SHEETS|[reference]||시트 개수
TYPE|value|TRUE|값의 자료형 번호`,
  '찾기·참조': `
FILTER|array, include, [if_empty]|{1;2},{TRUE;FALSE}|조건이 참인 행이나 열을 동적 배열로 반환
SORT|array, [sort_index], [sort_order], [by_col]|{2;1}|값을 정렬한 동적 배열 반환
UNIQUE|array, [by_col], [exactly_once]|{1;1;2}|중복을 제거한 행이나 열 반환
SEQUENCE|rows, [columns], [start], [step]|3|지정한 크기와 간격의 연속 숫자 배열 생성
ADDRESS|row_num, column_num, [abs_num], [a1], [sheet_text]|3,2,4,TRUE,"Sheet1"|셀 주소 문자열 생성
CHOOSE|index_num, value1, [value2], …|2,"a","b","c"|순번으로 값 선택
COLUMN|[reference]|C3|열 번호
COLUMNS|array|A1:C5|열 개수
FORMULATEXT|reference|B1|셀의 수식을 텍스트로 반환
HLOOKUP|lookup_value, table_array, row_index_num, [range_lookup]|2,{1,2,3;10,20,30},2,FALSE|첫 행에서 가로로 찾아 반환
INDEX|array, row_num, [column_num]|{1,2;3,4},2,2|범위의 지정한 위치 값 또는 행·열 참조
INDIRECT|ref_text, [a1]|"A"&3|문자열을 셀 참조로 변환
LOOKUP|lookup_value, lookup_vector, [result_vector]|2.5,{1,2,3},{10,20,30}|정렬된 범위에서 근삿값 찾기
MATCH|lookup_value, lookup_array, [match_type]|2,{1,2,3},0|일치하는 값의 위치
OFFSET|reference, rows, cols, [height], [width]|A1,1,0|셀에서 행·열만큼 이동한 참조
ROW|[reference]|C3|행 번호
ROWS|array|A1:C5|행 개수
VLOOKUP|lookup_value, table_array, col_index_num, [range_lookup]|2,{1,10;2,20;3,30},2,FALSE|첫 열에서 세로로 찾아 반환
XLOOKUP|lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]|2,{1,2,3},{10,20,30}|방향·일치 방식을 지정하여 찾기
XMATCH|lookup_value, lookup_array, [match_mode], [search_mode]|2,{1,2,3},0|방향·일치 방식을 지정하여 위치 찾기`,
  '공학': `
BESSELI|x, n|1.5,1|수정 베셀 함수 I
BESSELJ|x, n|1.5,1|베셀 함수 J
BESSELK|x, n|1.5,1|수정 베셀 함수 K
BESSELY|x, n|1.5,1|베셀 함수 Y
BIN2DEC|number|"1010"|이진수를 십진수로 변환
BIN2HEX|number, [places]|"1010",4|이진수를 십육진수로 변환
BIN2OCT|number, [places]|"1010",4|이진수를 팔진수로 변환
BITAND|number1, number2|13,7|비트 AND
BITLSHIFT|number, shift_amount|5,2|비트 왼쪽 이동
BITOR|number1, number2|13,7|비트 OR
BITRSHIFT|number, shift_amount|20,2|비트 오른쪽 이동
BITXOR|number1, number2|13,7|비트 XOR
COMPLEX|real_num, i_num, [suffix]|3,4|복소수 문자열 생성
CONVERT|number, from_unit, to_unit|1,"mi","km"|단위 변환
DEC2BIN|number, [places]|10,8|십진수를 이진수로 변환
DEC2HEX|number, [places]|255,4|십진수를 십육진수로 변환
DEC2OCT|number, [places]|64,4|십진수를 팔진수로 변환
DELTA|number1, [number2]|3,3|두 수가 같으면 1
ERF|lower_limit, [upper_limit]|0,1|오차 함수
ERF.PRECISE|x|1|오차 함수
ERFC|x|1|상보 오차 함수
ERFC.PRECISE|x|1|상보 오차 함수
GESTEP|number, [step]|3,2|기준값 이상이면 1
HEX2BIN|number, [places]|"A",8|십육진수를 이진수로 변환
HEX2DEC|number|"FF"|십육진수를 십진수로 변환
HEX2OCT|number, [places]|"FF",4|십육진수를 팔진수로 변환
IMABS|inumber|"3+4i"|복소수 절댓값
IMAGINARY|inumber|"3+4i"|복소수 허수부
IMARGUMENT|inumber|"3+4i"|복소수 편각
IMCONJUGATE|inumber|"3+4i"|켤레 복소수
IMCOS|inumber|"1+i"|복소수 코사인
IMCOSH|inumber|"1+i"|복소수 쌍곡 코사인
IMCOT|inumber|"1+i"|복소수 코탄젠트
IMCSC|inumber|"1+i"|복소수 코시컨트
IMCSCH|inumber|"1+i"|복소수 쌍곡 코시컨트
IMDIV|inumber1, inumber2|"3+4i","1+i"|복소수 나눗셈
IMEXP|inumber|"1+i"|복소수 지수
IMLN|inumber|"1+i"|복소수 자연로그
IMLOG10|inumber|"1+i"|복소수 상용로그
IMLOG2|inumber|"1+i"|복소수 밑 2 로그
IMPOWER|inumber, number|"1+i",2|복소수 거듭제곱
IMPRODUCT|inumber1, [inumber2], …|"1+i","2+i"|복소수 곱
IMREAL|inumber|"3+4i"|복소수 실수부
IMSEC|inumber|"1+i"|복소수 시컨트
IMSECH|inumber|"1+i"|복소수 쌍곡 시컨트
IMSIN|inumber|"1+i"|복소수 사인
IMSINH|inumber|"1+i"|복소수 쌍곡 사인
IMSQRT|inumber|"3+4i"|복소수 제곱근
IMSUB|inumber1, inumber2|"3+4i","1+i"|복소수 뺄셈
IMSUM|inumber1, [inumber2], …|"1+i","2+i"|복소수 합
IMTAN|inumber|"1+i"|복소수 탄젠트
OCT2BIN|number, [places]|"12",8|팔진수를 이진수로 변환
OCT2DEC|number|"77"|팔진수를 십진수로 변환
OCT2HEX|number, [places]|"77",4|팔진수를 십육진수로 변환`,
};

export const FUNCTION_ALIASES = Object.freeze({
  STDEV:'STDEV.S', STDEVP:'STDEV.P', VAR:'VAR.S', VARP:'VAR.P', MODE:'MODE.SNGL',
  RANK:'RANK.EQ', PERCENTILE:'PERCENTILE.INC', PERCENTRANK:'PERCENTRANK.INC',
  QUARTILE:'QUARTILE.INC', FORECAST:'FORECAST.LINEAR', COVAR:'COVARIANCE.P',
  BETAINV:'BETA.INV', BINOMDIST:'BINOM.DIST', CHIDIST:'CHISQ.DIST.RT', CHIINV:'CHISQ.INV.RT',
  CONFIDENCE:'CONFIDENCE.NORM', CRITBINOM:'BINOM.INV', EXPONDIST:'EXPON.DIST',
  FDIST:'F.DIST.RT', FINV:'F.INV.RT', GAMMADIST:'GAMMA.DIST', GAMMAINV:'GAMMA.INV',
  LOGINV:'LOGNORM.INV', NORMDIST:'NORM.DIST', NORMINV:'NORM.INV', NORMSINV:'NORM.S.INV',
  POISSON:'POISSON.DIST', TINV:'T.INV.2T', WEIBULL:'WEIBULL.DIST', ZTEST:'Z.TEST',
});

const entries = Object.entries(groups).flatMap(([category, data]) => data.trim().split('\n').map(line => {
  const [name, parameters, exampleArgs, description] = line.split('|');
  const parts = parameters ? parameters.split(', ').filter(p => p !== '…') : [];
  return { name, category, parameters, syntax:`${name}(${parameters})`, description,
    example:`=${name}(${exampleArgs})`, min:parts.filter(p => !p.startsWith('[')).length,
    max:name==='LET'?253:parameters.includes('…') ? 255 : parts.length };
}));
for (const [name, target] of Object.entries(FUNCTION_ALIASES)) {
  const base = entries.find(entry => entry.name === target);
  entries.push({...base, name, category:'이전 버전 호환', syntax:base.syntax.replace(target,name),
    example:base.example.replace(target,name), description:`${target}의 이전 버전 이름 · ${base.description}`, aliasOf:target});
}
export const FUNCTION_CATALOG = Object.freeze(entries.sort((a,b) => a.name.localeCompare(b.name,'en')).map(Object.freeze));
export const FUNCTION_MAP = new Map(FUNCTION_CATALOG.map(entry => [entry.name,entry]));
// MS-XLSX future functions need this prefix in the XML even when Excel's formula bar
// displays the ordinary name. Existing/unknown _xlfn functions are left untouched.
export const XLSX_FUTURE_FUNCTIONS = new Set(FUNCTION_CATALOG.filter(f=>f.name.includes('.')&&f.name!=='ERROR.TYPE').map(f=>f.name).concat(`
ACOT ACOTH BASE BITAND BITLSHIFT BITOR BITRSHIFT BITXOR COMBINA CONCAT COT COTH CSC CSCH
DAYS DECIMAL FORMULATEXT GAMMA GAUSS IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH
IMSINH IMTAN ISFORMULA ISOWEEKNUM MAXIFS MINIFS NUMBERVALUE PDURATION PERMUTATIONA PHI
FILTER SORT UNIQUE SEQUENCE ANCHORARRAY SINGLE LET RRI SEC SECH SHEET SHEETS SWITCH TEXTAFTER TEXTBEFORE TEXTJOIN UNICHAR UNICODE XLOOKUP XMATCH XOR
`.trim().split(/\s+/)));
export function normalizeFunctionName(name) {
  return String(name).toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?/, '').replace(/^_XLWS\./,'');
}
export function searchFunctions(query = '', category = '') {
  const words = query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
  return FUNCTION_CATALOG.filter(f => (!category || f.category === category) && words.every(word => `${f.name} ${f.description} ${f.category}`.toLocaleLowerCase().includes(word)));
}
