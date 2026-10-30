# 추가 핵심 개선 — 5시간 작업 계획

시작: **2026-10-31 01:11 KST**, 목표 종료: **06:11 KST**. 기준 커밋 `48ee754`, 앱 0.7.30. [기존 계획](CORE-IMPROVEMENT-PLAN.md)을 이어서 진행한다. 사용자 문서와 실제 클립보드는 검증용으로 변경하지 않는다. 공개 릴리즈·리토스는 별도의 “배포해줘” 요청 전까지 갱신하지 않는다.

## 시간 배분과 완료 기준

| 구간 | 우선 작업 | 확인할 결과 |
|---|---|---|
| 01:11–02:11 | 계산·파일 재현과 동적 배열 함수 | FILTER/SORT/UNIQUE/SEQUENCE의 가로·세로·2차원, 중첩, 빈 결과/오류, 선택 인수와 크기 한도 |
| 02:11–03:11 | 동적 스필과 참조 | 결과 확장·축소, 원본 변경, `A1#`, 막힌 셀·병합·표·시트 끝의 오류, 기존 고정 배열과의 구분 |
| 03:11–04:11 | 편집·파일 연결 | 결과 셀 보호, 시작 셀 편집/삭제, 복사, 실행 취소·다시 실행, XLSX 동적 배열 정보와 계산 캐시의 반복 왕복 |
| 04:11–05:11 | 인쇄·PDF 후속 | 머리글·바닥글 또는 페이지 설정 편집을 원본 보존과 함께 연결하고 PDF 실제 렌더링으로 확인 |
| 05:11–06:11 | 회귀·대용량·Git·앱 | 관련 전체 검사와 3개 OS CI, 큰 배열의 시간·메모리 한도, 로컬 패키지 검증, Git 반영, 잠금이 해제되어 있으면 앱 갱신·실행 |

시간 배분은 예상이며 결함 재현 결과에 따라 조정한다. 동적 배열은 엔진·편집·파일 저장을 함께 완성하는 것을 우선한다. 반쪽 기능을 노출해 기존 문서의 계산이나 저장을 바꾸지 않는다. 실제 Microsoft Excel/Windows VBA의 새 실행 검증은 현재 환경에서 할 수 없으며, 공식 규칙·기존 Excel 측정 자료·OOXML 검사·독립 엔진 비교의 근거를 구분한다.

## 주요 설계 원칙

- 시작 셀 수식만 사용자 원본으로 저장하고 파생 결과는 계산 세션에서 관리해 대용량 기록을 불필요하게 복제하지 않는다.
- 이전 CSE 고정 배열과 동적 배열의 파일 표식을 구분한다. 저장 과정에서 동적 배열을 고정 배열로 조용히 바꾸지 않는다.
- 지원하지 않는 XLSM 작업은 원본 매크로·확장 정보를 보존하는 기존 정책을 따른다.
- 발견한 문제는 실패 사례 → 원인 → 수정 → 같은 시나리오 재검사 순서로 아래에 기록한다.

## 진행

- [x] 기준 커밋·원격 상태·기존 계획 확인, 독립 작업용 체크아웃 준비
- [x] 추가 계획과 시간 배분 기록
- [x] 함수와 배열 엔진
- [x] 편집·복사·XLSX 왕복
- [x] 인쇄 후속
- [ ] 대용량·플랫폼 회귀
- [ ] Git 반영·검증한 로컬 앱 준비

## 재현과 수정 기록

착수 시 FILTER/SORT/UNIQUE/SEQUENCE와 `A1#` 스필 참조는 미지원이다. 원본 고정 배열을 읽는 기능은 있으나 새 배열 수식 결과를 주변 셀에 동적으로 표시하지 않는다.

### 1차 구현: 동적 배열과 파일 캐시

- FILTER/SORT/UNIQUE/SEQUENCE와 스필 확장·축소, `A1#`, `@` 교차 선택을 연결했다. XLOOKUP은 일치한 행의 여러 열 또는 열의 여러 행을 반환한다. 기존 CSE는 고정 범위를 유지한다.
- 결과 셀 일부 덮어쓰기를 막고 첫 셀 편집·삭제, 복사(원본 포함 시 수식 / 결과만 선택 시 값), 실행 취소·다시 실행, 영역 이동, 합계 표시와 인쇄를 연결했다. 스필 경계와 막힌 이유를 표시한다.
- XLSX의 XLDAPR/동적 배열 메타데이터와 미래 함수 접두사, ANCHORARRAY/SINGLE 표현을 왕복한다. 독립 XlsxWriter 생성 파일과 기존 CSE를 구분하고 원본의 다른 cm/vm 기록·연결 파일을 보존한다.
- 지원하지 않는 수식의 원본 캐시는 표시 근거를 구분한다. 셀 값이 바뀌면 오래된 캐시를 새 계산값처럼 쓰지 않으며 실행 취소가 그 상태도 복원한다. 동적 배열의 오래된 원본 캐시는 파일에 보존하되 이 앱의 저장/재열기에서도 오래되었다는 표식을 유지한다.
- 계산 가능한 가져온 스필의 값 전용 캐시 셀을 정리한다. 스타일·메모·원본 셀 메타데이터는 유지해, 큰 배열을 줄인 뒤 편집 셀 한도가 불필요하게 차 있는 문제를 피한다. XLSM은 원본 보존을 위해 이 정리에서 제외한다.

추가 재현: 잘못된 방향으로 `A1#` 수식을 복사하면 `#REF!#`가 만들어졌다. 참조가 무효화될 때 스필 연산자도 소비해 `#REF!`로 유지하도록 복사·전치·행/열 삭제·시트 삭제를 수정했다. 또 원본이 오래된 미지원 배열을 `SUM(A1#)`로 읽으면 옛 합계가 살아나는 경로를 차단했다.

기준: [Microsoft 동적 배열](https://support.microsoft.com/en-us/excel/dynamic-array-formulas-and-spilled-array-behavior), [FILTER](https://support.microsoft.com/en-us/excel/functions/filter-function), [SORT](https://support.microsoft.com/en-us/excel/functions/sort-function), [UNIQUE](https://support.microsoft.com/en-us/excel/functions/unique-function), [SEQUENCE](https://support.microsoft.com/en-us/excel/functions/sequence-function), [스필 참조](https://support.microsoft.com/en-us/excel/spilled-range-operator), [XLOOKUP](https://support.microsoft.com/en-us/excel/functions/xlookup-function), [OOXML 동적 배열 속성](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/d1ef676d-d970-4ddd-b642-62f5e0c291e0), [XlsxWriter 배열 저장](https://xlsxwriter.readthedocs.io/working_with_formulas.html).

로컬 검증: 관련 단위 검사, 실제 Electron의 직접 편집·원본 변경·실행 취소/다시 실행·부분/전체 복사·저장/재열기, CSE 회귀, 10,000셀 의존 체인과 F9 회귀를 확인했다. LibreOffice 26.8의 별도 프로필에서 13개 XLSX 사례에 의도적으로 틀린 캐시를 넣고 다시 계산해 배열의 모든 원소가 일치했다. 이는 새 Microsoft Excel 실행 검증을 대신하지 않는다. 100만 결과의 계산·XLSX 저장·재열기 성능 수치와 최종 플랫폼 결과는 마지막 검증 후 기록한다.

한도: 파생 결과 합계 1,000,000개, 한 시트의 활성 배열 정의 10,000개, XLSX 기록 합계 1,000,000개다. 그 이상은 오류를 표시한다. 모든 최신 Excel 배열 함수나 LAMBDA, 복수 검색값과 복수 반환 열을 동시에 조합하는 XLOOKUP 전체를 구현한 것은 아니다. 지원하지 않는 rich-data/사용자 함수는 원본 보존과 캐시 정책을 적용한다.


### 2차 구현: 페이지 설정과 실제 PDF

페이지 레이아웃에 페이지 설정 창을 추가했다. 용지·방향·백분율/페이지 맞춤·시작 페이지 번호·페이지 순서·눈금선, 여백, 인쇄 영역·반복 행/열, 기본/첫/짝수 페이지의 머리글·바닥글을 편집한다. 적용 전 검증은 원본을 변경하지 않으며 한 번의 실행 취소로 되돌릴 수 있다. 잘못된 저장 인쇄 참조도 설정 창에서 보고 수정할 수 있다. XLSM의 원본 설정 변경은 기존 원본 보존 정책에 따라 제한한다.

머리글·바닥글은 왼쪽/가운데/오른쪽 구역, 페이지/전체 페이지/날짜/시간/파일·시트 이름, 글꼴·크기·굵게·기울임·밑줄·색상과 첫/짝수 페이지 구분을 출력한다. 가져온 서식은 해당 텍스트 구역을 바꾸지 않으면 유지한다. 첫 페이지 또는 짝수 페이지의 명시적인 빈 머리글과, 원본에 짝수 머리글이 없는 경우를 구분한다. 원본 경로와 머리글 그림 등 표시할 수 없는 정보는 실제 출력의 제한을 알리고 파일에는 보존한다.

PDF 재현 과정에서 여백에 배치한 바닥글이 추가 페이지를 만드는 문제를 잡았다. 용지 전체 크기 안에 여백·본문·머리글·바닥글을 배치해 실제 PDF 페이지 수를 계산 결과에 맞췄다. 셀 표 뒤의 차트·이미지는 각각 별도 페이지로 출력해 페이지 번호/전체 수에 포함한다. 지정 공간을 넘는 머리글은 잘린 채 인쇄하지 않고 조정을 안내한다. 이는 차트를 시트의 원래 위치 그대로 인쇄하는 기능까지 구현한 것은 아니다.

검증: 인쇄 설정의 적용/취소/실행 취소·다시 실행/저장·재열기, 잘못된 참조·여백의 원자적 거부, 차트 전용 시트의 설정 왕복, 4페이지의 실제 PDF(첫/짝수 머리글·한글·반복 제목·페이지 번호·차트)와 기존 인쇄·페이지 나눔·프린터/머리글 그림 연결 보존 회귀를 통과했다. Poppler 렌더링으로 새 PDF 4페이지와 기존 표·차트·이미지 출력도 시각 확인했다.

근거: [Microsoft 머리글/바닥글 코드](https://learn.microsoft.com/en-us/office/vba/excel/concepts/workbooks-and-worksheets/formatting-and-vba-codes-for-headers-and-footers), [첫 페이지 머리글](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.firstheader), [짝수 페이지 머리글](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.evenheader).

1차 동적 배열 커밋 `8ba775f`는 [Windows·Mac·Ubuntu 핵심 회귀 검사](https://github.com/beaur8-prog/sheet-studio/actions/runs/41096830969)를 모두 통과했다. 후속 인쇄 설정의 플랫폼 검사와 남은 대용량/UI 검증은 계속 진행한다.
