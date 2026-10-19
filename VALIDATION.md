# 검증 기록

## 2026-10-19 · 0.7.11 차트 첫 숫자 열·행 누락 수정 (개발, 미배포)

- 숫자 표 `10,20 / 30,40 / 50,60`을 선택하면 기존 코드는 첫 열과 첫 행을 무조건 제목으로 빼서 두 번째 열의 `40,60`만 그렸습니다. 수정본은 두 계열 `[10,30,50]`, `[20,40,60]`을 모두 그립니다. 문자 항목명은 축 이름으로 유지하며 제목 없는 표와 숫자 한 열도 지원합니다.
- 첫 행/첫 열 사용을 자동 구분하고 차트 설정에서 직접 지정합니다. 숫자 연도·번호를 항목명으로 사용하거나 숫자 제목을 쓸 수 있고, 새 차트는 이 선택을 저장해 값 수정 후 임의로 재해석하지 않습니다. 가져온 Excel 원본 계열은 명시적으로 새 차트로 변환하기 전까지 유지합니다.
- 로컬 단위 시험 **127개 통과**, Windows 전용 1개 건너뜀. 첫 열·행, 문자 제목/항목, 한 열·한 행, 0과 수식, 명시적 제목/항목 설정, JSON/XLSX 왕복과 범위 제한을 검사했습니다.
- Mac arm64 소스와 0.7.11 패키지에서 실제 드래그/역방향 선택 후 삽입, 모든 열의 막대 수·값, 첫 열 수정, XLSX 저장·재열기, 항목 지정과 실행 취소·닫기·재시작을 통과했습니다. 기존 10개 차트 유형, 원본 보존·변환, 8개 크기 조절점·이동·확대·키보드·JSON 검사도 통과했습니다.
- [차트 회귀 CI 39232223131](https://github.com/beaur8-prog/sheet-studio/actions/runs/39232223131): Windows·Ubuntu 22.04·24.04 소스 Electron 실행 **모두 통과**. Windows 단위 시험 128개, Ubuntu 각각 127개 통과(Windows 전용 1개 건너뜀). 세 환경에서 선택/첫 열·행, 10개 유형/원본 보존, 1024×768 창의 크기 조절·이동·확대 검사를 통과했습니다. Windows Excel worker의 PowerShell 문법 검사도 통과했으며 설치 파일 또는 실제 Office 실행을 검증한 CI는 아닙니다. 검증 소스는 `60324eb`입니다.
- Windows Excel 저장 계획도 렌더러/XLSX와 같은 계열별 범위를 전달하도록 수정하고 계획 유효성을 검사했습니다. [Series.Values](https://learn.microsoft.com/en-us/office/vba/api/excel.series.values)와 [Series.XValues](https://learn.microsoft.com/en-us/office/vba/api/excel.series.xvalues)로 범위를 지정하고, 계열 이름·거품 크기는 문서의 R1C1 참조 방식을 따릅니다. **실제 Microsoft Excel/VBA 실행·왕복은 이번에 수행하지 않았습니다.**
- Mac 앱은 ad-hoc 서명 무결성을 확인했고 로컬 캐시에 준비했습니다. 공개 릴리즈·리토스는 0.7.10을 유지하며 이번에는 배포하지 않았습니다.

## 2026-10-19 · 0.7.11 Ubuntu 앱 연결·방향키 입력 수정 (개발, 미배포)

- 기존 0.7.10 DEB를 읽어 아이콘이 `1024x1024` 한 크기에만 설치되고, desktop 파일의 `StartupWMClass=Sheet Studio`와 파일명 `sheet-studio.desktop`이 다르며 런타임 `desktopName` 설정이 없음을 확인했습니다. 공통 아이콘 크기 9개(16·24·32·48·64·128·256·512·1024)를 설치하고, 패키지 메타데이터·시작 시 `app.setDesktopName()`·desktop 파일·실행 창의 식별자를 `sheet-studio`로 통일했습니다. [Electron의 Linux desktopName 설명](https://www.electronjs.org/docs/latest/api/app#appsetdesktopnamename-linux)을 따릅니다.
- 일반 셀 입력에서 방향키로 확정하지 않던 문제를 기존 0.7.10에서 재현했습니다. 수정본은 네 방향 확정·이동, 숫자·완성 수식·실행 취소, 숨김 행/열·병합 셀·시트 가장자리를 검사했습니다. F2·더블클릭·수식 입력줄의 글자 커서와 수식 참조 선택은 유지합니다. 입력 이벤트가 편집 모드를 임의로 바꾸지 않게 했고 IME 조합 중 확정을 막습니다.
- 로컬 단위 시험 **117개 통과**, Windows 전용 1개 건너뜀. 소스 앱의 입력 이동·직접 편집·수식·함수·parity 감사 UI와 Mac arm64 0.7.11 패키지의 입력 이동·수식 UI가 통과했습니다. Mac 패키지의 실행 코드 10개 파일이 소스와 일치하고 ad-hoc 서명 무결성도 확인했습니다.
- [Ubuntu 22.04·24.04 설치 검증 39227854352](https://github.com/beaur8-prog/sheet-studio/actions/runs/39227854352) **모두 통과**. DEB를 실제 설치하고 desktop 파일 구문·GIO 앱 목록 등록·GTK 아이콘 조회(16~512px)와 실제 X11 `WM_CLASS = "sheet-studio", "sheet-studio"`를 확인했습니다. 설치 앱에서 입력 이동·직접 편집·수식·함수 회귀 검사도 통과했습니다.
- 두 Ubuntu의 실제 Wayland 창에서 `set_app_id("sheet-studio")`를 확인했습니다. 검사 환경은 Xvfb 위에 입력 장치를 갖춘 Weston을 실행한 중첩 Wayland이며, 창을 명시적으로 표시해 식별자를 수집합니다. GNOME Shell의 실제 **Pin to Dash 메뉴를 수동 클릭한 시험이나 Wayland 전체 기능 검증은 아닙니다.** 사용자 Ubuntu에서의 고정 결과는 새 배포본 설치 후 확인해야 합니다.
- 테스트 환경만 바꾼 최종 검사는 앞서 설치한 DEB를 재사용했으며, 설치된 ASAR의 버전·desktopName·실행 코드가 현재 소스와 일치하는지 다시 검사했습니다. 검증 자료는 `.verification/0.7.11/`, GitHub Actions의 `linux-validation-ubuntu-*` 산출물에 보관합니다.
- 공개 릴리즈와 리토스는 **0.7.10을 유지**합니다. 0.7.11 배포는 요청받은 뒤 진행합니다. 화면 잠금이 확인되어 사용 중인 Mac 앱의 교체·재실행은 하지 않았고, 검증한 0.7.11 Mac 빌드는 로컬 캐시에 준비했습니다.

## 2026-10-19 · 0.7.10 공개 배포 검증

- [GitHub Actions 39189920681](https://github.com/beaur8-prog/sheet-studio/actions/runs/39189920681): Mac arm64·Windows x64·Ubuntu 22.04 x64 빌드와 배포 앱 검사 모두 통과. Windows 단위 시험 118개, Mac·Linux 각각 117개 통과(Windows 전용 1개 건너뜀).
- 세 플랫폼에서 기본 실행·시트 수명주기·색상·셀 서식·OS 클립보드·parity 감사 UI 회귀를 통과했습니다. Windows 설치 및 시작 메뉴 바로가기의 실제 실행 경로·앱 식별자와 설치 앱 실행, Linux DEB 설치 후 실행도 확인했습니다.
- 좁은 창에서 홈 버튼으로 이동하며 발생한 지연 스크롤이 새 메뉴를 닫던 문제를 수정했습니다. 1024×768 창에서 메뉴 유지와 실제 스크롤 후 닫힘을 검사했습니다. 클립보드 배포 검사는 실제 패키지 실행 파일을 선택하도록 수정했습니다.
- 공개 Mac ZIP을 다시 내려받아 체크섬·압축 CRC·ad-hoc 서명 무결성을 확인하고 추출한 앱으로 Sheet↔Word, Sheet↔Slide, Word↔Slide 여섯 방향을 재검사했습니다. 병합·색·테두리·표시값·줄바꿈·Word 실행 취소 모두 통과했습니다. 상대 앱은 Word 1.2.11·Slide 1.6.1의 검증된 Mac 패키지이며 사용자 문서는 열지 않았습니다.
- 공개 설치 파일(DMG·Windows Setup·Linux DEB)과 Mac ZIP을 SHA-256 검증했습니다. 릴리즈에 플랫폼별 매니페스트·런타임 출처와 통합 SHA256SUMS.txt를 제공합니다.
- Windows/Linux에서 세 앱을 동시에 실행하는 여섯 방향 시험, 실제 Windows Excel/VBA 실행과 Microsoft Excel 파일 왕복은 이번 배포에서 수행하지 않았습니다. Mac은 Developer ID 미서명·미공증이며 Windows도 Authenticode 미서명입니다.

## 2026-10-19 · 0.7.10 Studio 표 교환 (앞선 소스·로컬 검증)

- 단위 시험 **117개 통과**, Windows 전용 1개 건너뜀. 빌드와 실제 OS 클립보드 UI 시험 통과. 외부 표 HTML의 10 MB 수신 한도는 UTF-8 바이트로 검사하며, Sheet·Slide의 공유 모듈은 같은 파일로 유지합니다.
- 격리한 Sheet 0.7.10·Word 1.2.11·Slide 1.6.1에서 **Sheet↔Word, Sheet↔Slide, Word↔Slide 여섯 방향**을 소스 앱과 패키징한 Mac 앱으로 검사했습니다. Word의 표 전체 선택·복사, Slide의 홈 복사, 세 앱의 홈 붙여넣기를 사용합니다. HTML과 같은 선택 범위의 TSV, 병합·색·줄바꿈·표시값 및 Word의 한 번 실행 취소를 확인했습니다.
- 클립보드 UI 시험의 복구도 표준 텍스트·HTML·RTF·이미지를 한 번의 `clipboard.write()`로 되돌립니다. 단일 형식을 반복 써서 다른 형식을 지우던 복구 경로를 제거했습니다. 사용자 문서 대신 격리 프로필과 합성 파일을 사용합니다.
- 결과: `.package/studio-clipboard-unit-final.log`, `.package/studio-clipboard-native-final.log`, `.package/studio-clipboard-six-directions-source.log`, `.package/studio-clipboard-six-directions-packaged.log`, `test-results/suite-clipboard-*.png`. Windows·Linux와 실제 Microsoft Office 앱 간 교환은 재실행하지 않았습니다. 공개 릴리즈·리토스 갱신은 배포 요청 때 진행합니다.

검증일: 2026-10-18 (Asia/Seoul).

## 0.7.3 — 셀 안 커서와 직접 편집

- 기존 0.7.2에서 각 마우스 누름을 70ms 유지하고 클릭 사이에 100ms를 두면 셀 DOM 교체 때문에 더블클릭 편집이 시작되지 않음을 재현했습니다. 셀의 포커스 기본 동작 차단을 mousedown 단계로 옮기고, 두 번째 mousedown에서 편집을 시작해 렌더링 전에 커서를 활성화했습니다.
- 일반 텍스트는 클릭 위치의 글자 오프셋에 커서를 두며, 입력창에 실제 셀의 글꼴·크기·굵기·색·배경을 반영합니다. 병합 영역 크기와 확대 비율도 반영하고 텍스트 커서/캐럿 색을 명시했습니다.
- `node tests/ui-cell-editing.mjs` 및 최종 패키지의 `--packaged`: 화면 갱신을 포함한 느린 더블클릭, 클릭 위치의 캐럿과 글자 삽입, 셀/입력창 좌표·크기·서식, Enter/Tab 확정, Esc 취소, 실행 취소/다시 실행, F2 및 선택 후 새 내용 입력, 수식·병합 셀, 다른 셀 클릭 시 확정, 확대/틀 고정 상태와 재시작을 검사했습니다. 입력기 조합 중 Enter 무시 검사는 합성 이벤트이며 OS 한국어 입력기 전수 검증은 아닙니다. 실제 입력 커서를 숨기지 않은 스크린샷을 검토했습니다.
- 최종 **0.7.3 패키지**에서 `tests/ui-formulas.mjs`, `tests/ui-functions.mjs`, `tests/ui-worksheet.mjs`, `tests/ui-dimensions.mjs`를 각각 `--packaged`로 실행해 모두 통과했습니다. 렌더러 예외 0개.
- `node --check src/main.mjs`, `npm run package:win`, `git diff --check` 통과. 최종 ASAR의 dist 및 Electron 스크립트 10개 파일이 작업 파일과 바이트 단위로 일치하고 패키지 버전이 0.7.3임을 확인했습니다. Mac CI에 직접 편집 검사를 추가했으며 **실제 macOS 실행은 미검증**입니다.

0.7.3 휴대용 실행 파일: `release/0.7.3/win-x64/Sheet-Studio-0.7.3-Windows-x64.exe` (92,103,694 bytes). SHA-256: `9154866d9dab5cc3693ed51a31a89312f8c32a6acffdd2bd055baca56261175f`.

## 0.7.2 — 마우스로 열 너비·행 높이 조절

- 머리글 경계의 드래그 영역을 클리핑되는 자식 요소에서 별도 레이어로 옮겼습니다. 경계 양쪽 5px에서 크기 조절 커서와 강조 표시가 나타납니다. 드래그 중에는 격자가 다시 그려져도 안정적인 부모 요소에서 포인터를 유지하며 점선과 현재 크기(px)를 표시합니다.
- 크기를 실제로 변경한 드래그에만 실행 취소 기록을 하나 생성합니다. Esc 또는 창 포커스 상실 시 원래 크기를 복원하며, 기존 더블클릭 자동 맞춤과 편집 중인 셀 내용도 유지합니다.
- `node tests/ui-dimensions.mjs` 및 최종 패키지의 `--packaged`: 경계 양쪽 실제 hit-test와 커서, 열·행의 실시간 드래그/화면 크기/안내선, DOM 교체 후 포인터 유지, 실행 취소/다시 실행, 단순 클릭과 원위치 복귀 시 기록 생략, Esc 취소, 더블클릭 자동 맞춤, 150% 배율, 틀 고정/스크롤/숨김, 편집 내용 보존 및 재시작을 검사했습니다. 창 포커스 상실 분기는 `blur` 이벤트를 보내 검증했습니다. 열/행 조절 중 스크린샷도 검토했습니다.
- 최종 **0.7.2 패키지**의 `tests/ui-worksheet.mjs --packaged`와 `tests/ui-formulas.mjs --packaged` 회귀 검사 통과. 렌더러 예외 0개.
- `node --check src/main.mjs`, `node --check src/dimension-interactions.mjs`, `npm run package:win`, `git diff --check` 통과. 최종 ASAR의 dist 및 Electron 스크립트 10개 파일이 작업 파일과 바이트 단위로 일치하고 패키지 버전이 0.7.2임을 확인했습니다.
- Mac CI에 크기 조절 검사를 추가했습니다. 공통 UI 코드이며 **실제 macOS 실행은 미검증**입니다.

0.7.2 휴대용 실행 파일: `release/0.7.2/win-x64/Sheet-Studio-0.7.2-Windows-x64.exe` (92,104,485 bytes). SHA-256: `4614f156c5d4b075e61d9e707bc16f766a3aee9cbc3de09a0096c9e472b66f28`.

## 0.7.1 — 채우기·글자색 즉시 반영

- 홈의 두 색상 입력을 `input` 단계에서 적용하고 최종 `change` 이벤트도 지원합니다. 색상 창을 한 번 연 동안의 연속 선택은 하나의 실행 취소 기록으로 묶습니다. 선택 범위의 강조 배경이 실제 셀 채우기를 덮지 않게 수정했습니다.
- Windows Electron의 실제 기본 색상 창에서 A1:B2를 선택한 상태로 채우기와 글자색을 바꿨습니다. 창을 닫기 전 `input`만 발생한 시점에 네 셀의 저장된 색과 화면의 계산된 CSS 색이 모두 바뀐 것을 확인했습니다. 채우기 색을 연속 변경한 뒤 창을 닫고 실행 취소 한 번으로 이전 서식이 복원되는 화면도 확인했습니다.
- `node tests/ui-worksheet.mjs --packaged`: 틀 고정, 범위 입력/실행 취소, 붙여넣기, 숨기기, 조건부 서식 및 재시작 검사 통과.
- `node tests/ui-formulas.mjs --packaged`: 자동 합계, 수식 편집/확정/취소, 참조 선택, 실행 취소 및 저장 복원 검사 통과.
- `node --check src/main.mjs`, `npm run package:win`, `git diff --check` 통과. 최종 **0.7.1 ASAR**의 dist 및 Electron 스크립트 10개 파일이 작업 파일과 바이트 단위로 일치하고 패키지 버전이 0.7.1임을 확인했습니다. 공통 UI 코드 수정이며 실제 macOS 색상 창은 미검증입니다.

0.7.1 휴대용 실행 파일: `release/0.7.1/win-x64/Sheet-Studio-0.7.1-Windows-x64.exe` (92,103,907 bytes). SHA-256: `ec84dcef7f04b9be7977a1d50137565bc205b07d7846123717a1a62626e7beb6`.

## 0.7.0 — 워크시트 편집·보기 보완

- `npm test`: **96/96 통과**. 틀 고정의 좌표·숨김·선택 분할, 값/수식/서식·전치·반복·연산 붙여넣기, 조건부 규칙의 재계산·우선순위·중지·중복/상위/평균/와일드카드, JSON/XLSX 왕복과 네이티브 계획 검증을 포함합니다.
- `node scripts/worksheet-native-smoke.mjs`: 설치된 **Windows Excel 16.0 build 19127**과 XLSX 고정/숨김, 조건부 채우기·글자색 및 값 변경 후 재계산, 전치 붙여넣기, 상대 수식 범위 입력을 대조해 통과했습니다. 별도 합성 XLSM에 고정·숨김·규칙 적용 및 제거 후 다시 열기도 통과했습니다. 사용자 매크로는 실행하지 않았고 원본 파일이 바뀌지 않음을 확인했습니다.
- 실제 Excel 대조 중 dxf 채우기 색의 bgColor, ExcelJS에서 누락되는 stopIfTrue/빈 숨김 행, 네이티브 표현식 규칙의 `=` 접두사를 보완했습니다. 초기 oracle의 CutCopyMode enum 대입 오류는 불필요한 대입을 제거해 수정했습니다.
- `node scripts/filters-xlsm-smoke.mjs`: 기존 XLSM 필터 적용/해제, 날짜 필터, 원본 파일 불변 및 VBA 파트 유지 회귀 검사 통과.
- `npm run package:win`: 통과. 최종 **0.7.0 ASAR**의 dist와 Electron 브리지/작업 스크립트 **10개 파일**이 현재 파일과 바이트 단위로 일치하고 패키지 버전이 0.7.0임을 확인했습니다.
- 아래 검사는 최종 **0.7.0 패키지**에서 모두 통과했습니다.
  - `tests/ui-worksheet.mjs`: 고정/스크롤/셀 클릭/해제, Ctrl+Enter 상대 참조와 실행 취소/다시 실행, 값 전치 붙여넣기, 점선/Esc, 행·열 숨기기/취소와 키보드 이동, 조건부 규칙 생성·값 수정·재시작. 클립보드는 OS 권한과 구별하기 위해 메모리 텍스트 대역을 사용합니다. 실제 격자 및 새 규칙 창 스크린샷을 검토했습니다.
  - `tests/ui-formulas.mjs`, `tests/ui-functions.mjs`, `tests/ui-filters.mjs`: 자동 합계/참조 편집, 함수 검색/완성, 필터 회귀 검사.
  - `tests/ui-charts.mjs`, `tests/ui-excel-charts.mjs`, `tests/ui-pivot.mjs`: 차트 조절/이동/확대/스크롤, 차트 10종과 XLSX 왕복, 피벗 편집/저장 검사. 첫 차트 검사에서 스크롤 직후 DOM 교체와 좌표 측정이 겹쳤습니다. 검사 코드가 애니메이션 프레임 완료를 기다리게 고친 뒤 통과했습니다.
  - `scripts/electron-smoke.cjs`, `scripts/vba-ui-smoke.cjs`: 파일 IPC와 격리, XLSM 열기/Excel 연결/결과 가져오기/VBA 보존 저장. 렌더러 예외 0개.
- 정적 Vite preview의 `tests/ui-features.mjs`도 통과했습니다. Mac CI에 새 워크시트 UI 검사를 추가했으며 **실제 Mac 빌드/실행은 미검증**입니다. 전체 Excel 기능 전수 검사나 완전 호환성 확인은 아닙니다.

0.7.0 휴대용 실행 파일: `release/0.7.0/win-x64/Sheet-Studio-0.7.0-Windows-x64.exe` (92,103,505 bytes). SHA-256: `a9b709f858707c2e92a369c8aa64f3235bec1077bdd41e02d44bb04c7425c087`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.6.0 — Excel 방식 자동 필터

- `npm test`: **89/89 통과**. 범위 추정(중간 위치/머리글 한 행/빈 구간), 값·빈 셀·와일드카드·숫자/날짜 두 조건, 여러 열 AND, 상·하위/평균, SUBTOTAL, 정렬의 레코드/수식/서식 이동, JSON·XLSX 왕복 및 네이티브 필터 계획 검증을 포함합니다.
- Windows Excel **16.0 build 19127**의 별도 시험 문서에서 16개 조건의 표시 행/부분합을 수집해 `tests/fixtures/excel-filters.json`에 보관했습니다. 필터 상태에서 정렬할 때 숨겨진 레코드를 이동하지 않는 동작도 실제 Excel에서 수집했습니다. 사용자 문서와 기존 Excel 세션은 사용하지 않았습니다.
- `node scripts/filters-native-smoke.mjs`: 16개 조건의 XLSX를 실제 Excel에서 열고 **다시 적용한 후에도** 표시 행과 SUBTOTAL이 일치했습니다. 날짜 선택 목록/날짜 조건 2개도 별도로 열기/재적용을 확인했습니다. 최초 검사에서 빈 셀 제외 조건의 OOXML 값과 날짜 목록 인코딩 차이를 찾아 수정했습니다.
- `node scripts/filters-xlsm-smoke.mjs`: 기존 시험 XLSM에 숫자/값 복합 필터와 날짜 선택 필터를 추가해 저장·재열기·다시 적용·필터 제거를 검증했습니다. 원본 파일 바이트는 불변이고 VBA 프로젝트 파트가 결과에 남아 있습니다. Excel은 저장 시 VBA 내부 메타데이터를 다시 쓸 수 있으므로 VBA 파트의 바이트 동일성을 주장하지 않습니다. 이 검사에서는 매크로를 실행하지 않았습니다.
- `npm run package:win`: **0.6.0 패키지 생성 성공**. 최종 ASAR의 HTML, CSS/JavaScript, 라이선스 고지와 Electron 브리지 9개 파일이 작업 결과와 바이트 단위로 일치하고 패키지 버전이 0.6.0임을 확인했습니다.
- 다음 검사를 최종 Windows 패키지에서 통과했습니다:
  - `node tests/ui-filters.mjs --packaged`: 버튼 토글/범위/머리글 화살표, 검색·체크/취소/빈 결과·기존 선택 추가, 복합 조건/날짜/상위 항목, 여러 열 필터, 지우기/다시 적용, 필터 밖 행 유지, 방향키, 복사, 정렬, 시트 전환, 실행 취소/복원 및 재시작. 복사 검사만 OS 클립보드 권한과 분리한 메모리 클립보드를 사용합니다.
  - `node tests/ui-formulas.mjs --packaged`, `node tests/ui-functions.mjs --packaged`: 자동 합계·참조 표시·함수 도우미, 필터 상태에서 시트 참조 편집 후 복귀.
  - `node tests/ui-charts.mjs --packaged`, `node tests/ui-excel-charts.mjs --packaged`, `node tests/ui-pivot.mjs --packaged`: 차트 크기/저장/피벗 회귀 검사.
  - `node scripts/electron-smoke.cjs --packaged`, `node scripts/vba-ui-smoke.cjs --packaged`: 데스크톱 시작, XLSM 연결/결과 가져오기/보존 저장.
- `node tests/ui-features.mjs`는 최종 dist를 Vite preview로 제공해 통과했습니다. 처음 브라우저용 차트/피벗 검사를 서버 없이 실행한 시도는 연결 거부였으며, 위 패키지 검사로 다시 확인했습니다.
- macOS CI에 필터 UI 검사를 추가했습니다. **실제 Mac 빌드·실행은 미검증**입니다. 색/아이콘 필터, 날짜 계층/상대 날짜, 구조화된 표별 필터 등 남은 범위는 [필터 안내](docs/FILTERS.md)에 명시했습니다.

0.6.0 휴대용 실행 파일: `release/0.6.0/win-x64/Sheet-Studio-0.6.0-Windows-x64.exe` (92,096,231 bytes). SHA-256: `0eb53a2847ac3acdae4dd90c2392382431b7dfbc95291d7b98403adb4e3fd70c`. Windows 코드 서명은 되어 있지 않습니다.

## 0.5.0 — Excel 함수 라이브러리 확장

- 기존 37개에서 **379개 함수 이름**으로 확장했습니다. 이전 버전 별칭 30개를 포함합니다. Formula.js 4.6.1을 고정하고 Excel의 범위/배열/논리값/빈 셀/오류 처리, 참조 함수, 날짜 번호, 조건 집계, 일부 통계·재무·공학 계산을 앱에서 보정했습니다. 전체 목록과 제한은 [함수 안내](docs/FUNCTIONS.md)에 있습니다.
- `npm test`: **77/77 통과**. Excel oracle 예시 **371개**와 추가 경계 사례 **111개**를 회귀 검사합니다. 오류·타입·범위·배열 상수/연산, 미평가 분기, 참조 좌표·순환, 1900년 날짜·주말·공휴일, 확률 분포·큰 정수·제한된 반복, 함수 검색·구문 분석 및 XLSX 왕복을 확인했습니다. UTC / America/New_York / Asia/Seoul의 날짜 검사도 통과했습니다.
- `node scripts/functions-native-smoke.mjs`: Windows Excel **16.0, build 19127**에서 직접 수식을 입력한 결과와 앱 계산 371개를 대조했습니다. 앱이 생성한 XLSX의 375개 수식을 실제 Excel에서 다시 열어 계산한 결과도 직접 입력 결과와 모두 일치했습니다. 검사 파일은 읽기 전용으로 열었고 바이트 불변을 확인했습니다.
- 로컬 Excel이 인식하지 않는 TEXTAFTER/TEXTBEFORE/XLOOKUP/XMATCH 네 함수는 위 371개 Excel 계산 대조에 포함하지 않고 문서 기반 기대값으로 따로 검사합니다. RAND/RANDBETWEEN/NOW/TODAY는 반환 형식·범위를 검사합니다. 수치 대조 허용 오차는 1e-7 × max(1, |Excel 값|), 복소수는 실수부/허수부 기준입니다. 111개 경계 사례에서 TYPE의 레거시 암시적 교차 동작 한 건은 제외했으며 이유를 fixture에 기록했습니다.
- XLSX 저장 시 최신 함수 토큰에 필요한 `_xlfn.` 접두사를 추가했습니다. 최초 대조에서 이 접두사 누락으로 105개가 Excel에서 #NAME?을 반환했으며 수정 후 375개 파일 재계산 결과가 모두 일치했습니다. 문자열 안의 함수 이름은 바꾸지 않고 반복 저장 시 접두사를 중복하지 않습니다.
- `npm run package:win`: 통과. 최종 배포 ASAR의 HTML·CSS·JavaScript·MIT 고지가 최종 dist와 바이트 단위로 일치함을 확인했습니다. 아래 모든 UI 검사를 **0.5.0 패키지**에서 통과했습니다.
  - `node scripts/electron-smoke.cjs --packaged`
  - `node tests/ui-functions.mjs --packaged`: 검색·범주·빈 결과, 구문/예시, 삽입 전 값 불변, 실행 취소, 키보드/마우스 자동 완성, 인수 강조, 수식 입력줄, 중첩 삽입, Esc, IME 이벤트, 재계산 및 재시작. 실제 함수 라이브러리와 인수 도움말 화면도 확인했습니다.
  - `node tests/ui-formulas.mjs --packaged`: 자동 합계/참조 테두리와 기존 수식 편집 회귀 검사.
  - `node tests/ui-charts.mjs --packaged`, `node tests/ui-excel-charts.mjs --packaged`, `node tests/ui-pivot.mjs --packaged`: 차트 크기·10종 차트·피벗·XLSX 왕복 회귀 검사.
  - `node scripts/vba-ui-smoke.cjs --packaged`: XLSM 열기, Excel 연결, 결과 가져오기와 VBA 프로젝트 보존 저장. 사용자 매크로 전체의 실행 호환성 검사는 아닙니다.
- macOS 워크플로에도 새 함수 UI 검사를 추가했습니다. **실제 Mac 빌드·실행은 미검증**입니다. 스필/LET/LAMBDA/외부 연결/모든 표시 형식과 Excel의 전체 경계 규칙을 지원하는 것은 아닙니다.

0.5.0 휴대용 실행 파일: `release/0.5.0/win-x64/Sheet-Studio-0.5.0-Windows-x64.exe` (92,086,187 bytes). SHA-256: `468fa36475c405b1297bacd2bdc7d0450d7990d4d079d96342dea9d04d399f20`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.4.1 — 자동 합계 범위 표시와 수식 편집

- `npm test`: **54/54 통과**. 작성 중인 수식의 A1 참조 스캔, 문자열/함수 이름 제외, 시트 이름의 작은따옴표, F4의 네 가지 참조 형식, 괄호 보완, 위/왼쪽 자동 합계, 여러 열/행 합계, 빈 결과 셀 및 덮어쓰기 방지를 추가 검사했습니다.
- `node tests/ui-formulas.mjs` 및 `--packaged`: 실제 Windows Electron에서 자동 합계 점선의 위치·크기, 확정 전 셀 불변, 마우스 범위 교체, Enter/Tab 확정, Esc 취소, 실행 취소/복원, 가로 합계, 여러 열 합계, 기존 결과 보호, 참조별 색상, F4, 방향키와 Shift 범위, 다른 시트 참조, 필터 상태 복원, 150% 배율·스크롤, 잘못된 주소 거부 및 재시작을 검증했습니다. IME 조합 중 Enter를 무시하는 이벤트 검사도 통과했으며, OS 입력기 전체의 실기 검사는 아닙니다.
- `node scripts/electron-smoke.cjs --packaged`, `node tests/ui-charts.mjs --packaged`, `node tests/ui-excel-charts.mjs --packaged`, `node tests/ui-pivot.mjs --packaged`: 파일 IPC, 메뉴, 8방향 차트 크기 조절, 차트 10종, 피벗의 편집/저장/복원 회귀 검사 통과. 마지막 필터 복원 수정 후 수식 UI를 다시 검사했습니다.
- `node scripts/vba-ui-smoke.cjs --packaged`: 테스트 앱 시작 중 Chromium 네트워크 서비스 오류와 180초 시작 시간 초과가 한 차례 발생했습니다. 다른 UI 검사를 종료한 뒤 단독 재실행하여 Excel 연결·결과 가져오기·VBA를 유지한 XLSM 저장이 통과했습니다. 렌더러 예외 0개입니다.
- 브라우저 `ui-smoke.mjs`의 동일 검사항목을 진행 로그와 15초 제한을 추가한 로컬 사본으로 재실행해 통과했습니다. 첫 실행은 대기 상태가 계속되어 중단했습니다. 별도로 개발 서버가 패키징 산출물 변경에도 페이지를 다시 불러오는 현상을 관찰했습니다. 이후 개발 서버를 종료하고 최종 패키지의 수식/XLSM 검사를 수행했습니다. `node tests/ui-features.mjs`도 통과하여 상대 수식 붙여넣기, 병합, 입력 중 시트 변경, 정렬·필터, 찾기, JSON 저장/복원을 확인했습니다.
- `test-results/formula-autosum.png`에서 실제 점선 테두리와 선택된 수식 인수를 시각적으로 확인했습니다. 추가 점검의 범위 및 남은 기능은 [기능 점검 기록](docs/FEATURE-REVIEW.md)에 정리했습니다. 전체 Excel 기능의 전수 검사나 macOS 실제 실행 검증은 아닙니다.

0.4.1 휴대용 실행 파일: `release/0.4.1/win-x64/Sheet-Studio-0.4.1-Windows-x64.exe` (92,028,845 bytes). SHA-256: `11f90df270fe0a45b4644c20f8b0b855151a04bf7bd4e2867a974a2788883351`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.4.0 — Excel 차트 호환성 확대

- `npm test`: **50/50 통과**. 18개 차트 조합의 네이티브 XLSX/JSON 왕복, 여러 계열, 참조·캐시·시트 이름 변경, 제목·범례·레이블·축 제목, 반복 저장, 정의된 이름, 알 수 없는 차트/연결 이미지/ChartEx 보존 및 제한된 네이티브 변경 계획을 검증했습니다.
- `node scripts/chart-native-smoke.mjs`: 설치된 Windows Excel에서 **새 차트 18개 조합**을 열고 차트 유형·첫 계열 값·한국어 제목을 확인한 뒤 재저장했습니다. 혼합형의 계열별 막대/꺾은선 유형과 주축/보조축도 확인했습니다. **Excel에서 만든 기존 차트 13개**(3D·표면형 포함)를 가져와 제목·크기·셀 값 변경을 반영하고, 차트 개수와 원본 파일 불변을 검증했습니다.
- `node scripts/chart-xlsm-smoke.mjs`: Windows Excel 보존 경로로 기존 차트 제목·크기·범례 수정, 새 차트 18개 생성, 삭제, 원본 XLSM 불변 검사를 통과했습니다. 인터넷 출처 표시를 실제 Windows 대체 데이터 스트림에 기록하도록 고쳤으며, 출처를 알 수 없는 입력에도 표시를 유지하는 저장 경로를 확인했습니다. VBA 매크로 실행 성공 검사는 아닙니다.
- `npm run package:win`: 통과. `node scripts/electron-smoke.cjs --packaged`, `node tests/ui-charts.mjs --packaged`, `node tests/ui-excel-charts.mjs --packaged`, `node tests/ui-pivot.mjs --packaged`, `node scripts/vba-ui-smoke.cjs --packaged` 모두 통과했습니다. 실제 배포용 앱에서 차트 10종 생성, 설정, XLSX 저장/열기, 셀 변경, 원본 차트의 명시적인 변환, 지원하지 않는 차트의 변환, 실행 취소·재시작과 기존 8방향 크기 조절을 확인했습니다. 렌더러 예외 0개이며 혼합형·도넛형 화면을 시각적으로 확인했습니다.
- 최신 ChartEx는 합성 OOXML의 파트·관계 보존만 검사했습니다. 로컬 Excel에서 최신 차트 생성 호출이 거부되어 실제 최신 차트 왕복은 미검증입니다. 주식형 생성용 네이티브 테스트 자료도 성립하지 않아 원본 보존은 합성 자료로 검사했습니다. 전체 Excel 차트 기능과의 동일성을 뜻하지 않습니다.
- macOS 공통 코드와 패키지 UI 검사 워크플로를 갱신했습니다. 실제 Mac 빌드·실행은 **미검증**입니다. 지원 유형과 제한은 [차트 안내](docs/CHARTS.md)에 있습니다.

0.4.0 휴대용 실행 파일: `release/0.4.0/win-x64/Sheet-Studio-0.4.0-Windows-x64.exe` (92,026,342 bytes). SHA-256: `ff60a771502e285f8344cef1e740500dce2a8cca83e951c96a4fffdcc9c93c4b`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.3.1 — 차트 크기 조절 / macOS 빌드 경로

- `npm test`: **43/43 통과**. 차트 최소·최대 크기, 반대편 고정점, JSON 크기 보존 및 이전 문서 기본 크기, Mac 아키텍처 선택과 메뉴 동작 검사를 추가했습니다.
- `node tests/ui-charts.mjs --desktop` 및 `--packaged`: 실제 Windows Electron에서 8개 조절점, 모서리·변 드래그, 150% 배율, 스크롤, 1회 단위 실행 취소/다시 실행, Escape 취소, 방향키 조절, 제목 변경, 이동, 최소 크기, 재시작, JSON 저장/열기, 크기가 없는 이전 문서, 꺾은선 차트까지 통과했습니다. 결과 화면의 조절점과 차트 배치를 시각적으로 확인했습니다.
- `node scripts/electron-smoke.cjs --packaged`: 파일 다이얼로그, 제한된 preload, 파일 저장 검증, Mac 메뉴와 같은 IPC 경로를 통한 입력란 전체 선택·셀 붙여넣기·실행 취소/복원·모달 차단 검사 통과. 실제 macOS 기본 메뉴 실행 검사는 아닙니다.
- `node tests/ui-pivot.mjs --packaged`, `node scripts/vba-ui-smoke.cjs --packaged`: 피벗과 Windows Excel 연결·결과 가져오기·VBA 보존 XLSM 저장 회귀 검사 통과. 렌더러 예외 0개입니다.
- `npm run build`, `node scripts/package.cjs --win`: 통과. 버전/플랫폼/아키텍처별 출력 폴더를 사용하여 이전 버전의 실행 파일을 유지합니다.
- macOS arm64/x64 패키징 및 기본 메뉴 코드를 추가했습니다. 실제 Mac 빌드·실행·서명·공증은 **미검증**입니다. 수동 GitHub Actions 워크플로와 [Mac 빌드 안내](docs/MACOS.md)를 제공합니다. VBA 실행 연결은 Windows 전용입니다.

0.3.1 휴대용 실행 파일: `release/0.3.1/win-x64/Sheet-Studio-0.3.1-Windows-x64.exe` (91,987,511 bytes). SHA-256: `00c02e5acf1274712aab4a31633a6965d64c59bac4f4b50eb29b10b1f14fc1da`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.3.0 — 피벗테이블

- `npm test`: **39/39 통과**. 교차 집계, 가중 총평균, 여러 값/필드, 빈 셀·자료형 구분, 계산 오류, 범위 변경, JSON 설정, 실제 XLSX 파트/왕복, 보고서 필터 위치와 다른 시트의 참조 수식 보존을 검증했습니다.
- `node tests/ui-pivot.mjs`: 실제 화면에서 생성, 필드/집계 변경, 미리 보기, 결과 셀 보호, 원본 편집 뒤 새로 고침, 재시작, XLSX 저장/열기, 필터, 일반 범위 변환과 실행 취소 통과. 필드 편집 및 결과 화면의 스크린샷을 시각적으로 확인했습니다.
- 기존 `ui-smoke.mjs` 및 `ui-features.mjs`: 통과.
- `node scripts/pivot-native-smoke.mjs`: 설치된 Excel에서 9가지 XLSX 사례(합계, 평균, 다중 행·값, 다중 열·값, 행 필터, 보고서 필터, 값만 집계, 총합계 끄기, 필터 결과 0행)를 개방했습니다. 비어 있지 않은 8개는 실제 네이티브 PivotTable 존재와 새로 고침, 총합계를 검증했습니다. 0행 결과는 정적 셀과 앱 설정만 저장하는 예외를 확인했습니다. Excel이 다시 저장한 파일에서 앱 피벗 설정을 복원했습니다.
- 같은 네이티브 검사에서 앱 피벗이 포함된 XLSM을 Excel 보존 경로로 저장하고 결과 셀·Custom XML 설정·원본 파일 불변을 확인했습니다. 이는 새 XLSM 네이티브 피벗 생성이나 VBA 실행 성공을 뜻하지 않습니다.
- `node scripts/electron-smoke.cjs --packaged`, `node tests/ui-pivot.mjs --packaged`, `node scripts/vba-ui-smoke.cjs --packaged`: 배포용 Windows 앱에서 통과. 렌더러 예외 0개, 기존 Excel 연결과 XLSM 보존 저장도 통과했습니다.
- 확인 중 고친 문제: 필드 선택 상자의 너비, Excel 피벗의 숫자 캐시 저장 방식, 보고서 필터 공간, 수식 참조 위치, 모든 항목을 제외한 Excel 필터, 한국어 Excel의 General 숫자 서식 및 텍스트 COM 변환.
- 범위/호환성 한계는 [피벗 안내](docs/PIVOT.md)에 기록했습니다. 날짜 그룹·소계·슬라이서, 외부/OLAP 피벗, 기존 Excel 피벗 설정의 역변환은 검증 범위가 아닙니다.

0.3.0 휴대용 실행 파일: `release/Sheet-Studio-0.3.0-Windows.exe` (91,997,987 bytes). SHA-256: `66c87c7c9e2d6cc3482f2a595fd01ba94fcae4fbf97951972260b21a4c09d8bb`. Windows 실행 파일은 코드 서명하지 않았습니다.

## 0.2.0 — XLSM / VBA 연결

- `npm test`: **31/31 통과**. 기존 25개에 원본 XLSM 바이트/JSON 보존, 명시적인 매크로 제외 변환, 최소 변경 계획, 구조 변경 비교, 유효하지 않은 native 요청 거부, 사용자 함수 저장값 표시와 Windows 출처 표시 유지 검사를 추가했습니다.
- `npm run build` 및 `node scripts/electron-smoke.cjs`: 통과. 확장된 preload IPC, Excel 설치 탐지와 기존 파일 작업을 확인했습니다.
- `node scripts/vba-ui-smoke.cjs`: 통과. 실제 앱에서 XLSM 열기, 개발 도구 패널, Excel 연결, 연결 중 앱 단축키 차단, 결과 가져오기, A1 편집 후 보존 저장과 A2 수식 재계산을 확인했습니다. 렌더러 예외 0개입니다.
- 패키징한 `release/win-unpacked/Sheet Studio.exe`에서도 `electron-smoke.cjs --packaged` 및 `vba-ui-smoke.cjs --packaged`가 통과했습니다. ASAR 내부의 PowerShell 작업 파일을 꺼내 실행하는 배포 경로도 확인했습니다.
- `node scripts/native-smoke.mjs .verification/native-source.xlsm`: 저장·연결·결과 반환·중단, 새 시트/행 삽입/크기/병합, 원본 파일 불변, 기존 차트와 이름 정의 및 VBA 프로젝트 유지 검사가 통과했습니다. 테스트 보고서에서 `macroExecuted`는 **false**입니다.
- 실제 Excel 편집기에서 직접 작성한 `Smoke` VBA 모듈을 oletools 0.60.2로 저장 전후 추출해 **모듈 소스와 기존 시트 CodeName 동일**을 확인했습니다. Excel이 VBA 프로젝트의 내부 메타데이터를 다시 기록하므로 바이너리 해시는 달라집니다. 시작 도중 앱 종료 요청을 받는 경우에도 작업을 취소하는 검사에 통과했습니다.
- 실제 `Application.Run` 요청은 Excel의 “매크로를 실행할 수 없습니다” 오류로 거부되었습니다. 테스트는 이 오류와 결과 셀 불변을 기록했으며 보안 설정을 완화하지 않았습니다. **실제 VBA 실행 성공, UserForm, 외부 참조 및 모든 업무 매크로 호환성을 검증한 결과가 아닙니다.**

로컬 테스트 산출물과 작업 사본은 `.verification/`, `test-results/`, Windows 임시 폴더에 있으며 공개 저장소에 포함하지 않습니다. 아래는 0.1.0 최초 구현의 검증 기록입니다.

## 자동 검증

GitHub 업로드 전 추적 대상 소스 27개만 별도 임시 폴더에 복사하여 `npm ci --ignore-scripts`, `npm test`, `npm run build`를 실행했고 모두 통과했습니다. 이 검증은 부모 작업 폴더의 node_modules를 사용하지 않았습니다. Electron 다운로드·설치 스크립트는 이 소스 설치 검사에서 생략했으며, 데스크톱 실행 검증은 아래 별도 기록을 참조하세요.

`npm test`: **25/25 통과** (계산 엔진 19개, 파일 입출력 6개).

검증한 계산 동작은 산술 우선순위, 빈 값·텍스트·논리 값 처리, 지연 평가, 오류 전파, 순환 참조, 날짜 일련번호, 조회 함수, 절대·상대 참조, 시트 간 참조, 삽입·삭제 시 범위 축소, 표시 형식 및 유효하지 않은 문서 모델 처리입니다.

파일 테스트는 실제 XLSX를 생성·재개방하여 수식, 캐시 값, 데이터 타입, 시트, 스타일, 병합, 행 높이·열 너비, 메모와 공유 수식을 검증합니다. CSV의 BOM, 쉼표, 따옴표, 여러 줄, 유니코드 및 텍스트 식별자도 포함합니다.

`node tests/ui-smoke.mjs`: **통과**. 설치된 Microsoft Edge를 사용해 실제 브라우저에서 셀 입력, 종속 수식 갱신, 수식 입력줄, 서식 실행 취소, 아래로 수식 채우기, 다른 시트 참조, 한글 텍스트, 행 삽입, 새로 고침 후 상태 복원, 마지막 셀 이동, 예제, 차트, XLSX 다운로드를 실행했습니다. 렌더러 예외 0개입니다.

`node tests/ui-features.mjs`: **통과**. 클립보드의 상대 수식, 병합, 작성 중인 셀의 시트 전환, 수식을 포함하는 행 정렬, 필터, 찾기 및 JSON 저장·재개방을 검증했습니다. 렌더러 예외 0개입니다.

`npm run test:desktop`: **통과**. 실제 Electron 창의 화면, preload 브리지, 렌더러의 Node 접근 차단, 제목 변경, 파일 열기·저장 IPC, 취소 및 잘못된 바이트 검증을 확인했습니다. 테스트에서 파일 대화상자의 선택 경로는 제어했습니다.

## 화면 확인

`artifacts/sheet-studio.png`를 열어 한국어 글꼴, 리본, 수식 입력줄, 선택 테두리, 표와 차트의 배치를 시각적으로 확인했습니다. `test-results/desktop.png`에는 실제 데스크톱 창 렌더링 결과를 보존했습니다.

검증 중 발견한 Enter 중복 처리, 입력 중인 이름 상자 덮어쓰기, 다른 시트 범위의 끝점 참조 이동, 차트 범위 이동, 사용자 지정 숫자 형식 및 잘못된 JSON 모델 처리를 수정했습니다.

## 네이티브 분석 검증

원본 EXCEL.EXE SHA-256이 분석 전후 동일함을 확인했습니다. REA 세션을 종료했고 자체 분석 프로세스도 정리했습니다. 정적 JSON 산출물의 파싱과 두 함수의 Ghidra `DECOMPILE COMPLETED: true` 표시를 확인했습니다.

Excel 원본과 새 앱의 모든 기능을 자동 비교한 결과는 아닙니다. Ghidra 분석은 보고서에 명시한 두 함수와 PE/COM/리소스 경계로 한정됩니다. 수식 테스트의 통과는 구현한 규칙의 검증이며 Excel 계산 엔진 전체와의 동등성을 의미하지 않습니다.

## 실행 파일

Electron 런타임과 UI를 포함한 portable EXE와 unpacked 앱 생성이 완료됐습니다. 실제 사용한 런타임 버전 및 공식 배포 ZIP 해시 검증은 로컬 `release/runtime-provenance.json`, 최종 파일 해시는 로컬 `release/artifact-manifest.json`에 기록했습니다. 최초 빌드는 기존 설치된 Electron 42.11.15를 사용했으며, 공개 소스의 package.json도 검증한 42.11.15로 고정했습니다. 실행 파일과 분석 원시 자료는 코드 저장소에 포함하지 않습니다.

- Portable: `release/Sheet-Studio-0.1.0-Windows.exe`, 91,951,284 bytes, SHA-256 `c5bead423e7708152e7265a708ac94479fa58dd23d93896afe82cf7b037cb3cc`.
- Unpacked: `release/win-unpacked/Sheet Studio.exe`, 234,328,576 bytes, SHA-256 `cc9e91337b0168d0e3e6d50389890a9570e49a4b3c35919c9df7932a978fa8ea`.
- 패키징한 앱에서도 데스크톱 화면과 파일 IPC 검증을 통과했습니다. 완성 앱을 실행하고 화면을 확인했습니다. 별도 코드 서명 인증서는 적용하지 않았습니다.

ExcelJS 4.4.0 의존성 감사에서 UUID의 중간 등급 전이 의존성 경고가 있습니다. 현재 앱은 해당 UUID API를 직접 호출하지 않습니다. 이 결과를 의존성 취약점이 전혀 없다는 검증으로 해석하지 않습니다.
