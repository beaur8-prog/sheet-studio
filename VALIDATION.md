# 검증 기록

검증일: 2026-10-18 (Asia/Seoul).

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
