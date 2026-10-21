# Sheet Studio

**[Mac·Windows·Linux 다운로드](https://github.com/beaur8-prog/sheet-studio/releases/latest)** · [설치·사용자 가이드](docs/USER-GUIDE.md) · [버전별 수정 내역](docs/CHANGELOG.md)

현재 배포 버전은 **0.7.17**입니다. 문서 전체 100만 셀, 변경분 실행 취소, 셀 묶음 자동 복구와 파일·계산·필터·정렬의 작업 스레드 처리를 포함합니다. 기존 색 기준 필터, 차트 선택값·세로 막대형 개선, 숫자 연속 자동 채우기와 한글 조합 입력 수정도 포함합니다. [이번 릴리즈 안내](docs/releases/v0.7.17.md)를 확인하세요.

선택 편집·붙여넣기·채우기는 한 번에 10만 셀까지입니다. 큰 문서의 첫 전달과 XLSX 열기·저장에는 시간이 걸릴 수 있습니다. [대용량 문서의 범위와 사용법](docs/CAPACITY.md)을 확인하세요.

현재 소스의 **0.7.19 개발판**은 열기·저장 창에서 현재 파일 폴더를 우선 표시하고, 새 문서의 마지막 사용 폴더와 재시작 후 위치를 기억합니다. 0.7.18의 XLSX 기본·숨김 행 높이와 외부 표 높이 보존 수정도 포함합니다. 공개 0.7.17에는 아직 포함되지 않습니다. [행 높이 점검 결과](docs/ROW-HEIGHT.md)를 확인하세요.

설치된 Microsoft Excel의 네이티브 실행 파일과 한국어 UI를 분석하고, 익숙한 작업 흐름을 독립적으로 구현한 Windows·macOS 스프레드시트 앱입니다. Microsoft Excel의 원본 코드 복원이나 전체 기능의 동일성을 주장하지 않습니다.

## 실행

Windows·macOS·Linux와 Node.js 22.12 이상이 필요합니다.

```powershell
git clone https://github.com/beaur8-prog/sheet-studio.git
cd sheet-studio
npm ci
npm run runtime:install
npm run desktop
```

Windows 실행 파일은 `npm run package:win`으로 만듭니다. 결과는 `release/0.7.19/win-x64/Sheet-Studio-0.7.19-Windows-x64-Setup.exe`입니다. `npm run package:dir`은 같은 폴더 아래 `win-unpacked/`을 생성하며 이 폴더 전체가 필요합니다.

**Mac 빌드:** Mac에서 `SHEET_STUDIO_ADHOC=1 npm run package:mac -- --arch=arm64`(Apple Silicon)를 사용합니다. DMG/ZIP 경로, 기본 메뉴, 서명 및 GitHub Actions 사용법은 [macOS 빌드 안내](docs/MACOS.md)를 참조하세요. macOS arm64의 실제 실행과 시트 수명주기 회귀 테스트를 확인했습니다. 빌드 결과와 설치된 의존성은 Git 저장소에 포함하지 않습니다.

로컬 브라우저에서는 `npm run dev` 후 `http://127.0.0.1:5192`로 접속하세요.

앱은 오프라인에서 동작합니다. 작업 내용은 기기에 자동 저장됩니다. **파일 → XLSX 저장**으로 다른 앱에서 열 수 있는 파일을 만들고, **앱 문서 저장 (.sheet.json)**으로 앱의 차트·메모를 포함한 편집 상태를 보존할 수 있습니다.

**도움말 → 예제 통합 문서**에서 수식과 서식을 적용한 예제를 열 수 있습니다. 예제의 데이터는 이 작업에서 만든 가상 데이터입니다.

## 구현한 동작

- 한국어 리본, 이름 상자, 수식 입력줄, 시트 탭, 상태 표시줄, 어두운/밝은 도구 모음
- 셀 더블클릭/F2로 셀 안에서 직접 편집, 키보드 이동, 범위 선택, 복사·잘라내기·붙여넣기, 아래/오른쪽 채우기, Ctrl+Enter 범위 입력, 실행 취소·다시 실행
- 자동 합계 범위 점선·확정/취소, 수식 참조 색상, 마우스·방향키 참조 선택, 다른 시트 참조 지정, F4 상대/절대 참조 전환 ([사용법과 점검 결과](docs/FEATURE-REVIEW.md))
- 틀 고정, 복사 점선, 값·수식·서식 선택하여 붙여넣기·행/열 바꿈, 행·열 숨기기/자동 맞춤 ([워크시트 안내](docs/WORKSHEET.md))
- 가상화 격자: A1부터 XFD1048576까지 이동, 머리글 경계 드래그로 열 너비·행 높이 조절, 실시간 크기 안내·Esc 취소
- 여러 시트, 이름 변경, 복제, 삭제, 다른 시트의 셀 참조
- 행·열 삽입과 삭제, 연결된 수식과 범위 참조 조정
- 글꼴, 굵게·기울임꼴·밑줄, 색, 정렬, 숫자·통화·날짜·백분율 표시, 셀 병합, 메모
- 정렬, Excel 방식의 머리글 자동 필터·값 검색·체크박스·복합 조건·다시 적용 ([필터 안내](docs/FILTERS.md)), 중복 행 제거, 표 서식, 값·수식 변화에 반응하는 조건부 서식 규칙 관리
- 차트 10종·18개 조합: 여러 계열, 누적·100% 누적, 혼합 보조축, 범례·데이터 레이블·축 제목, 이동·8방향 크기 조절, 실제 XLSX 차트 입출력 ([차트 안내](docs/CHARTS.md))
- 피벗테이블: 여러 행·열·값 필드, 6종 집계, 다중 항목 필터, 총합계, 새로 고침 및 XLSX 피벗 내보내기 ([사용 안내](docs/PIVOT.md))
- 함수 검색·자동 완성·인수 안내, 배열 상수와 범위 연산을 사용하는 SUMPRODUCT, Excel 최신 함수 접두사 보존
- XLSX·CSV 읽기/쓰기, JSON 저장/복원, 인쇄/PDF
- 기존 XLSM 열기, 설치된 Excel의 VBA 실행 연결, 결과 가져오기, VBA 프로젝트를 유지하는 XLSM 저장

## 기존 VBA 매크로 사용

**Windows 데스크톱 Microsoft Excel이 설치되어 있어야 합니다.** XLSM을 연 다음 **개발 도구 → Excel에서 매크로 실행 → Excel 연결**을 누르세요. 열린 Excel 작업 사본에서 매크로를 실행하거나 앱에서 `Module1.MacroName`을 입력합니다. **결과 가져오기**로 셀 변경을 반영하고 **XLSM 저장**으로 저장합니다.

실행에는 Excel의 보안 정책이 그대로 적용됩니다. 차단된 매크로는 실행되지 않으며 앱은 Trust Center 설정이나 VBA 프로젝트 접근 권한을 변경하지 않습니다. 작업 사본의 폴더는 원본과 달라서 상대 경로를 쓰는 매크로에는 조정이 필요할 수 있습니다. 지원 범위와 검증 한계는 [VBA 안내](docs/VBA.md)를 참조하세요.

계산 엔진은 JavaScript `eval`을 사용하지 않습니다. 산술, 비교, 문자열, 절대·상대 참조, 범위, 오류 전파, 순환 참조 감지와 다음 함수 등을 지원합니다.

**379개 함수 이름**(이전 버전 별칭 30개 포함)을 지원합니다. `SUMIFS/COUNTIFS/AVERAGEIFS`, `INDEX/MATCH/XMATCH`, `INDIRECT/OFFSET`, `PMT/IRR/XIRR`, `STDEV/VAR`와 확률 분포, 공학·텍스트·날짜 함수가 포함됩니다. **수식 → 함수 삽입**에서 이름·한국어 설명·범주로 검색하고 구문과 예시를 확인하세요. 수식 입력 중에는 ↑↓로 함수를 선택하고 Tab으로 완성하며, 현재 인수 안내가 표시됩니다. [전체 함수 목록과 호환 범위](docs/FUNCTIONS.md)를 참조하세요.

## 범위와 차이

독립 VBA 런타임, Power Query 편집기, 공동 편집, 외부 연결, 동적 배열, 전체 열 수식(`A:A`), Excel 전체 함수 및 모든 계산 규칙은 구현하지 않았습니다. VBA는 설치된 Excel에서 실행합니다. 큰 문서의 호환성이나 Excel과의 모든 경계 조건 일치는 보장하지 않습니다. 0.7.17은 기록된 셀과 수식 범위를 최대 1,000,000셀, 선택 편집·붙여넣기·채우기를 한 번에 최대 100,000셀로 제한합니다. XLSM 원본과 실행 결과는 각각 최대 15 MB이며 Excel 연결의 셀 변경 계획은 200,000개 이하입니다. [기능별 제한](docs/CAPACITY.md)을 참조하세요.

차트 테두리의 조절점을 끌면 크기가 바뀝니다. 조절점에 키보드 포커스를 두고 방향키로 10px, Shift+방향키로 1px씩 조절할 수 있습니다. Esc로 드래그를 취소합니다.

차트는 첫 행을 계열 이름, 첫 열을 범주, 나머지 열을 값으로 사용합니다. 분산형은 첫 열이 X값이며 거품형은 X·Y·크기 열이 필요합니다. 새 차트는 XLSX·JSON에 저장됩니다. 기존 차트는 원본 XML과 연결 파트를 보존하며, 3D는 2D 미리 보기, 주식형·표면형·최신 ChartEx 유형은 원본 보존 안내로 표시합니다. 원본 차트의 전체 Excel 렌더링이나 모든 외부 연결을 지원하지는 않습니다. 조건부 서식은 지원하는 비교·텍스트·중복·상위/평균·수식 규칙을 다시 계산하며 XLSX/XLSM에 저장합니다. 색조·데이터 막대·아이콘 집합은 미지원입니다. 머리글 선택은 사용 영역을 기준으로 한 제한된 범위를 선택합니다. 자동 필터는 선택 범위 또는 현재 데이터 영역의 첫 행을 머리글로 사용합니다. 복사/채우기, 정렬, 잘라내기 및 인쇄의 모든 Excel 세부 동작을 재현하지는 않습니다.

지원하지 않는 XLSX 개체·서식은 다시 저장할 때 유지되지 않을 수 있습니다. 자세한 파일 형식 지원 범위는 [IO-NOTES.md](IO-NOTES.md)를 참조하세요.

## 분석과 검증

- [분석 요약](docs/ANALYSIS.md): 네이티브 분석과 화면 관찰의 범위 및 구현과의 구분
- [검증 기록](VALIDATION.md): 테스트 실행과 실행 파일 확인
- `node tests/ui-cell-editing.mjs --packaged`: 화면 갱신 중 더블클릭, 셀 안 커서·서식·입력·취소·병합·확대 상태 검사
- `node tests/ui-dimensions.mjs --packaged`: 열·행 경계 드래그, 안내선, 실행 취소, 자동 맞춤과 확대/고정 상태 검사
- `node tests/ui-worksheet.mjs --packaged`: 틀 고정·붙여넣기·범위 입력·숨기기·조건부 서식 UI 검사
- `node scripts/worksheet-native-smoke.mjs`: 설치된 Excel과 새 워크시트 기능 및 XLSM 저장 비교
- `node tests/ui-functions.mjs --packaged`: 함수 검색·범주·삽입·자동 완성·인수 안내·수식 저장 검사
- `node scripts/functions-native-smoke.mjs`: 설치된 Windows Excel과 계산 결과 및 XLSX 실제 열기 대조
- `npm test`: 계산, 파일 왕복, 차트 유형·참조·원본 보존·크기 및 플랫폼 설정 테스트
- `node tests/ui-excel-charts.mjs --packaged`: 10종 차트의 생성·설정·XLSX 저장/열기·변환 검사
- `node scripts/chart-native-smoke.mjs`: 설치된 Windows Excel에서 18개 새 차트 조합과 기존 차트 왕복 검사
- `node tests/ui-charts.mjs --desktop`: 실제 Electron에서 차트 조절·저장·복원 검사 (`--packaged`로 배포 앱 검사)
- `node tests/ui-smoke.mjs`: 로컬 개발 서버에 대한 실제 브라우저 상호작용 테스트 (설치된 Microsoft Edge 사용)

원본 Microsoft 실행 파일, 리소스, 아이콘 또는 디컴파일 코드는 앱과 이 저장소에 포함하지 않았습니다. 상세 원본 분석 산출물은 로컬에 보관하며 공개 저장소에는 독립 앱 코드와 분석 요약만 포함합니다.

수학·통계·공학 함수의 일부는 MIT 라이선스의 [Formula.js](https://github.com/formulajs/formulajs) 4.6.1을 사용하며, 셀 참조·공백·오류·날짜·조건 집계는 앱의 호환 계층에서 처리합니다. 배포 앱에 [오픈 소스 고지](public/THIRD-PARTY-NOTICES.txt)를 포함합니다.

Linux x64 DEB는 Linux에서 `npm run package:linux`로 빌드합니다. 설치·실행 방법은 [사용자 가이드](docs/USER-GUIDE.md)를 참고하세요.

## Studio 앱 간 표 복사 (0.7.10)

선택 범위를 복사해 Word Studio 1.2.11 또는 Slide Studio 1.6.1에 붙이면 편집 가능한 표가 됩니다. 반대 방향도 지원합니다. [사용법·교환 형식·상대 앱 구현·검증 안내](docs/STUDIO-CLIPBOARD.md)를 확인하세요.
