# Sheet Studio

설치된 Microsoft Excel의 네이티브 실행 파일과 한국어 UI를 분석하고, 익숙한 작업 흐름을 독립적으로 구현한 Windows 스프레드시트 앱입니다. Microsoft Excel의 원본 코드 복원이나 전체 기능의 동일성을 주장하지 않습니다.

## 실행

Windows와 Node.js 22.12 이상이 필요합니다.

```powershell
git clone https://github.com/beaur8-prog/sheet-studio.git
cd sheet-studio
npm ci
npm run desktop
```

설치 없는 Windows 실행 파일을 직접 만들려면 `npm run package`를 실행하세요. 결과는 `release/Sheet-Studio-0.1.0-Windows.exe`에 생성됩니다. `npm run package:dir`은 `release/win-unpacked/` 폴더를 생성하며 이 폴더 전체가 필요합니다. 빌드 결과와 설치된 의존성은 Git 저장소에 포함하지 않습니다.

로컬 브라우저에서는 `npm run dev` 후 `http://127.0.0.1:5192`로 접속하세요.

Windows 앱은 오프라인에서 동작합니다. 작업 내용은 기기에 자동 저장됩니다. **파일 → XLSX 저장**으로 다른 앱에서 열 수 있는 파일을 만들고, **앱 문서 저장 (.sheet.json)**으로 앱의 차트·메모를 포함한 편집 상태를 보존할 수 있습니다.

**도움말 → 예제 통합 문서**에서 수식과 서식을 적용한 예제를 열 수 있습니다. 예제의 데이터는 이 작업에서 만든 가상 데이터입니다.

## 구현한 동작

- 한국어 리본, 이름 상자, 수식 입력줄, 시트 탭, 상태 표시줄, 어두운/밝은 도구 모음
- 셀 편집, 키보드 이동, 범위 선택, 복사·잘라내기·붙여넣기, 아래로 채우기, 실행 취소·다시 실행
- 가상화 격자: A1부터 XFD1048576까지 이동, 열 너비·행 높이 조절
- 여러 시트, 이름 변경, 복제, 삭제, 다른 시트의 셀 참조
- 행·열 삽입과 삭제, 연결된 수식과 범위 참조 조정
- 글꼴, 굵게·기울임꼴·밑줄, 색, 정렬, 숫자·통화·날짜·백분율 표시, 셀 병합, 메모
- 정렬, 텍스트 필터, 중복 행 제거, 표 서식, 양수·음수 색상 표시
- 선택 범위의 막대·꺾은선 차트, 차트 이동·제목 변경
- XLSX·CSV 읽기/쓰기, JSON 저장/복원, 인쇄/PDF

계산 엔진은 JavaScript `eval`을 사용하지 않습니다. 산술, 비교, 문자열, 절대·상대 참조, 범위, 오류 전파, 순환 참조 감지와 다음 함수 등을 지원합니다.

`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `ABS`, `SQRT`, `POWER`, `MOD`, `INT`, `LEN`, `LEFT`, `RIGHT`, `MID`, `CONCAT`, `CONCATENATE`, `TRIM`, `UPPER`, `LOWER`, `SUMIF`, `COUNTIF`, `VLOOKUP`, `XLOOKUP`, `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY`.

## 범위와 차이

VBA, Power Query, 피벗 테이블, 공동 편집, 외부 연결, 동적 배열, 전체 열 수식(`A:A`), Excel 전체 함수 및 모든 계산 규칙은 구현하지 않았습니다. 큰 문서의 호환성이나 Excel과의 모든 경계 조건 일치는 보장하지 않습니다. 선택 편집/수식 범위는 최대 100,000셀, 저장 모델은 최대 200,000개의 기록된 셀로 제한합니다.

차트는 범위의 첫 열을 범주로, 두 번째 열을 값으로 사용합니다. 앱의 차트는 JSON에 보존되며 XLSX 차트로 내보내지 않습니다. 조건부 서식은 현재 값에 색을 적용합니다. 머리글 선택은 사용 영역을 기준으로 한 제한된 범위를 선택합니다. 첫 행은 단일 셀에서 실행하는 정렬·필터의 머리글로 사용합니다. 복사/채우기, 정렬, 잘라내기 및 인쇄의 모든 Excel 세부 동작을 재현하지는 않습니다.

지원하지 않는 XLSX 개체·서식은 다시 저장할 때 유지되지 않을 수 있습니다. 자세한 파일 형식 지원 범위는 [IO-NOTES.md](IO-NOTES.md)를 참조하세요.

## 분석과 검증

- [분석 요약](docs/ANALYSIS.md): 네이티브 분석과 화면 관찰의 범위 및 구현과의 구분
- [검증 기록](VALIDATION.md): 테스트 실행과 실행 파일 확인
- `npm test`: 계산 및 파일 왕복 테스트
- `node tests/ui-smoke.mjs`: 로컬 개발 서버에 대한 실제 브라우저 상호작용 테스트 (설치된 Microsoft Edge 사용)

원본 Microsoft 실행 파일, 리소스, 아이콘 또는 디컴파일 코드는 앱과 이 저장소에 포함하지 않았습니다. 상세 원본 분석 산출물은 로컬에 보관하며 공개 저장소에는 독립 앱 코드와 분석 요약만 포함합니다.
