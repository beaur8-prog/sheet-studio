# 기존 XLSM과 VBA

Sheet Studio의 Windows 데스크톱 앱은 설치된 데스크톱 Excel을 별도 COM 인스턴스로 연결합니다. 기존 `.xlsm`에 들어 있는 VBA 프로젝트를 유지하고 Excel의 실제 VBA 엔진에 실행을 요청합니다. Excel 없이 동작하는 VBA 인터프리터나 Excel 전체 복제품은 아닙니다.

0.7.29 개발에서는 원본 상태의 행·열 기본 서식 검증과 미지원 상세 서식 이동의 사전 차단을 보강했습니다. 아래 과거 Excel 실행 기록과 이번 저장 계획·UI 검사는 구분합니다. **이번 작업에서 실제 Windows Excel/VBA를 새로 실행하지 않았습니다.**

## 사용 순서

1. 데스크톱 앱의 **열기**로 로컬 `.xlsm` 파일을 선택합니다.
2. 필요한 셀을 편집하고 **개발 도구 → Excel에서 매크로 실행 → Excel 연결**을 선택합니다.
3. Excel 작업 사본에서 매크로를 실행합니다. 앱의 **이름으로 실행**은 인수가 없는 공개 매크로 `MacroName` 또는 `Module1.MacroName`을 호출합니다. **Excel 매크로 목록**은 Excel의 기본 실행 창을 엽니다.
4. 작업이 끝나면 **결과 가져오기**를 누릅니다. Excel 작업 사본을 저장해 Sheet Studio에 다시 읽습니다.
5. **XLSM 저장** 또는 Ctrl+S로 VBA를 유지한 파일을 저장합니다. XLSX/CSV 변환에는 매크로 제외 확인이 표시됩니다.

연결하는 순간 Excel이 허용한 Workbook_Open 등 열기 이벤트가 실행될 수 있습니다. 실행 중에는 앱의 통합 문서 편집을 막습니다. **중단·닫기**는 이 연결이 만든 Excel 프로세스만 종료하므로 아직 가져오지 않은 작업 사본의 변경은 앱에 반영되지 않습니다.

## 보존 방식

XLSM을 읽을 때 원본 OOXML ZIP 바이트와 편집 전 셀 모델을 보관합니다. ExcelJS로 XLSM을 새로 만들지 않습니다. Excel은 원본의 별도 작업 사본을 열고 변경된 셀 값, 수식, 기본 서식, 메모, 시트 이름, 행·열 작업, 크기와 병합을 반영합니다. 기존 VBA, 시트 CodeName, 이름 정의와 Excel 개체는 원래 패키지에서 유지됩니다. Excel이 저장 과정에서 내부 데이터를 다시 기록하므로 파일이나 VBA 바이너리의 해시가 동일하다는 뜻은 아닙니다.

원본 파일은 실행 중 덮어쓰지 않습니다. 파일 저장 대화상자에서 원본 경로를 직접 선택한 경우에는 사용자가 선택한 파일을 교체합니다. 기존 다른 Excel 창에 연결하거나 그 프로세스를 종료하지 않습니다.

앱의 JSON 문서에도 원본이 포함됩니다. 재시작 또는 JSON 복원 후에는 파일 출처를 다시 확인할 수 없으므로 네이티브 실행 시 인터넷 출처로 취급합니다. 원본 XLSM을 앱의 파일 열기 대화상자로 다시 열면 해당 파일의 실제 출처 정보를 읽습니다. 큰 XLSM은 브라우저 로컬 저장 용량을 넘을 수 있으므로 파일 저장을 사용하세요.

## 0.7.29에서 상세 서식 이동

현재 Windows 저장 계획은 기본 글꼴·색상·숫자 표시와 네 방향 테두리 등 지원 속성만 보냅니다. 대각선·회전/축소·특수 채우기·취소선·잠금, 문자별 서식·링크를 다른 셀로 옮기면서 누락될 수 있는 경우에는 변경 전에 안내하거나 저장 계획을 거부합니다. **선택하여 붙여넣기 → 값/수식**을 사용하거나 Excel에서 편집하고, 매크로가 필요 없는 사본은 별도 확인 후 XLSX로 저장하세요. 원래 상세 서식을 그대로 둔 값 편집과 지원 서식 변경은 사용할 수 있습니다.

표·이름·저장 배열·차트 전용 시트·상세 축/배치의 일부 구성 변경도 XLSM 연결 저장에서는 제한합니다. 앱의 XLSX 편집 지원 범위와 모든 XLSM 저장 동작이 같다는 뜻은 아닙니다.

## 보안과 호환성 한계

- 초기 편집과 XLSM 저장은 `AutomationSecurity = 3` 및 이벤트 비활성 상태에서 수행합니다. 명시적인 실행 연결은 `AutomationSecurity = 2`로 사용자의 Excel 정책을 따릅니다. 낮은 보안 수준(1), AccessVBOM 변경, 신뢰 위치 추가, 인터넷 출처 표시 제거는 사용하지 않습니다.
- 인터넷 출처(Mark of the Web)가 있는 원본의 표시는 작업 사본과 저장 결과에도 유지합니다. 정책·보호된 보기·권한 때문에 열기 또는 실행이 차단될 수 있습니다. 앱은 해당 오류를 표시하며 보안 결정을 대신하지 않습니다.
- Excel 4.0 XLM 매크로 시트는 지원하지 않습니다. VBA 비활성화 설정만으로 XLM 실행이 막히지 않는다는 Excel API 특성 때문에 사전에 거부합니다.
- 작업 사본은 임시 폴더에 있습니다. `ThisWorkbook.Path`, 상대 경로, 같은 폴더의 데이터 파일, 추가 기능, COM/ActiveX 구성 요소, 외부 연결 및 Office 비트 수에 의존하는 매크로는 별도 환경 확인이 필요합니다. 임시 작업 파일은 `%TEMP%/sheet-studio-vba-*`에 남을 수 있습니다.
- UserForm이나 다른 Excel UI가 필요한 코드는 Excel 창에서 상호작용합니다. 모든 UserForm, 외부 참조, 이벤트 또는 기존 업무용 매크로의 호환성을 검증한 것은 아닙니다.
- VBA 편집기, 중단점·디버거, 매크로 기록기, 인수 입력 UI는 Sheet Studio에 구현하지 않았습니다. Excel의 편집 기능을 사용하세요.
- VBA 사용자 정의 함수는 Sheet Studio의 JavaScript 계산 엔진에서 실행되지 않습니다. 지원하지 않는 함수의 기존 Excel 계산값은 셀 툴팁에 저장된 값이라고 표시합니다. 이 값은 입력 변경 후 오래된 값일 수 있으므로 Excel에서 재계산해야 합니다. 앱의 합계·정렬 등은 독립 계산 엔진을 사용합니다.
- XLSM 내 차트·피벗·컨트롤은 Sheet Studio 화면에 완전하게 표시되지 않습니다. Sheet Studio에서 새로 만든 차트나 시트 복제는 기존 Excel 개체·시트 코드를 복제하지 않습니다. 복잡한 구조 변경은 Excel에서 확인하세요.
- 열 너비는 CSS 픽셀과 Excel 글꼴 단위 사이에서 환산하므로 몇 픽셀 차이가 생길 수 있습니다.
- 입력과 결과 파일은 각각 15 MB 이하, 한 번의 변경 계획은 200,000셀 이하입니다. 네이티브 필터 범위/숨김 행에도 별도 200,000개 제한이 있습니다. 암호화 파일, `.xls`, `.xlsb`, `.xlam`은 현재 지원하지 않습니다. Windows PowerShell 및 Excel COM이 정책상 허용되어 있어야 합니다.

## 이전 버전의 Excel 실행 검증

계산/입출력/변경 계획의 자동 테스트와 실제 Excel에서의 셀 편집·수식 재계산·차트 및 이름 보존·원본 불변 검사를 수행했습니다. Excel 편집기로 직접 작성한 작은 VBA 모듈도 native 저장 전후에 추출해 소스가 동일함을 확인했습니다.

검증 PC에서 이름으로 매크로를 호출했을 때 Excel은 “매크로를 실행할 수 없습니다” 오류를 반환했습니다. 매크로 이름과 보안 정책 등에서 발생하는 Excel 오류이며, 테스트 중 정책을 완화하지 않았습니다. 따라서 이 환경에서 VBA 실행 성공이나 모든 기존 매크로의 호환성을 검증했다고 주장하지 않습니다.

`node scripts/native-smoke.mjs`는 Excel로 테스트 문서를 만들고 저장/연결을 검사합니다. 현재 정책에서 VBA 프로젝트 접근이 가능할 때만 테스트 모듈을 생성하며, 설정을 변경하지 않습니다. 별도로 작성한 **동일한 테스트 코드**의 파일은 `node scripts/native-smoke.mjs <fixture.xlsm>`으로 검사할 수 있습니다. 이 명령은 파일의 `Smoke` 매크로 실행을 요청하므로 임의의 업무 문서를 테스트 입력으로 사용하지 마세요. 매크로 실행 여부와 차단 메시지는 `.verification/native-report.json`에 분리하여 기록합니다.

```vb
Public Sub Smoke()
    ThisWorkbook.Worksheets("Data").Range("B2").Value = ThisWorkbook.Worksheets("Data").Range("A1").Value * 3
End Sub
```

## API 근거

- [Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity): 보안 모드와 XLM 예외.
- [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run): 이름으로 매크로 호출.
- [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open): 파일 열기, 링크 갱신 및 매크로 보안.
- [Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas): 열린 문서의 사본 저장.
- [XlBuiltInDialog](https://learn.microsoft.com/en-us/office/vba/api/excel.xlbuiltindialog): Excel 기본 매크로 대화상자 `xlDialogRun = 17`.
