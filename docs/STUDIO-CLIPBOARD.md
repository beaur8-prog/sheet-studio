# Studio 앱 사이에서 편집 가능한 표 복사하기

2026-10-19 · Sheet Studio 0.7.10 / Word Studio 1.2.11 / Slide Studio 1.6.1의 배포 기능입니다. 세 앱을 각각 새 버전으로 업데이트해야 모든 경로와 홈 버튼이 동작합니다.

## 사용하는 방법

1. Sheet Studio에서 표 전체의 셀 범위를 선택합니다. **홈 → 복사**, **우클릭 → 복사**, **⌘C(Mac) / Ctrl+C(Windows·Linux)** 중 하나를 사용합니다.
2. Word Studio에서 본문에 삽입할 위치를 클릭하고 **홈 → 붙여넣기** 또는 **⌘V / Ctrl+V**를 누릅니다. 문서의 편집 가능한 표가 됩니다. 기존 표의 셀 범위를 선택했다면 Word Studio의 기존 셀 붙여넣기 동작을 따릅니다.
3. Slide Studio에서는 슬라이드의 빈 공간을 클릭하고 붙여넣습니다. **편집 가능한 표 개체**가 생깁니다. 표 개체는 이동·크기 조절이 가능하며 셀 내용·색·병합을 수정할 수 있습니다. 기존 표의 셀을 선택한 상태에서는 선택 위치부터 내용·서식·병합을 붙이며, 표 일부만 걸치는 대상 병합 셀은 오류를 안내하고 변경하지 않습니다.
4. 실행 취소 한 번으로 붙여넣기를 되돌립니다. 원본 Sheet와 연결되는 실시간 링크는 없습니다.

반대 방향도 지원합니다. Word Studio에서는 **표 안 우클릭 → 표 편집 → 표 전체 선택** 후 복사합니다. Slide Studio에서는 표 개체 하나 또는 표의 셀 범위를 선택하고 복사합니다. Sheet Studio의 시작 셀, 다른 Word 문서, 다른 Slide 슬라이드에 붙일 수 있습니다. Slide에서 여러 개체를 한꺼번에 선택한 복사는 기존 Slide 내부 개체 복사를 따르며, 외부 앱으로 표를 보낼 때는 표 하나를 선택하세요.

브라우저 개발 모드에서 클립보드 읽기 권한을 주지 않았다면 **Ctrl+V / ⌘V**를 사용하세요. Slide의 브라우저 홈 버튼은 클립보드를 읽지 못할 때 마지막 앱 내부 복사를 붙이고 이를 안내합니다. 설치된 데스크톱 앱의 홈 버튼은 OS의 현재 클립보드로 일반 붙여넣기 이벤트를 발생시킵니다.

## 보존 범위와 제한

| 항목 | 처리 |
| --- | --- |
| 셀 내용·빈 셀·줄바꿈·한글 | 보존합니다. HTML 특수문자는 이스케이프합니다. |
| 병합 | `rowspan` / `colspan`으로 전달합니다. 겹치거나 한도를 벗어나는 병합은 거부합니다. |
| 글꼴·크기·굵게·기울임·밑줄·글자색·채우기 | 셀의 명시된 서식을 전달합니다. 수신 기기에 없는 글꼴은 대체될 수 있습니다. |
| 가로·세로 맞춤, 열 너비·행 높이 | HTML에 기록합니다. 앱의 페이지·슬라이드 배치 및 크기 조절에 따라 실제 크기는 달라질 수 있습니다. |
| 테두리 | 네 변의 위치·굵기·색을 전달합니다. Word는 지원하는 HTML 선 종류를 유지합니다. 현재 Slide의 표 모델은 점선·파선·이중선을 실선으로 변환합니다. Sheet로 되돌릴 때도 수신 모델의 선 종류·굵기 제한을 따릅니다. |
| 수식·숫자 표시 형식 | **현재 표시값**을 외부 앱에 보냅니다. 통화·백분율·날짜는 보이는 글자로 전달하며, 외부 표를 Sheet로 받으면 셀 텍스트로 저장합니다. `=...` 모양의 글자를 수식으로 실행하지 않습니다. Sheet 내부 복사는 원래 수식과 참조 이동을 계속 지원합니다. |
| 조건부 서식·숨김 | Sheet는 현재 계산된 색과 표시된 행·열을 복사합니다. 조건부 규칙 자체는 전달하지 않습니다. 숨김이 포함된 병합 범위의 복사는 기존 제한을 따릅니다. |
| 셀 내부의 서로 다른 글자 서식 | Word 내부에서는 유지하지만 Sheet·Slide로 표를 가져오면 셀의 첫 텍스트 구간 서식으로 단순화합니다. |
| 기타 | 메모·매크로·외부 연결·차트·그림·대각선 테두리·테마 정의는 이 표 교환에 포함되지 않습니다. 외부 앱으로 잘라내기 이동을 완료하는 기능은 제공하지 않습니다. |

수신하는 표 HTML은 UTF-8 기준 최대 10 MB입니다. 세 앱 모두 한글의 글자 수와 바이트 수를 구별해 검사합니다. Sheet와 Word는 최대 100,000셀입니다. Slide는 100행·50열, 셀당 2,000자·표 전체 50,000자 제한을 적용합니다. 제한 초과나 잘못된 병합 범위는 문서를 변경하기 전에 거부합니다. Word의 페이지 너비보다 넓은 표는 수신 후 열 너비를 줄이세요.

## 클립보드 계약 v1

OS 클립보드에 두 형식을 **동시에** 씁니다. 두 값은 같은 선택 범위의 같은 시점 자료입니다.

- `text/html`: 하나의 표, 행·셀, `rowspan` / `colspan`, 열 `<col width>` 및 셀 `colwidth` / `data-colwidth`, 명시적 인라인 CSS. `data-studio-table="1"`은 교환 형식 표식입니다. 표식이 없는 일반 Word HTML 표도 읽습니다.
- `text/plain`: 표시값의 TSV. 탭·줄바꿈·따옴표가 들어간 값은 따옴표로 감싸고 내부 따옴표를 두 번 씁니다. HTML을 지원하지 않는 수신 앱의 대체 자료입니다.

글꼴 크기는 HTML에서 pt, 열·행 크기와 테두리 굵기는 px입니다. 표의 텍스트는 수식이 아니며 원본 통합 문서의 수식, 파일 경로, VBA 또는 전체 문서를 메타데이터에 넣지 않습니다. HTML 문자열을 문서에 직접 삽입하지 않습니다.

Sheet의 `data-studio-copy`와 Slide의 `data-slide-studio-copy`는 같은 앱의 현재 내부 복사를 구분하는 일회성 식별자입니다. 원본 수식·개체를 유지하고 오래된 내부 복사가 외부 자료를 덮지 않게 사용합니다. 다른 앱은 이를 무시해야 합니다. 외부 앱에 Sheet 원본의 잘라내기를 완료시키는 신호를 보내서는 안 됩니다.

## 송신 쪽 구현

Sheet의 `src/table-exchange.mjs`가 셀 값·서식·병합을 교환용 표로 변환하고, `src/studio-table.mjs`의 `tableToClipboard()`가 HTML과 TSV를 생성합니다. `src/main.mjs`의 `copy()`가 계산된 표시값·조건부 색·가시 셀을 사용합니다.

Electron에서는 `electron/preload.cjs`의 제한된 `writeClipboard({text, html})`를 호출합니다. `electron/main.cjs`는 송신 창과 메인 프레임, 자료 타입과 크기를 확인한 후 `clipboard.write({text, html})`를 한 번 호출합니다. 웹에서는 `ClipboardItem`에 두 MIME 형식을 넣습니다. 일반 텍스트만 덮어 쓰면 표 서식은 전달되지 않습니다.

Word는 ProseMirror의 기본 셀/문서 선택 복사가 표 HTML을 생성합니다. 표만 선택해 복사할 때 `src/clipboard.ts`의 `tableClipboardText()`가 병합의 빈 칸과 인용된 줄바꿈을 포함한 TSV를 함께 만듭니다. 일반 글·혼합 선택의 텍스트 복사는 기존 직렬화를 따릅니다. Slide의 `src/table-exchange.ts`와 `copyStudioTable()`은 표 개체·선택 셀을 같은 HTML로 내보냅니다. 기존 Slide 내부 개체 식별자는 함께 유지합니다.

## 수신 쪽 구현

### Word Studio

홈과 표 메뉴는 `src/clipboard.ts`의 공통 붙여넣기를 사용합니다. 데스크톱에서는 기존 `wordDesktop.paste()` → `webContents.paste()` → 편집기의 정상 paste 이벤트 경로를 사용합니다. 브라우저에서는 같은 `ClipboardItem`의 HTML을 우선 읽고 `sanitizeImportedHtml()`를 거친 뒤 `editor.view.pasteHTML()`을 호출하며, HTML이 없으면 `pasteText()`를 씁니다. 비동기 읽기 중 문서·선택·포커스나 편집 상태가 바뀌면 적용하지 않습니다. 선택을 옮겼다가 같은 위치로 돌아온 경우도 취소합니다. 키보드 및 데스크톱 붙여넣기는 DOM paste 이벤트에서 크기·병합을 먼저 검사하고 기존 `transformPastedHTML`과 표 확장을 사용합니다. 잘못된 병합을 ProseMirror가 고치면서 기존 선택 본문을 바꾸지 않도록 파싱 전에 거부합니다.

셀 스타일은 `src/table-format.ts`, 병합과 열 너비는 Tiptap Table 확장이 처리합니다. `colwidth`와 `<col width>`가 필요하며 `data-colwidth`만 보내면 이 Tiptap 버전에서 열 너비가 사라질 수 있습니다.

### Slide Studio

`src/table-exchange.ts`의 `readStudioTable()`이 표 HTML을 읽고 기존 `TableData`에 매핑한 뒤 `validateTable()`로 검증합니다. `src/App.tsx`의 정상 paste 이벤트에서 내부 복사 식별자를 먼저 확인하고, 외부 HTML 표를 텍스트 상자보다 우선 처리합니다. 빈 슬라이드 영역에서는 `createTableElement()`로 삽입하고 기존 표 셀에서는 `pasteStudioTable()`로 범위에 적용합니다.

홈과 개체 우클릭 붙여넣기는 제한된 `pasteTableClipboard()` IPC → `webContents.paste()`를 사용합니다. 텍스트 편집용 `pasteAndMatchStyle()`를 쓰면 HTML이 제거되므로 표 붙여넣기에 재사용하지 않습니다. 발표 중·모달·텍스트 편집·잠금 상태를 기존 편집 규칙에 따라 보호합니다.

### Sheet Studio

`pasteContent()` → `spreadsheetTable()` → 기존 `pastePlan()` 경로입니다. 외부 HTML은 DOM에 부착하지 않은 `<template>`에서 읽고 허용된 셀 값·CSS 속성만 새 자료로 구성합니다. 스크립트와 능동 요소는 제외하고, 병합·크기·글자 수를 검사하며, 텍스트는 `scalarRaw()`로 수식처럼 실행되지 않게 저장합니다. 숨김 셀 보호·병합 검사·실행 취소는 기존 붙여넣기 계획을 사용합니다.

0.7.18 개발판은 외부 표의 행 높이도 전달합니다. 원본 행 높이를 대상 행의 최소 높이로 적용하고 더 큰 대상 행은 유지합니다. 숨김 행을 건너뛰는 매핑과 반복 선택 범위를 반영하며 셀·병합·높이가 함께 실행 취소·자동 복구됩니다. 값·수식·서식만 붙이기, 전치·연산 붙여넣기에서는 높이를 바꾸지 않습니다. XLSX 문서의 기본 행 높이도 화면과 외부 복사에 보존합니다. [행 높이 점검 결과](ROW-HEIGHT.md)를 참조하세요.

`src/studio-table.mjs`는 Sheet와 Slide 저장소에 같은 파일로 포함했습니다. 계약 변경 시 두 파일과 양쪽 테스트를 함께 갱신하세요. 각 저장소는 형제 저장소 없이 빌드·개별 테스트가 가능하며, 통합 검사는 아래처럼 세 저장소를 나란히 둡니다.

## 검증 방법

```text
work/
  sheet-studio/
  word-studio/
  slide-studio/
```

- Sheet: `npm test`, `npm run build`, `node tests/ui-studio-clipboard.mjs`
- Word: `npm run build`, `npx playwright test tests/studio-clipboard.spec.ts tests/table.spec.ts`
- Slide: `npm run build`, `npx playwright test tests/studio-clipboard.spec.ts tests/priority-safety.spec.ts tests/native-files.spec.ts tests/office-shell.spec.ts`
- 세 앱을 빌드한 후 Sheet 저장소에서 `node scripts/studio-clipboard-smoke.mjs`

통합 검사는 격리한 앱 프로필을 사용해 실제 OS 클립보드로 **Sheet↔Word, Sheet↔Slide, Word↔Slide 여섯 방향**을 확인합니다. Word의 실제 표 전체 선택·홈 복사, Slide의 홈 복사, 각 앱의 홈 붙여넣기를 사용합니다. 사용자 문서를 열지 않으며 종료 시 클립보드의 표준 텍스트·HTML·RTF·이미지를 한 번의 `clipboard.write()`로 복원합니다. 패키지 검사 시 `SHEET_STUDIO_EXECUTABLE`, `WORD_STUDIO_EXECUTABLE`, `SLIDE_STUDIO_EXECUTABLE`에 각 앱 실행 파일을 지정하면 같은 검사를 설치 번들로 수행합니다. 검증 스크린샷은 `sheet-studio/test-results/suite-clipboard-*.png`이며 Slide→Word 결과는 `suite-clipboard-slide-to-word.png`입니다. Mac에서 검증했으며 실제 Windows·Linux 앱 간 클립보드 실행 검증은 수행하지 않았습니다. 실제 Microsoft Office와의 호환성 검증을 뜻하지 않습니다.
