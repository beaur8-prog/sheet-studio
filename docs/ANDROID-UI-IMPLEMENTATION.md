# Android 전용 UI 구현·검증 기록

- 날짜: 2026-11-08
- 개발 버전: **0.7.46**
- 기준 계획: [Android UI 전환 계획](ANDROID-UI-ADAPTATION-PLAN.md)
- 범위: Android 웹 렌더러의 보기·편집 UI와 호스트 연결 지점. APK/AAB·서명·스토어 배포는 별도다.

## 적용 조건

[src/android-mode.mjs](../src/android-mode.mjs)가 Android 런타임을 판별한다. Electron의 `window.desktop`이 있으면 항상 기존 데스크톱 경로를 사용한다. Capacitor가 네이티브 플랫폼을 보고하면 그 값을 우선하고, 일반 브라우저/WebView에서는 Android 플랫폼·User-Agent를 사용한다. 창 크기·터치 지원만으로 Android UI를 켜지 않는다.

Android에서만 `body.sheet-android`와 모드 클래스를 붙이고 [전용 CSS](../src/android.css)를 적용한다. Windows·macOS·Linux·iPad에서 화면 폭을 줄여도 Android 메뉴·보기 모드·별도 복구 정책을 적용하지 않는다. 복구 DB·원본 스냅샷 DB·파일 식별자·세션 키도 Android용으로 분리했다.

## 구현한 사용 흐름

| 항목 | Android 동작 |
|---|---|
| 시작·열기 | 파일 요청 없는 새 실행은 빈 문서. 일반 문서는 입력 모드로 열되 시작만으로 키보드를 띄우지 않음. XLSM은 보기로 열고 명시적 읽기 전용 보기도 제공. 잘못된 파일·취소 시 열린 문서 유지 |
| 보기 화면 | 화면 모드와 문서 권한을 분리. 일반 문서는 전체화면에서도 셀 입력·수식·실행 취소 가능. 입력 중에는 입력줄만 표시하고 Enter의 포커스 유지. 전체 화면을 격자에 배정. 상단 연필·더보기는 격자 위 오버레이로 표시하고 2.8초 뒤 숨김. 탭/포인터/키 입력으로 다시 표시. 더보기에서 찾기·시트·확대·주소·차트·공유 사용. Esc/연필/Android Back으로 일반 화면 복원(XLSM의 읽기 전용 권한은 유지) |
| 읽기 | 긴 값·수식·메모와 저장된 계산값/피벗 제한 상세, 주소/범위/정의된 이름 이동, 검색, 여러 시트 이동, 차트 크게 보기와 계열별 처음 100개 자료 표 |
| 읽기 전용 불변 | 선택·시트 이동·임시 틀 고정·확대가 문서·dirty·실행 취소를 바꾸지 않음. 명령 가드와 문서 Proxy 가드로 수정 경로 차단. 검색 입력창의 입력은 허용 |
| 터치 | 브라우저 팬을 사용하며 터치 스와이프를 데스크톱 범위 드래그로 처리하지 않음. 큰 확대 버튼·너비 맞춤, 선택 손잡이와 주소 대안. 핸들 취소·캡처 상실·회전 시 선택 복원 |
| 편집 | 일반 문서는 셀 단일 탭으로 입력. 같은 입력창을 유지하며 Enter/다음 셀 이동, 123/문자 전환, 완료/취소. 색·서식·테두리의 적용/취소와 범위 유지, 실행 취소/다시 실행 |
| 필터·입력 목록 | 큰 ‘이 열 필터’·‘입력 목록’ 명령. 필터 창의 값/색/조건과 확인·취소를 큰 화면 조작 대상으로 제공. 보기 모드의 임시 행 표시(G2b)는 아직 노출하지 않음 |
| 수식·가로 화면 | 합계/평균/함수 버튼을 입력 전과 입력 중 모두 제공. 범위/결과 셀 미리 보기, 기존 셀 덮어쓰기·순환 참조 차단. 가로 입력줄 48 CSS px, 키보드 높이와 회전 분리 |
| 태블릿·DeX | 840×600 CSS px 이상 창에서 마우스 또는 외부 키보드가 있으면 리본·수식줄·시트 탭·상태줄 표시. 편집 화면 자동/터치/데스크톱 선택 저장. 작은 창·휴대폰 가로 화면은 터치 배치 유지. 넓은 화면의 첫 한글 입력 대기와 마우스 차트 이동/크기 조절 지원 |
| 모드 전환 | 전체화면 전환은 dirty·확정한 값·실행 취소·미확정 초안을 유지. Esc도 초안을 취소하지 않고 화면만 복원. 별도 읽기 전용 권한 전환은 기존 초안 확인/취소를 거침. 조합 중에는 전환을 유보 |
| 공유 | 미편집 열람은 보관한 원본 File/Blob을 공유. 편집한 세션은 검증된 내보내기 파일 공유를 기본으로 하며 ‘변경 전 원본’은 별도 선택 |
| 저장 | XLSX·CSV·앱 문서의 사본 저장. 브라우저 다운로드 요청을 OS 저장 완료로 표시하지 않음. 원본 덮어쓰기 경로는 제공하지 않음 |
| 복구 | 문서별 최대 20개 복구본과 원본 스냅샷 보관, 자동 퇴출 방지. 사용자가 명시적으로 복구본을 삭제할 수 있음. 현재 편집 중인 복구본 삭제는 제한 |
| 입력 초안 | 조합/입력 이벤트로 받은 초안을 문서 ID별로 별도 보관. 마지막 완료한 문서 복구와 구분해 다시 확인/취소할 수 있음 |
| 접근성 기반 | 셀 주소·값·읽기 전용/선택 상태와 논리 활성 셀 제공. 필수 조작에 주소·버튼 대안, 48 CSS px 기본 조작 영역, 변경된 선택 요약만 안내 |

XLSM은 Android에서 보기·원본 공유로 제한한다. Android에 Windows Excel/VBA 실행, 데스크톱 설치 도우미·업데이트, 인쇄/PDF 명령을 제공하지 않는다. 차트/행열의 작은 직접 드래그 손잡이도 비활성화하고 메뉴·사본 작업을 우선한다.

## 코드와 데이터 보호

- [보기 세션](../src/android-mode.mjs): 활성 시트·시트별 선택/스크롤/배율·임시 틀 고정을 원본 모델과 분리한다. 문서 ID·리비전·모드 세대로 오래된 내보내기 결과를 거부한다.
- [주 화면 연결](../src/main.mjs): Android 분기만 별도 열기·공유·사본 저장·복구 경로를 호출한다. 보기용 파일 교체는 편집 이력을 새로 만들지 않고 Worker에는 `persist:false`로 계산/조회 모델만 동기화한다.
- [작업 브리지](../src/workbook-tasks.mjs): 기존 리비전/계산 세대 검사를 유지한다. 비영구 동기화는 복구 flush를 예약하지 않는다.
- [변경 기록](../src/history.mjs): Android에서만 `canMutate`를 전달한다. 보기에서 모델 대입·삭제·문서 교체를 거부하고 undo도 실행하지 않는다. 데스크톱 호출에는 가드를 전달하지 않는다.
- [복구 저장소](../src/recovery-store.mjs): 문서별 보관 옵션은 Android DB에서만 활성화한다. 문서별 세대가 다른 파일 열기/저장으로 지워지지 않으며 한도 초과는 실패로 돌려준다. 원본 파일 삭제와 복구본 삭제는 연결하지 않는다.
- [원본 보관](../src/android-files.mjs): 원본 바이트를 담은 불변 File/Blob을 별도 IndexedDB에 보관한다. 공유 때 URI를 다시 읽어서 외부 변경 파일을 옛 원본으로 표시하는 문제를 피한다. 보관 실패는 안내하고 현재 문서와 원본을 유지한다.

초안 보관은 브라우저에 전달된 문자열 기준이다. OS 내부의 아직 전달되지 않은 조합 문자열이나 저장 완료 전 강제 종료까지 복구를 보장하지 않는다. 보기 재개는 reload 또는 호스트의 명시적 `resumeSession`에서만 시도하고, 일반 새 실행에는 이전 문서를 자동으로 열지 않는다.

## Android 네이티브 호스트 계약

현재 저장소에는 Android 네이티브 프로젝트가 없다. 아래는 패키징 단계에서 구현할 **연결 계약**이며, 실제 Android 파일 선택기/공유창을 이미 구현했다는 뜻이 아니다. 브라우저에서 지원되는 파일 input·Web Clipboard·Web Share/다운로드 대체 경로로 UI를 검사했다.

호스트는 Android WebView가 페이지를 시작하기 전에 `window.sheetStudioAndroid`를 주입할 수 있다. `protocolVersion: 1`, `platform: 'android'`가 필요하며 메서드는 Promise를 반환한다. 지원하지 않는 기능은 메서드를 생략한다.

| API | 계약 |
|---|---|
| `setFullscreen({enabled,hideStatusBar,hideNavigationBar,edgeToEdge,transientBarsBySwipe})` | 실행 시 위 상태줄·아래 탐색줄 숨김과 전체 창 사용 요청, 버튼의 표시/복원. 실제 적용 확인 후 `{enabled}` 응답. 렌더러는 네이티브 구현을 포함하지 않으며 미제공 시 버튼에서 웹 Fullscreen API 사용 |
| `closeApp()` | 종료 확인과 셀 입력 해결·복구 flush 후에만 호출. Android 호스트의 `close` 요청으로 Activity를 finish하며 실패 시 문서와 화면을 유지 |
| `openFile()` | 취소는 null. 성공은 `{name, mime, data}`. `data`는 읽기 완료한 ArrayBuffer/Uint8Array/Blob이며 URI 문자열만 전달하지 않음 |
| `saveCopy({name,mime,data})` | SAF의 새 문서로 저장. 취소 `{cancelled:true}`, 실제 제공자 완료 확인 시만 `{completed:true}`. 예외/불확실한 응답은 편집본 보존·실패 안내 |
| `shareFile({name,mime,data})` | 임시 파일의 읽기 권한·수명을 보장해 공유. 취소/오류는 원본·편집본을 변경하지 않음 |
| `writeClipboard({text,html})` / `readClipboard()` | 시스템 클립보드와 연결. 읽기는 `{text,html}`. 지원하지 않으면 브라우저 기능 사용 |
| `takeRequestedFile()` / `onOpenFile(callback)` | `{requestId,name,mime,data}`. 같은 요청 ID는 한 번만 소비. 작업·패널 중인 요청은 대기 후 문서 보호 절차를 거침 |
| `resumeSession` | 기존 Activity/작업의 재생성일 때만 true. 일반 새 실행은 false/생략. 새 파일 intent는 별도 요청으로 전달 |

시스템 IME의 뒤로 가기를 먼저 처리한 뒤, 호스트는 취소 가능한 `sheet-studio-android-back` 이벤트를 발행한다. `defaultPrevented`이면 앱이 메뉴/입력/문서를 처리한 것이므로 Activity 종료를 반복하지 않는다. 편집/보기 상태 자체를 바꾸거나 빈 문서로 교체하지 않는다. 입력·패널 처리를 마친 뒤에는 앱 종료 확인을 띄운다. 확인의 `닫기`는 복구 flush 후 `closeApp()`을 호출하고 `취소`는 현재 문서로 돌아간다. 브라우저에서는 앱 종료 API가 없다는 안내를 표시하며 종료 성공을 가장하지 않는다.

`sheet-studio-android-insets`는 매번 전체 스냅샷을 CSS px로 전달한다. 권장 detail은 `{systemBars:{top,right,bottom,left},displayCutout:{top,right,bottom,left}}`다. 보이는 시스템줄과 실제 카메라 cutout을 구분하며 누락된 변은 0으로 초기화한다. WebView가 이미 모든 여백을 소비했다면 `{consumedByWebView:true}`를 전달해 중복 패딩을 방지한다. 기존 평면 `{top,right,bottom,left}`는 시스템줄 여백으로 호환 처리하며 전체화면에서는 무시한다. `visualViewport`가 이미 반영한 IME 높이를 중복해서 빼지 않는다. `sheet-studio-android-input`의 `{hardwareKeyboard:boolean}`은 연결/해제 상태를 알려 준다. 실제 dp 대응·edge-to-edge·분할 창·IME 동작은 호스트와 함께 검증해야 한다.

## 전체화면과 상태줄

Android에만 manifest 링크와 `viewport-fit=cover`를 적용한다. `public/android.webmanifest`는 홈 화면 실행의 `fullscreen` 표시 방식을 요청한다. 일반 브라우저의 전체화면 버튼은 사용자 동작에서 `requestFullscreen({navigationUI:'hide'})`를 호출하고 다시 눌러 복귀한다. 네이티브 호스트는 `setFullscreen`이 있으면 실행 시 상태줄·탐색줄 숨김과 `edgeToEdge:true`를 요청한다. UI가 안테나 줄을 CSS로 가린 것처럼 처리하거나 기기 높이를 임의로 더하지 않는다. [Fullscreen 표준](https://fullscreen.spec.whatwg.org/), [Web App Manifest 표시 모드](https://www.w3.org/TR/appmanifest/#display-member)

APK 호스트의 실제 상태줄 숨김은 아래와 같은 네이티브 처리가 필요하다. 이 예시는 패키징 연결 지침이며 현재 저장소에서 빌드한 Android Activity가 아니다. WebView는 부모 전체 크기를 사용하고, 루트 padding/fitsSystemWindows로 같은 여백을 다시 예약하지 않아야 한다. 적용 후 호스트는 시스템줄과 cutout 스냅샷을 다시 전달한다. [Android 공식 immersive mode](https://developer.android.com/develop/ui/views/layout/immersive)

```kotlin
WindowCompat.enableEdgeToEdge(activity.window)
val controller = WindowCompat.getInsetsController(activity.window, activity.window.decorView)
controller.systemBarsBehavior =
    WindowInsetsControllerCompat.BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE
if (enabled) controller.hide(WindowInsetsCompat.Type.systemBars())
else controller.show(WindowInsetsCompat.Type.systemBars())
```

전체화면에서는 위·아래 시스템줄을 함께 숨기며 가장자리 스와이프로 일시 표시할 수 있다. IME를 숨기거나 키보드 높이까지 격자 공간으로 계산하지 않는다. 실제 OS 표시 방식·Gboard/삼성 입력기·상태줄과 cutout은 기기 검증 대상이다. 설치 앱이 이미 fullscreen 표시 방식이고 웹 API로 종료할 수 없는 경우에는 불필요한 전체화면 전환 버튼을 숨긴다.

## 검증과 남은 단계

자동 검사는 격리한 Chrome/Electron 프로필과 시험용 클립보드를 사용한다. 실제 사용자 문서·복구 DB·OS 클립보드는 사용하지 않는다.

- `npm test`: 기존 엔진·파일·서식·변경 기록 검사와 Android 판별·보기 불변·모드 전환 검사를 실행한다.
- `npm run test:android`: 프로덕션 빌드를 임시 로컬 서버에서 열고 Android 전화/태블릿의 User-Agent·viewport·터치를 모의한다. 원본 바이트 공유, 저장된 계산값 안내, 수정 차단, 검색, 시트 전환, 스와이프, 범위 취소, 한글 이벤트, 색 적용, undo/redo, 편집본 XLSX 재열기, 저장 실패/취소, 입력 초안 reload, 여러 문서 복구·용량 거부/삭제를 검사한다. 좁은 Windows/iPad 창에 전용 UI가 생기지 않는지도 확인한다.
- `node scripts/electron-smoke.cjs`와 기존 UI 검사는 실제 Electron의 버전·메뉴·열기/저장·편집 경로를 확인한다.

0.7.40의 2026-11-07 실행 결과(기준 검사):

| 검사 | 결과 |
|---|---|
| 프로덕션 빌드 | 통과, 0.7.40 manifest 확인 |
| 단위 검사 | 525개 중 524개 통과, 실패 0. Windows 전용 인터넷 출처/신뢰 토큰 검사 1개는 macOS에서 건너뜀 |
| Android UI 통합 | 휴대폰 412×915·회전/크기 변경 400×850·태블릿 1280×800 CSS px, Android User-Agent/터치 모의에서 통과. Windows/iPad User-Agent의 같은 좁은 창에는 Android UI가 없음 |
| 저장·보호 | 미편집 원본 공유 바이트 일치, 편집본 XLSX를 실제 파일 해석기로 다시 열어 값·색 확인. 저장 제공자 실패/취소/완료 응답은 모의하고 편집본 유지 확인 |
| 복구 | 다른 파일 열람 후 문서별 복구, 입력 초안 reload, 보관 한도 초과 시 거부·명시적 삭제 후 재저장 통과 |
| 대용량·실패 | 단순 숫자 100,000셀 JSON 열기·J10000 이동·잘못된 XLSX 열기 실패 후 기존 문서 유지 통과. desktop Chrome 156.0.8078.80의 한 차례 측정 약 1초. Android 성능/메모리 보장 수치가 아님 |
| 실제 데스크톱 앱 | macOS Electron smoke 통과: 버전·메뉴·파일 열기/저장·편집 IPC. 기존 UI parity 검사도 통과: 숨김 셀·여러 행·병합·잘라내기/붙여넣기·채우기·시트 복제·수식 참조·틀 고정·AutoSum 이력 |
| 문서·격리 | 버전 파일 일치, 문서 링크·28개 계획 시나리오·diff 공백 검사 통과. 시험 프로필과 클립보드 모의 사용 |

터치 검사 중 차트를 탭한 직후 합성 클릭이 새 배경을 눌러 창을 닫는 문제를 재현해 수정했다. 창을 여는 터치의 추가 클릭만 소비하며 키보드로 연 문맥 메뉴의 Enter는 유지한다. 범위 드래그 취소 검사는 창 크기 변경 이벤트가 처리된 뒤 원래 선택으로 돌아가는지 확인한다.

브라우저 모의 결과를 실제 Android Gboard·삼성 IME·TalkBack·SAF·저메모리 검증으로 집계하지 않는다. 실제 Android 기기/SDK 검증은 이 작업 환경에서 수행하지 않았고 아래 단계에 남긴다.

| 계획 단계 | 이번 구현 상태 |
|---|---|
| G0 | Android 연결 계약·브라우저 실행 경로 준비. SDK/기기 없는 환경이므로 실제 WebView/IME/SAF 적합성 확인은 남음 |
| G1 | Android 전용 모드·상태 분리·수정 가드·화면 구현 및 브라우저 검증 |
| G2a | 핵심 보기 UI·공유·터치/확대·읽기·복구 보호 구현. 실제 휴대폰 완료 판정은 보류 |
| G2b | 선택 사항인 임시 행 표시 미제공. 기존 문서 필터를 보기에서 수정하지 않음 |
| G3 | 기본 편집·서식·사본/공유·복구 UI 구현. 실제 파일 제공자 완료·실패 검증 필요 |
| G4 | 태블릿 리본·키보드 이벤트·입력 초안 보존 경로 구현. 실제 외부 키보드 첫 한글과 연결/분리 검증 필요 |
| G5·후속 | 실제 Android 접근성·메모리·시스템 제스처·장시간 문서 검사, APK/AAB 패키징·서명·배포는 남음 |

0.7.41의 연속 입력·집계·가로 화면 검증은 [0.7.41 기록](releases/v0.7.41.md), 0.7.42의 위아래 여백 제거·격자 확장 검증은 [0.7.42 기록](releases/v0.7.42.md)을 따른다.

0.7.42부터 앱 전체의 안전 영역 padding은 0이며 실제 가시 viewport를 채운다. 필요한 inset은 상단·하단 조작부 또는 격자의 cutout 가장자리에서만 처리한다. 전체화면 진입·복원은 inset과 크기를 다시 계산한다. 키보드가 페이지를 이동시킨 경우 `visualViewport.offsetTop/offsetLeft`와 크기를 함께 반영하며, pinch zoom 중에는 문서 레이아웃을 축소하지 않는다. 브라우저가 웹 콘텐츠 밖에 남긴 주소줄·OS 창 여백은 CSS로 회수할 수 없으므로 실제 실행 환경도 확인해야 한다. [Chrome edge-to-edge 안내](https://developer.chrome.com/docs/css-ui/edge-to-edge), [Android edge-to-edge 안내](https://developer.android.com/develop/ui/views/layout/edge-to-edge)


## 실제 Android 호스트 연결 (0.7.43)

별도 [Office Studio Android 수정 브랜치](https://github.com/beaur8-prog/office-studio/tree/codex/android-sheet-lifecycle-20261108)에 APK 호스트가 있다. 앞선 0.7.41/0.7.42의 렌더러 변경만으로 배포된 APK의 네이티브 창 설정과 포함된 웹 자산이 바뀌지는 않는다. 새 패키지는 Sheet 소스와 아래 호스트 변경을 함께 포함해야 한다.

Sheet에서만 WebView를 시스템줄·카메라 영역 전체로 확장하고 네이티브 부모의 카메라 padding을 제거했다. IME 높이만 네이티브 뷰포트에서 소비하며, 시스템줄·카메라 여백과 `displayCutoutRects`를 CSS px로 웹에 전달한다. 카메라 옆에 공간이 있으면 상단 메뉴를 좌우에 배치해 별도 빈 줄을 추가하지 않는다. 해당 메타데이터가 없거나 메뉴가 들어갈 수 없는 넓은 cutout이면 안전 여백을 유지한다.

입력 미리보기는 Android의 현재 편집 셀에 초안 문자열을 그리는 표시 처리다. HTML은 이스케이프하며 문서 값·수식 계산·실행 취소는 Enter/완료 때만 갱신된다. 입력 취소·유효성 오류·한글 조합·다음 셀에서도 기존 문서 보호 경로를 유지한다. [검증 기록](releases/v0.7.43.md)

0.7.43의 네이티브 수정은 [Android API 35 에뮬레이터 3개 검사](https://github.com/beaur8-prog/office-studio/actions/runs/42448953777)로 실행 검증했다. 실제 제조사 폰 검증과 정식 APK 배포는 별도다. [이번 결과](releases/v0.7.43.md)

## 0.7.44 모드와 네이티브 연결

여기서 데스크톱 편집은 **Android 태블릿·DeX의 넓은 창**만 뜻한다. Electron과 Windows/macOS/Linux에는 새 모드 버튼이나 화면 클래스를 생성하지 않는다. `src/viewer.css`의 모든 선택자도 Android 루트에 제한한다. `src/editor-layout.mjs`는 전체 창 크기로 판단하며 소프트 키보드가 열린 높이를 화면 전환 기준으로 삼지 않는다. 터치만 사용하는 넓은 태블릿은 자동 설정에서도 큰 버튼을 유지한다.

편집→뷰어 전환은 미확정 입력 확인/취소와 한글 조합 보호를 먼저 거친다. 모드 전환 자체는 문서·실행 취소·dirty를 바꾸지 않는다. 뷰어 진입 때 시스템 전체화면을 요청하며 복귀 때 이전 전체화면 상태를 복원한다. 패널이 열려 있으면 Esc는 패널부터 닫고 다음 Esc가 편집으로 돌아간다. XLSM은 기존대로 편집을 허용하지 않는다.

별도 Office Android 호스트는 Sheet에 한해 실제 하드웨어 Esc의 down/up/repeat를 소비하고 `OfficeStudioHost.escape()`로 웹에 전달한다. 키가 Android Back으로 바뀌어 Activity가 닫히는 것을 막는다. 구성 변경과 페이지 준비 때 `sheet-studio-android-input` 이벤트로 물리 키보드 연결 여부를 전달한다. 다른 앱의 키 처리와 배포본은 변경하지 않는다. 검사와 제한은 [0.7.44 기록](releases/v0.7.44.md)을 참조한다.

## 0.7.45 전체화면 입력·태블릿 동작

`viewer`는 Android UI 내부의 화면 상태다. `androidSession.mode`는 기존 읽기 전용/편집 권한만 결정한다. 전체화면 전환이 쓰기 권한을 바꾸지 않으며, XLSM·보호 셀·배열·피벗 보호도 그대로 적용한다. 사용자가 더보기 → 읽기 전용으로 전환을 선택하면 일반 파일도 수정 없이 열람할 수 있다. `sheetStudio.isViewer()`와 `getMode()`는 각각 화면과 권한을 보고한다.

넓은 화면의 입력 방식은 `android-desktop-input`, 실제 리본 표시는 `android-desktop-editor`로 분리한다. 전체화면에서도 마우스·외부 키보드 입력 준비를 유지한다. 셀 한 번 탭은 선택, F2/두 번 탭은 내용 편집, 바로 타이핑한 뒤 방향키는 확정 후 이동한다. 우클릭은 기존 셀 메뉴를 사용하고 행·열 크기 조절, 차트 이동, 채우기, Alt+= 자동 합계도 기존 편집 엔진을 사용한다. 작은 화면의 단일 탭 입력·큰 수식 도구는 유지한다. Android에 없는 데스크톱 파일 서비스·VBA·인쇄를 새로 제공하는 변경은 아니다.

네이티브 호스트는 Sheet의 하드웨어 Esc를 입력기 전 WebView와 Activity 양쪽에서 소비하고 `sheet-studio-android-escape`를 직접 보낸다. Back도 Sheet 전용 이벤트로 전달한다. 공통 브리지가 없거나 false를 반환했다고 Activity를 바로 닫지 않으며, 종료는 문서 보호를 거친 명시적 닫기 요청으로만 진행한다. 실제 검사 결과는 [0.7.45 기록](releases/v0.7.45.md)에 남긴다.

## 0.7.46 표시와 키 전달

Android 렌더러만 `displayResult`/`displayStyle`을 사용한다. 일반 값의 표시에는 평가기의 동일한 리터럴 변환을 쓰고 수식·배열/유출 셀·조건부 서식은 작업 스레드를 유지한다. 재계산 대기 중에는 같은 셀/수식의 직전 표시를 사용하며 `aria-busy`로 대기 상태를 구분한다. 표시 캐시는 최대 20,000개이며 문서를 바꾸거나 셀 컨테이너를 교체하면 정리한다. 유효하지 않은 원본 파일 계산값은 다시 표시하지 않는다. 실제 계산·복사·검증·저장은 이 표시 캐시를 사용하지 않는다. 처음 보는 수식은 계산 결과가 오기 전 잠깐 비어 있을 수 있다.

네이티브 호스트와 `setViewer`로 화면 상태를 동기화한다. Esc가 Android `KEYCODE_BACK`으로 오는 경우도 키 누름에서 화면 복원을 요청하고, 화면 복원이 끝난 후 도착하는 키 뗌을 소비한다. 전환 응답 대기 중의 Esc도 버리지 않고 복원 요청을 보관한다. Sheet의 정상 종료는 `closeApp` → `sheetClose` 경로로만 허용한다. 이전 공통 브리지의 `close` 요청은 Sheet를 종료하지 못하며 Word/Slide의 기존 close 처리는 유지한다. 문서 정보에는 Android 패키지 버전과 편집 엔진 버전을 함께 표시한다.
