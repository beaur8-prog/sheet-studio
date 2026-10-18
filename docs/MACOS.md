# macOS 빌드

Mac에서 Node.js 22.12 이상을 설치한 뒤 저장소 루트에서 실행합니다. Windows 패키지는 Windows에서, DMG/ZIP은 Mac에서 만듭니다.

```sh
npm ci
npm test
npm run desktop
```

로컬 테스트용 앱은 인증서 없이 ad-hoc 서명으로 패키징할 수 있습니다. 아래 두 명령 중 대상 Mac에 맞는 명령을 실행하세요.

```sh
# Apple Silicon (M 시리즈)
SHEET_STUDIO_ADHOC=1 npm run package:mac -- --arch=arm64

# Intel
SHEET_STUDIO_ADHOC=1 npm run package:mac -- --arch=x64
```

아키텍처를 생략하면 빌드하는 Mac의 아키텍처를 사용합니다. 다른 아키텍처를 요청하면 electron-builder가 해당 Electron 런타임을 내려받습니다. `--dir`을 추가하면 DMG/ZIP 없이 `.app`만 만듭니다.

결과:

- Apple Silicon: `release/0.7.2/mac-arm64/Sheet-Studio-0.7.2-macOS-arm64.dmg` 및 `.zip`
- Intel: `release/0.7.2/mac-x64/Sheet-Studio-0.7.2-macOS-x64.dmg` 및 `.zip`
- 각 폴더의 `artifact-manifest.json`: 결과 파일의 크기 및 SHA-256

배포 앱을 실행해서 확인하려면 빌드한 아키텍처와 같은 Mac에서 다음을 실행합니다.

```sh
node scripts/electron-smoke.cjs --packaged
node tests/ui-charts.mjs --packaged
node tests/ui-excel-charts.mjs --packaged
node tests/ui-formulas.mjs --packaged
node tests/ui-pivot.mjs --packaged
```

서로 다른 아키텍처의 실행 검사는 각각의 Mac에서 해야 합니다. `.app`은 내부 실행 파일만 복사하지 말고 번들 전체를 유지하세요.

## GitHub Actions

저장소의 **Actions → macOS packages → Run workflow**에서 두 아키텍처의 빌드와 패키지 UI 검사를 실행합니다. 성공하면 DMG, ZIP, 해시 목록과 테스트 스크린샷을 Artifacts에서 받을 수 있습니다. 실행 환경은 [GitHub의 macOS 호스트 목록](https://docs.github.com/en/actions/reference/runners/github-hosted-runners)을 기준으로 각각 `macos-15`(arm64), `macos-15-intel`(x64)을 사용합니다.

## 메뉴와 기능 범위

- Command+N/O/S, Command+Z/Shift+Z, Command+C/X/V/A/F와 macOS 기본 앱·파일·편집·창 메뉴를 제공합니다. 셀 선택에서는 통합 문서를 편집하고, 입력란에서는 텍스트를 편집합니다.
- 창을 모두 닫아도 Mac 앱은 유지되고, Dock 아이콘을 누르면 창을 다시 엽니다.
- 셀·수식·차트·피벗·XLSX/CSV/JSON 기능은 공통 JavaScript 코드입니다.
- **VBA 실행과 VBA를 보존하면서 변경 내용을 반영하는 XLSM 저장은 Windows Excel COM 연결 전용입니다.** Mac에서는 기존 XLSM을 읽고 원본 매크로 바이트를 포함한 `.sheet.json`으로 보관할 수 있지만, Mac용 Excel의 매크로 실행 연결은 구현하지 않았습니다. 매크로를 제외한 XLSX/CSV 사본은 앱의 확인 화면을 거쳐 내보낼 수 있습니다.

## 서명과 검증 범위

`SHEET_STUDIO_ADHOC=1`은 로컬 테스트용 서명이며 Apple의 Developer ID 서명·공증을 대신하지 않습니다. 이 모드에서는 hardened runtime과 공증을 사용하지 않습니다. 일반 사용자 배포 시 이 환경 변수를 제거하고 Developer ID 인증서 및 공증 자격 증명을 설정하세요. [electron-builder 26 macOS 설정](https://www.electron.build/v26/docs/mac/)과 [공증 안내](https://www.electron.build/docs/notarization/)를 따릅니다. 앱은 macOS 보안 설정을 변경하지 않습니다.

이번 수정은 Windows에서 단위 테스트, Electron UI 테스트와 Windows 패키징을 검증했습니다. macOS 실제 패키징·실행·서명·공증은 아직 실행하지 않았습니다. 위 워크플로 또는 실제 Mac에서 결과를 확인해야 합니다.
