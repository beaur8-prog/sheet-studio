# XLSX 구조 검사

검증 전용 .NET 8/Open XML SDK 3.5.1 도구다. 앱 런타임에는 포함하지 않는다. 합성 XLSX를 읽기 전용으로 열고 Office 2021 대상 스키마·SDK 의미 제약과 원본 해시를 검사한다. Microsoft Excel의 계산·화면·복구 경고 검증은 아니다.

저장소 루트에서 실행한다.

```sh
node tests/ooxml-sdk-fixtures.mjs
dotnet run --project tests/ooxml-sdk/Validator.csproj -- test-results/ooxml-sdk
```

결과는 `test-results/ooxml-sdk/sdk-validation.json`과 같은 폴더의 합성 자료에 남긴다. GitHub Actions의 `OOXML SDK structure validation`을 수동 실행할 수도 있다. 앱 피벗·필터 피벗·외부 평균 피벗·앱/원본 공존·표 확장·18개 차트 조합·빈 셀 정책·날짜 기준·서식을 각각 최초 저장과 재저장 두 번까지 검사한다.

API 근거: [SpreadsheetDocument.Open](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.spreadsheetdocument.open), [OpenXmlValidator.Validate](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator.validate). 의존성: [공식 NuGet 3.5.1](https://www.nuget.org/packages/DocumentFormat.OpenXml/3.5.1).
