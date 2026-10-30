# Spreadsheet files and desktop app — 0.7.29 development

This independent implementation does not contain Microsoft Excel binaries, resources or recovered source code. Development 0.7.29 is not a public release. Current behavior and verification limits are recorded in [the improvement log](docs/EXCEL-IMPROVEMENTS.md).

## Integration

```js
import {readSpreadsheet,writeSpreadsheet} from './src/io.mjs';
const workbook=await readSpreadsheet(new Uint8Array(file.data),file.name);
const {data,mime,extension}=await writeSpreadsheet(workbook,'xlsx');
await window.desktop.saveFile({name:workbook.name+'.'+extension,data:Array.from(data)});
```

Read accepts XLSX, XLSM, CSV and TSV. Write accepts XLSX, CSV and TSV; project JSON persistence belongs to the UI. XLSX includes ordinary worksheets and preserved chart sheets in their original order. CSV/TSV exports active-sheet calculated values. Legacy XLS and encrypted workbooks are unsupported.

## Model and value identity

The sparse model uses zero-based row/column dimension keys in CSS pixels and A1 cell addresses. Cells contain `raw` and optional `style`, note and preservation metadata. A leading apostrophe marks a text literal. Dates retain original numeric serials and `book.date1904`; code must pass the workbook date epoch into formatting/calculation. Changing a number format alone does not change the value type.

Detailed evaluation returns `{value,error}` so text such as `"#CALC!"` remains distinct from an error with that spelling. Worker calculations, clipboard values and file caches carry this distinction. Imported formula caches also carry their source formula and optional `cachedError` marker. Export recalculates supported formulas with the current session and asks Excel to recalculate on load. Unsupported formulas retain original caches only when their reference baseline is still valid. No macros, external connections or embedded scripts execute during ordinary export.

`xlsxStyles` stores deduplicated full Excel style models and editable baselines. `xlsxDefaultStyle` identifies Normal; sheet `xlsxStyleDefaults` stores styled row/column templates. Cell `xlsxStyle`/`xlsxNormal` distinguish an explicit original style or Normal override from inherited formatting. The exporter reapplies only changed supported properties over original details such as diagonal borders, text rotation, shrink-to-fit and protection. Rich text and hyperlinks retain original payloads when appropriate. Notes retain plain text. The app does not render every advanced font/fill/border/locale effect.

## XLSX preservation and editing

- Real tables include names, stable column IDs, calculated/totals formulas, styles and per-table filters. Structured references and defined workbook/sheet names are retained; supported references can be edited. Query-table connection graphs remain in the file but do not refresh in the app. See [tables](docs/TABLES.md).
- Stored array definitions retain source formula attributes and declared ranges. Supported arrays recompute followers and guard partial edits. Unknown modern functions remain unsupported; this is not a dynamic spill engine or full preservation of every dynamic-array metadata extension.
- Validation ranges/formulas, protection attributes/hashes, original differential/unsupported standard conditional-format XML and indexed colors are retained. Supported validation/protection works in the editor. Unsupported conditional rule types are preserved with limited previews/editing.
- Embedded charts retain original XML and related parts, supported series/name references, axes, themes and placement. Real chart sheets keep sheet order/state/names. Image/shape anchors and relationship graphs are retained; only basic raster images have previews. See [charts](docs/CHARTS.md).
- Page setup, margins, print options, header/footer XML, row/column breaks, ignored errors and linked printer/header-image parts remain in the file. Orientation edits persist. Original author/creation/document properties are retained where supported. App PDF printing is an A4 flow of cells followed by charts/images, not Excel's exact original print layout.
- This remains a selective OOXML editor. OLE/ActiveX, arbitrary extensions, external/OLAP pivot dependencies, all chart types and all unsupported workbook objects are not guaranteed to round-trip unchanged.

CSV parsing supports UTF-8/UTF-16 BOMs, quoted fields, escaped quotes, CR/LF and embedded line breaks. CSV output uses UTF-8 BOM for Korean and escapes formula-looking text. Leading-zero identifiers remain text. CSV cannot carry workbook styles or relationships.

Limits include 100 MB files, 512 MB expanded ZIP content, 1,000,000 recorded cells across all sheets and 100,000,000 raw characters. Selection edits/paste/fill/table conversion are bounded at 100,000 cells, printing at 20,000, and bulk structural operations at 10,000 rows/columns. XLSM original/result files remain 15 MB with 200,000 changed-cell plans. See [capacity](docs/CAPACITY.md).

## Desktop bridge

`window.desktop` exposes the following restricted operations:

- `openFile(): Promise<{name: string, data: number[]} | null>`
- `saveFile({name: string, data: number[]}): Promise<string | null>`
- `setTitle(title: string): void`
- `vbaProbe`, `vbaStart`, `vbaStatus`, `vbaCommand`, `vbaResult`, `vbaCancel` for the validated Excel session protocol

XLSM open/result payloads also carry an opaque `originToken`. The main process binds this token to the file hash and its original Zone.Identifier; a renderer-provided token alone cannot confer trust.

File paths come from native user dialogs. The renderer has context isolation, sandboxing, no Node integration and no unrestricted IPC API. New windows, navigation and permission requests are blocked. Saving writes a temporary sibling file and atomically replaces the selected destination.

## Build and run

From the cloned repository directory (Windows or macOS, Node.js 22.12 or newer):

```powershell
npm ci
npm test
npm run build
npm run desktop
npm run package
```

`npm run start` launches the current build, building first if none exists. `npm run dev` starts the browser development server at `http://127.0.0.1:5192`. `npm run package:dir` makes an unpacked app for the host platform. See [macOS builds](docs/MACOS.md) for DMG/ZIP and architecture selection.

`npm ci` installs the declared tools and Electron 42.11.15 locally. The repository does not require another application checkout. Runtime resolution prefers the local installation; parent and sibling lookups remain available as development fallbacks. Packaging helpers explicitly declare their JSZip and resedit dependencies.

The resolved Windows runtime is **Electron 42.11.15**, verified against the cached official release archive. Archive SHA-256: `f85fb7fb361508161bf0c3431e96518e34a9c10997048df21400a7da96658737`. The installed executable matches the archive's `electron.exe` at SHA-256 `8997a726173fb70b53bfb63b4f19865db759f9b247a66cd3b79b0b45c4eb626e`. Packaging writes `release/<version>/<platform>-<arch>/runtime-provenance.json` and changes only native product metadata/icon resources, checking that executable code-section hashes remain identical.

Windows artifacts appear under `release/<version>/win-x64/`: `Sheet-Studio-<version>-Windows-x64-Setup.exe` and the complete directory `win-unpacked/Sheet Studio.exe`. The packaged app bundles the renderer and Electron bridge; ordinary spreadsheet editing does not require Node.js, npm, Excel or an internet connection. VBA execution and preserving XLSM edits require installed Windows desktop Excel.

## Dependency provenance and verification

Formula support in 0.5.0 covers 379 names, including 30 compatibility aliases, using an explicit registry and Formula.js 4.6.1 behind Excel-specific adapters. See [the function catalogue and limitations](docs/FUNCTIONS.md). The renderer bundle and `THIRD-PARTY-NOTICES.txt` are included in both Windows and macOS packages. No network service or installed Excel is needed for these functions.

XLSX exports add the MS-XLSX `_xlfn.` prefix to future-function tokens, preserving strings and existing prefixes. This is required for Excel to recognize many newer functions after reopening the file. Unknown formulas and stored results retain the existing preservation behavior. `node scripts/functions-native-smoke.mjs` compares direct Excel calculations with actual reopening of an app-generated XLSX; the optional native check requires Windows Excel. The committed expected-value fixtures run on every platform through `npm test`.

XLSX support uses [ExcelJS](https://github.com/exceljs/exceljs), pinned at version 4.4.0. Its browser bundle is compiled locally by Vite. Registry metadata and official repository documentation were inspected for the dependency choice and formula-result contract. ExcelJS stores formulas/results but does not calculate formulas itself; the app's independent engine supplies calculations.

`tests/io.test.mjs` verifies multi-sheet XLSX round trips, formula source and cached results, number/text/boolean/date distinctions, formats, dimensions, merged ranges, shared-formula translation, Korean CSV text, embedded quotes/newlines, active-sheet CSV export and invalid-file errors.

## Pivot tables

App pivots include definitions in JSON/Custom XML and, for nonempty reports, native PivotTable/cache parts. Page-filter rows and cross-sheet formula offsets are handled during export/import. XLSM plans do not create new native pivot objects. Native pivots created elsewhere are not imported into the app field editor. See [pivot support](docs/PIVOT.md).

## Filters, worksheet and native preservation

Worksheet and table filters retain supported values, custom/number/text/date/color conditions, sort state and hidden rows. Filtered rows affect SUBTOTAL and chart visibility. Freeze panes, manual hidden dimensions and saved sizes round-trip. Supported conditional rules recalculate without permanently painting cells. See [filters](docs/FILTERS.md) and [worksheet behavior](docs/WORKSHEET.md).

XLSM imports retain original macro bytes and use the Windows Excel bridge for VBA-preserving saves. The bridge handles only its validated change-plan features, so table definitions, names, array edits, full format reset/painting and new axis/placement edits are restricted when necessary. This development run did not newly verify actual Microsoft Excel/VBA execution. Cross-platform Electron tests and OOXML/ExcelJS checks are recorded separately from earlier real Excel oracle evidence.
