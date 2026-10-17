# Spreadsheet files and desktop app

This is an independent spreadsheet implementation. It does not contain Microsoft Excel code, binaries, icons or other extracted product assets.

## Integration

```js
import { readSpreadsheet, writeSpreadsheet } from './src/io.mjs';

const workbook = await readSpreadsheet(new Uint8Array(file.data), file.name);
const { data, mime, extension } = await writeSpreadsheet(workbook, 'xlsx');
await window.desktop.saveFile({ name: `${workbook.name}.${extension}`, data: Array.from(data) });
```

Both exports are asynchronous. Read supports `.xlsx`, `.xlsm`, `.csv` and `.tsv`; write supports `xlsx`, `csv` and `tsv`. JSON project persistence belongs to the UI. XLSX includes every worksheet; CSV/TSV includes the active worksheet with calculated values.

The sparse workbook model is:

```js
{
  name: '통합 문서', activeSheetId: 'sheet-id',
  sheets: [{
    id: 'sheet-id', name: 'Sheet1',
    cells: {
      A1: { raw: '12.5', style: { format: 'currency' } },
      A2: { raw: '=SUM(A1:B1)', style: { bold: true } },
      A3: { raw: "'00123", style: {} }
    },
    columnWidths: { 0: 120 }, rowHeights: { 0: 28 },
    merges: ['A5:C5'], charts: []
  }]
}
```

Column and row dimension keys are zero based, and dimensions are CSS pixels. Cell addresses and merge ranges use A1 notation. Dates are numeric Excel serials with `style.format: 'date'`. A leading apostrophe in `raw` marks a text literal. Optional `cachedValue` retains imported formula results. `style.numberFormat` retains the exact imported XLSX number format; it is reused only while its format category agrees with `style.format`.

Preserved styles: bold, italic, underline, font family/size, font color, solid background, horizontal/vertical alignment, wrapping, basic cell borders and number formats. Cell notes are preserved as plain text. Named formats are `general`, `number`, `currency`, `accounting`, `percent` and `date`. Fonts use point sizes as defined in XLSX.

Formula export calls `evaluateCell(workbook, sheet.id, address)` from `src/engine.mjs`, includes the resulting cached value and asks Excel to recalculate on load. An imported cached result is retained when the engine does not support that formula. Formula source strings are preserved. Export does not execute macros, external workbook links or embedded scripts.

CSV parsing supports UTF-8 BOM, UTF-16 BOM, RFC-style quoted fields, escaped quotes, CR/LF and embedded line breaks. CSV export uses UTF-8 BOM for Korean text. Leading-zero CSV identifiers are read as text. Plain-text values beginning with a formula marker are escaped on CSV export.

Limits: 100 MB input, 200,000 imported nonempty XLSX cells or CSV fields, and a 2 million-position CSV export rectangle. Embedded worksheet charts are imported and exported as native DrawingML. Supported chart types render in the app; advanced types retain their original chart parts with limited previews. Chart themes and direct chart dependencies are retained, but this is not full package preservation. Third-party pivot tables, standalone images, conditional formatting, validation, protection and other unsupported workbook objects may be lost. Cell theme colors may be approximated. See [chart support](docs/CHARTS.md). Password-encrypted files and legacy `.xls` are unsupported.

XLSM has a separate preservation path: the original package is retained and edited through installed Excel. `writeSpreadsheet` refuses native-source workbooks unless `allowMacroLoss: true` is explicitly supplied. See [VBA support](docs/VBA.md).

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

Windows artifacts appear under `release/0.4.0/win-x64/`: the portable `Sheet-Studio-0.4.0-Windows-x64.exe`, and the standalone directory `win-unpacked/Sheet Studio.exe`. The packaged app bundles the renderer and Electron bridge; ordinary spreadsheet editing does not require Node.js, npm, Excel or an internet connection. VBA execution and preserving XLSM edits require installed Windows desktop Excel.

## Dependency provenance and verification

XLSX support uses [ExcelJS](https://github.com/exceljs/exceljs), pinned at version 4.4.0. Its browser bundle is compiled locally by Vite. Registry metadata and official repository documentation were inspected for the dependency choice and formula-result contract. ExcelJS stores formulas/results but does not calculate formulas itself; the app's independent engine supplies calculations.

`tests/io.test.mjs` verifies multi-sheet XLSX round trips, formula source and cached results, number/text/boolean/date distinctions, formats, dimensions, merged ranges, shared-formula translation, Korean CSV text, embedded quotes/newlines, active-sheet CSV export and invalid-file errors.

## Pivot tables

Sheet Studio pivots include editable definitions in JSON and Custom XML. Nonempty XLSX reports also include native PivotTable/cache parts. Excel can refresh these; empty reports preserve cells and app settings only. Report-filter rows are reserved on export and cross-sheet formulas are adjusted on export/import. XLSM preservation writes app settings and result cells, without creating new native pivot objects. Native pivots created elsewhere are not imported into the app field editor. See [pivot support](docs/PIVOT.md).
