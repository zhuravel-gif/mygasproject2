# mygasproject2 — Project Reference

## Stack
Google Apps Script (GAS) + HTML/CSS/JS webapp, deployed via clasp.

## Branch
`claude/sheetjs-import-refactor` (pushed to origin, PR-ready against `main`)

## Files
| File | Role |
|---|---|
| `Код.js` | Entry point: `doGet()`, `include()`, `onOpen()`, `openWebApp()` |
| `DataService.js` | Server-side data R/W for Google Sheets |
| `CalcEngine.js` | Cost calculation engine (server-side) |
| `WebApp.html` | Main HTML shell, tab layout, import modal |
| `Scripts.html` | All client-side JS (included via GAS `include()`) |
| `Styles.html` | All CSS |
| `SheetJsLib.html` | SheetJS xlsx 0.18.5 minified (~1.3 MB, CDN copy) |

## Data model — sheet `1cData` (18 columns)

| Col | Index | Key | Label |
|---|---|---|---|
| A | 0 | NAME | Номенклатура.Наименование |
| B | 1 | ART | Артикул |
| C | 2 | ART_WB | Артикул ВБ |
| D | 3 | CAT | Категория товаров |
| E | 4 | VOL | Объем тары |
| F | 5 | GRP2 | Товарная группа 2 |
| G | 6 | GRP3 | Товарная группа 3 |
| H | 7 | RAW | Основное сырье |
| I | 8 | FLAKON | Тара (флакон) |
| J | 9 | SET_QTY | Количество лаков в наборе |
| K | 10 | ART_MP | Артикул МП |
| L | 11 | IS_SET | Это набор |
| M | 12 | WEIGHT | Вес |
| N | 13 | GRP1 | Товарная группа 1 |
| O | 14 | COST_1C | Себестоимость 1С |
| P | 15 | PRICE | Цена поставщика |
| Q | 16 | NDS | НДС |
| R | 17 | TAX | Пошлина |

## Other sheets
- `Флаконы` — flakon list with price/nds/delivery/label
- `basis` — raw material basis
- `Результаты` — saved calculation results
- `__meta` — hidden, stores JSON: import mappings + rate defaults

## CalcEngine.js — COL constants
```js
var COL = { NAME:0, ART:1, ART_WB:2, CAT:3, VOL:4, GRP2:5, GRP3:6,
            RAW:7, FLAKON:8, SET_QTY:9, ART_MP:10, IS_SET:11,
            WEIGHT:12, GRP1:13, COST_1C:14, PRICE:15, NDS:16, TAX:17 };
```
Public functions: `calculateAll(params)`, `calculateManual(input)`,
`calculateByIndex(index,params)`, `getVerification(index,params)`, `determineType(row)`.

## DataService.js — CFG + public API

```js
CFG = { DATA:'1cData', FLAKONS:'Флаконы', BASIS:'basis',
        RESULTS:'Результаты', META:'__meta',
        BASE_COLS:14, COST_1C_COL:14, PRICE_COL:15, NDS_COL:16, TAX_COL:17, TOTAL_COLS:18 }
```

| Function | Description |
|---|---|
| `getData()` | Returns `{headers, rows}` from 1cData |
| `getParams()` / `saveParams(p)` | Load/save calc params (usd/rmb/log/com) |
| `getFlakonList()` / `saveFlakonData(fl)` | Flakon CRUD |
| `saveResults(results, params)` | Write calc results to Результаты sheet |
| `importNomenclature(data, nds, tax)` | **Wrapper** → `importBaseNomenclature` — replaces entire 1cData |
| `importCost1C(mappedData)` | **Wrapper** → `importCurrentCost` — fills col O by name match |
| `importSupplierPrice(mappedData)` | **Wrapper** → `importSupplierPriceData` — fills col P by name match |
| `updateNdsTax(updates[])` | Batch update NDS/TAX by row index: `[{row,nds,tax}]` |
| `updateAllNdsTax(nds, tax)` | Apply same NDS/TAX to all rows |
| `getImportSettings()` | Returns saved mappings + rate defaults from `__meta` |

`mappedData` format: `[{ name: string, value: number }]`
Matching: by name (normalized), fallback by article.

## WebApp.html — UI structure

**Tabs:** Импорт данных · Параметры · Флаконы · Себестоимость · Калькулятор · Проверка

**Tab "Импорт данных":**
- Buttons: `startImport('nomenclature')`, `startImport('cost')`, `startImport('supplier')`, `loadExportData()`, `exportCSV()`
- Rate toolbar: `#bulk-nds`, `#bulk-tax` → `applyRatesToAllUI()`, `saveRateEditsUI()`
- Hidden file input: `id="import-file-input"` `onchange="handleImportFileSelected(event)"`
- Modal: `id="import-modal"` (class `modal`) — title `#import-modal-title`, subtitle `#import-modal-subtitle`, body `#import-mapping-grid`, footer: `closeImportModal()` / `confirmImport()`
- Table: `#export-table` → thead `#export-thead`, tbody `#export-tbody`
- Status: `#export-status`, stats: `#export-stats`

**Tab "Параметры":** `#p-usd`, `#p-rmb`, `#p-log`, `#p-com`

**Includes at bottom:** `SheetJsLib` then `Scripts`

## Scripts.html — client-side JS

**STATE:** `{ exportData, calcResults, flakons, params, xlsxReady, importType, importFileData }`

**Import flow:**
1. `startImport(type)` — sets `STATE.importType`, lazy-loads SheetJS, triggers `#import-file-input`
2. `handleImportFileSelected(event)` — FileReader → `XLSX.read()` → `showMappingModal()`
3. `showMappingModal(fileName, headers, rowCount)` — builds mapping UI into `#import-mapping-grid`
   - `nomenclature`: 14-field table with auto-detect via `NOMENCLATURE_FIELDS[].hints`
   - `cost`/`supplier`: 2-field table (name col + value col)
4. `confirmImport()` — reads selects, calls `google.script.run.importNomenclature()` / `importCost1C()` / `importSupplierPrice()`

**NDS/TAX editing:**
- Cols 16/17 rendered as `<input class="nds-tax-input">` with closure event listeners
- `onNdsTaxChange(rowIdx, field, value)` → 1.5s debounce → `flushNdsTaxUpdates()` → `updateNdsTax(updates)`
- `applyRatesToAllUI()` → confirm dialog → `updateAllNdsTax(nds, tax)`
- `saveRateEditsUI()` → calls `flushNdsTaxUpdates()` immediately

**Constants:**
```js
EXPORT_NDS_COL = 16;  EXPORT_TAX_COL = 17;
FIELD_ORDER = ['NAME','ART','ART_WB','CAT','VOL','GRP2','GRP3','RAW',
               'FLAKON','SET_QTY','ART_MP','IS_SET','WEIGHT','GRP1'];
```

**Other functions:** `loadExportData()`, `renderExportTable()`, `filterExportTable()`,
`loadParamsUI()`, `saveParamsUI()`, `loadFlakonsUI()`, `saveFlakonDataUI()`,
`runCalculation()`, `renderCostTable()`, `populateCalcSelects()`, `onCalcSelect()`,
`calcManual()`, `runVerification()`, `renderVerification()`.

**Utilities:** `fmt(n)` — toFixed(2), `esc(s)` — XSS-safe escape, `colLabel(idx)` — A/B/C…

## Sample data files (.sample data/)
| File | Description |
|---|---|
| `1.xlsx` | Nomenclature export from 1C (4911 rows, 14 cols A-N) |
| `2.xls` | 1C cost metadata (2 rows — config format, not a plain list) |
| `3.xlsx` | Supplier prices (4511 rows, 12 cols; name=A, price=H) |

## clasp setup
- Config: `.clasp.json` in project root, `rootDir: ""`
- Executable: `C:\Users\Пользователь\AppData\Roaming\npm\clasp.cmd`
- Node: `C:\Program Files\nodejs`
- Auth: `~/.clasprc.json` (account: `zhuravel@rocknail.ru`)
- Push cmd: `clasp push` from project root
