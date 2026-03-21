# mygasproject2 - Project Reference

## Purpose

Google Apps Script WebApp for себестоимость calculation inside the same Google Sheets document from which the script is launched. The app imports local Excel files, maintains working reference sheets, calculates cost for raw items, finished goods, flakons, and bundles, and writes calculation snapshots back to the active spreadsheet.

## Stack

- Google Apps Script
- HTML / CSS / vanilla JavaScript WebApp UI
- `clasp` for local development and deployment
- local `SheetJS` bundle in `SheetJsLib.html` for `.xls/.xlsx` import and `.xlsx` export

## Main files

| File | Role |
|---|---|
| `Код.js` | GAS entry points, menu, modal open, project sheet reset |
| `DataService.js` | Server-side imports, sheet I/O, params, flakons, bundles, result sheets, meta storage |
| `CalcEngine.js` | Core calculation logic for all item types, flakons, bundles, verification payloads |
| `WebApp.html` | Main WebApp layout and tab structure |
| `Scripts.html` | Client logic for tabs, filters, sorting, import flow, exports, editable tables |
| `Styles.html` | Full WebApp styling |
| `SheetJsLib.html` | Local SheetJS bundle |
| `appsscript.json` | GAS manifest |
| `.clasp.json` | `clasp` project config |

## Google Sheets usage

The project works only with the current spreadsheet through `SpreadsheetApp.getActiveSpreadsheet()`. It does not use `openById()` and does not connect to any external spreadsheet.

### Sheets used in the active spreadsheet

| Sheet | Purpose | Created automatically |
|---|---|---|
| `1cData` | Main import sheet. Stores imported nomenclature and calculation source fields. | Yes |
| `Флаконы` | Flakon reference and flakon cost calculation sheet. | Yes |
| `Наборы` | Bundle composition sheet with active specification, cost sources and manual fallback costs. | Yes |
| `basis` | Parameters sheet. Stores global calculation parameters and default rates. | Yes |
| `Результаты dd.MM.yyyy HH:mm` | Snapshot sheets with saved calculation results. | Yes |
| `__meta` | Hidden technical sheet for saved import mappings and defaults. | Yes |

## Data model

### Sheet `1cData`

Fixed 18-column working schema:

| Col | Index | Key | Meaning |
|---|---|---|---|
| A | 0 | `name` | Item name |
| B | 1 | `article` | Article |
| C | 2 | `articleWb` | WB article |
| D | 3 | `category` | Category |
| E | 4 | `volume` | Container volume |
| F | 5 | `group2` | Product group 2 |
| G | 6 | `group3` | Product group 3 |
| H | 7 | `raw` | Main raw material |
| I | 8 | `flakon` | Flakon / tara |
| J | 9 | `setQty` | Quantity inside set |
| K | 10 | `articleMp` | Marketplace article |
| L | 11 | `isSet` | Set flag (`Да` / `Нет`) |
| M | 12 | `weight` | Weight |
| N | 13 | `group1` | Product group 1 |
| O | 14 | `cost1C` | Cost from 1C |
| P | 15 | `supplierPrice` | Supplier price |
| Q | 16 | `nds` | NDS |
| R | 17 | `tax` | Duty |

### Sheet `basis`

Stored as one parameter row:

| Col | Key | Meaning |
|---|---|---|
| A | `usd` | USD exchange rate |
| B | `rmb` | RMB exchange rate |
| C | `log` | Logistics rate for regular items |
| D | `com` | Commission multiplier for regular items |
| E | `logFl` | Logistics rate for flakons |
| F | `comFl` | Commission multiplier for flakons |
| G | `importNds` | Default import NDS |
| H | `importTax` | Default import duty |
| I | `flakonNds` | Default flakon NDS |
| J | `flakonTax` | Default flakon duty |

### Sheet `Флаконы`

| Col | Key | Meaning |
|---|---|---|
| A | `name` | Flakon name |
| B | `volume` | Volume |
| C | `weight` | Weight |
| D | `supplierPrice` | Supplier price |
| E | `nds` | NDS |
| F | `tax` | Duty |
| G | `label` | Label cost |
| H | `rawFl` | Flakon raw cost in RUB |
| I | `deliveryFl` | Flakon delivery in RUB |
| J | `taxDutyFl` | Flakon NDS + duty |
| K | `totalFl` | Final flakon cost |

### Sheet `Наборы`

| Col | Key | Meaning |
|---|---|---|
| A | `component` | Component name inside the bundle |
| B | `bundle` | Bundle name |
| C | `specification` | Bundle specification name |
| D | `quantity` | Quantity of the component in the bundle |
| E | `active` | Active specification flag |
| F | `cost1C` | Component 1C cost pulled from `1cData` |
| G | `calcCost` | Component calculated cost pulled from calculation engine |
| H | `manualCost` | Manual RUB fallback cost |
| I | `usedCost` | Effective cost used for the bundle calculation |
| J | `source` | Cost source (`Расчёт`, `Ручная стоимость`, `Нет данных`, etc.) |

### Result sheets `Результаты ...`

Saved result schema:

| Col | Key | Meaning |
|---|---|---|
| A | `name` | Item name |
| B | `flakonName` | Flakon name |
| C | `category` | Category |
| D | `type` | Item type |
| E | `raw` | Raw material cost |
| F | `taxDuty` | NDS + duty |
| G | `delivery` | Delivery |
| H | `rawFl` | Flakon raw cost |
| I | `deliveryFl` | Flakon delivery |
| J | `taxDutyFl` | Flakon NDS + duty |
| K | `label` | Label cost |
| L | `totalFl` | Flakon total |
| M | `total` | Final total cost |
| N | `cost1C` | Cost from 1C |
| O | `diff` | Difference vs 1C |
| P | `diffPct` | Difference percent |

## Current WebApp tabs

### 1. `Параметры`

Purpose:
- maintain all calculation parameters and default import/flakon rates

Current fields:
- `usd`
- `rmb`
- `log`
- `com`
- `log_fl`
- `com_fl`
- `НДС для импорта`
- `Пошлина для импорта`
- `НДС для флаконов`
- `Пошлина для флаконов`

### 2. `Импорт данных`

Purpose:
- rebuild base `1cData`
- import 1C cost
- import supplier prices
- import bundle compositions
- manually edit `Цена поставщика`, `НДС`, `Пошлина`

Current behavior:
- local file import through `input[type=file]` + local SheetJS
- user selects file, sheet, header row and mapping
- mappings are stored in `__meta`
- cost and price imports match first by normalized item name, then by article
- table supports:
  - search
  - slicers
  - header sorting
  - sticky first displayed columns
- import stats include:
  - total rows
  - raw / finished / sets counts
  - filled 1C cost count
  - filled supplier price count
  - loaded bundle composition count

### 3. `Флаконы`

Purpose:
- maintain flakon list
- calculate flakon cost block used in item costing

Current behavior:
- list is built from unique values of `Тара (флакон)` in `1cData`
- source mapping for `Объём`, `Вес`, `Цена поставщика` uses:
  - `Флакон` = `Наименование` in `1cData`
- user can edit:
  - `Цена поставщика`
  - `НДС`
  - `Пошлина`
  - `Этикетка`
- buttons:
  - `Обновить`
  - `Пересчитать`
  - `Сохранить`

### 4. `Наборы`

Purpose:
- import and maintain bundle compositions
- select active specification per bundle
- assign manual fallback costs for repeated components without found себестоимость

Current behavior:
- bundle import reads local Excel report rows with:
  - component
  - bundle
  - specification
  - quantity
- one bundle may have multiple specifications
- active specification is chosen manually in the top table
- top table now:
  - shows all bundle positions found in `1cData`
  - highlights bundles without matched compositions
  - supports simple filtering and header sorting
- manual cost block stores one fallback RUB cost per unresolved component
- lower composition table shows:
  - component
  - bundle
  - specification
  - quantity
  - active flag
  - 1C component cost
  - calculated component cost
  - manual fallback
  - used effective cost
  - source
- lower table supports:
  - simple slicers
  - search
  - header sorting

### 5. `Себестоимость`

Purpose:
- run full cost calculation for all rows from `1cData`

Current behavior:
- supports search, slicers, sorting and zero-value highlighting
- export to `.xlsx`
- save results to dated sheets
- types currently supported:
  - `Сырьё`
  - `Готовый товар`
  - `Флакон`
  - `Наборы`

### 6. `Калькулятор`

Purpose:
- manual one-item calculation

Current behavior:
- can work from selected `1cData` item or fully manual input
- supports flakon overrides
- uses the same engine as batch calculation
- bundle logic is also routed through the same engine

### 7. `Проверка`

Purpose:
- show step-by-step formula breakdown for a selected item

Current behavior:
- uses the same engine payload as calculator verification
- includes bundle calculation breakdown by active specification and component costs

## Server-side API

### Entry points in `Код.js`

- `doGet()`
- `include(filename)`
- `onOpen()`
- `openWebApp()`
- `resetProjectSheets()`

### Main public functions in `DataService.js`

| Function | Purpose |
|---|---|
| `getData()` | Return `{ headers, rows }` from `1cData` |
| `getImportSettings()` | Return saved import mappings and default rates |
| `importBaseNomenclature(payload)` | Rebuild `1cData` from mapped nomenclature import |
| `importCurrentCost(payload)` | Update `Себестоимость 1С` in `1cData` |
| `importSupplierPriceData(payload)` | Update `Цена поставщика` in `1cData` |
| `importBundleCompositions(payload)` | Import and normalize bundle composition rows into sheet `Наборы` |
| `updateImportRows(updates)` | Batch-save manual import table edits |
| `getParams()` | Load params from `basis` |
| `saveParams(params)` | Save params to `basis` |
| `getFlakonList()` | Build / return normalized flakon list |
| `recalculateFlakonData(flakons)` | Recalculate flakon computed columns only |
| `saveFlakonData(flakons)` | Save flakon table to `Флаконы` |
| `getBundleData()` | Return normalized bundle UI payload |
| `getBundleStats()` | Return bundle import counters for the import tab |
| `saveBundleData(payload)` | Save active specifications and manual component costs |
| `saveResults(results)` | Create a dated result sheet |

## Calculation engine

### Type detection

Defined in `CalcEngine.js`:
- `Флакон` if item name matches a flakon row
- `Наборы` if `Это набор = Да`
- `Готовый товар` if there is no raw material
- `Сырьё` otherwise

### Main public functions in `CalcEngine.js`

| Function | Purpose |
|---|---|
| `determineType(row)` | Detect item type |
| `calculateAll(params)` | Calculate all rows from `1cData` |
| `calculateByIndex(index, params)` | Calculate one row by index |
| `calculateManual(input)` | Calculate one manual item |
| `getVerification(index, params)` | Return verification payload |

### Core item formulas

For `Готовый товар`:

- `raw = price * rmb * com`
- `delivery = log * weight * 1.15 * usd * com`
- `taxDuty = (raw + delivery + raw * tax) * nds + raw * tax`

For `Сырьё`:

- `raw = (price * rmb * com / 1000) * volume`
- `delivery = log * volume / 1000 * 1.15 * usd * com`
- `taxDuty = (raw + delivery + raw * tax) * nds + raw * tax`
- flakon block is added from `Флаконы`

For flakons:

- `rawFl = supplierPrice * rmb * comFl`
- `deliveryFl = logFl * weight * 1.15 * usd * comFl`
- `taxDutyFl = (rawFl + deliveryFl + rawFl * tax) * nds + rawFl * tax`
- `totalFl = rawFl + deliveryFl + taxDutyFl + label`

For bundles:

- active specification is selected per bundle
- bundle total = sum of all active specification component lines
- for each component line:
  - first try calculated component total
  - if missing, use manual fallback cost
  - if still missing, use `0`
- line total = `usedCost * quantity`
- bundle total = sum of all line totals

Final item output:

- `total = raw + taxDuty + delivery + totalFl`
- for bundles, `total` is the bundle total from active composition
- `diff = total - cost1C`
- `diffPct = (total - cost1C) / cost1C`

## Client-side notes

Main client state in `Scripts.html` includes:

- imported `1cData`
- calculated result set
- flakon list
- bundle UI payload
- params
- import mapping context
- table filter and sorting state

Important client capabilities:

- local Excel import through SheetJS
- `.xlsx` export for cost results
- editable import table
- editable flakon table
- bundle specification selection and manual fallback costs
- search, slicers and sorting on working tables
- instruction blocks on operational tabs

## Sample input files

Folder: `.sample data`

| File | Meaning |
|---|---|
| `1.xlsx` | Base nomenclature import example |
| `2.xls` | Current cost import example |
| `3.xlsx` | Supplier price import example |
| `4.xlsx` | Bundle composition import example |

## `clasp` notes

- project config is in `.clasp.json`
- `rootDir` is the repository root
- deploy from project root with `clasp push`
- if the WebApp URL still shows old behavior after push, update the Apps Script deployment version in GAS
