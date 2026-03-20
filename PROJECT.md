# mygasproject2 — Project Reference

## Purpose

Google Apps Script WebApp for cost calculation inside the same Google Sheets document from which the script is launched. The project imports nomenclature and prices from local Excel files, calculates item cost, calculates flakon cost separately, and writes calculation results back to sheets in the active spreadsheet.

## Stack

- Google Apps Script (server-side)
- HTML/CSS/JavaScript WebApp UI
- `clasp` for local development and deployment
- local `SheetJS` bundle in `SheetJsLib.html` for `.xls/.xlsx` import and `.xlsx` export

## Main files

| File | Role |
|---|---|
| `Код.js` | GAS entry points: `doGet()`, `include()`, `onOpen()`, `openWebApp()` |
| `DataService.js` | Server-side work with Google Sheets, imports, params, flakons, result sheets, meta storage |
| `CalcEngine.js` | Core calculation logic for items and flakons |
| `WebApp.html` | Main WebApp layout and tabs |
| `Scripts.html` | All client-side logic for tabs, filters, sorting, import flow, exports |
| `Styles.html` | All styles for the WebApp |
| `SheetJsLib.html` | Local SheetJS bundle |
| `appsscript.json` | GAS manifest |
| `.clasp.json` | `clasp` project config |

## Google Sheets usage

The project works only with the current spreadsheet via `SpreadsheetApp.getActiveSpreadsheet()`. It does not use `openById()` and does not connect to an external spreadsheet.

### Sheets used in the active spreadsheet

| Sheet | Purpose | Created automatically |
|---|---|---|
| `1cData` | Main import sheet. Stores imported nomenclature plus cost fields used by all calculations. | Yes |
| `Флаконы` | Flakon reference and flakon cost calculation sheet. Stores manual and calculated flakon data. | Yes |
| `basis` | Parameters sheet. Stores `usd`, `rmb`, `log`, `com`, `log_fl`, `com_fl`. | Yes |
| `Результаты dd.MM.yyyy HH:mm` | Snapshot sheets with saved calculation results. A new sheet is created on each save. | Yes |
| `__meta` | Hidden technical sheet for saved import mappings and default import rates. | Yes |

## Data model

### Sheet `1cData`

Project schema is fixed to 18 columns.

| Col | Index | Key | Label |
|---|---|---|---|
| A | 0 | `name` | `Номенклатура.Наименование` |
| B | 1 | `article` | `Артикул` |
| C | 2 | `articleWb` | `Артикул ВБ` |
| D | 3 | `category` | `Категория товаров` |
| E | 4 | `volume` | `Объем тары` |
| F | 5 | `group2` | `Номенклатура.Товарная группа 2 (Общие)` |
| G | 6 | `group3` | `Номенклатура.Товарная группа 3 (Общие)` |
| H | 7 | `raw` | `Номенклатура.Основное сырье (Общие)` |
| I | 8 | `flakon` | `Номенклатура.Тара (флакон) (Общие)` |
| J | 9 | `setQty` | `Номенклатура.Количество лаков в наборе (Общие)` |
| K | 10 | `articleMp` | `Номенклатура.Артикул МП` |
| L | 11 | `isSet` | `Номенклатура.Это набор (RockNail)` |
| M | 12 | `weight` | `Номенклатура.Вес (числитель)` |
| N | 13 | `group1` | `Номенклатура.Товарная группа 1 (Общие)` |
| O | 14 | `cost1C` | `Себестоимость 1С` |
| P | 15 | `supplierPrice` | `Цена поставщика` |
| Q | 16 | `nds` | `НДС` |
| R | 17 | `tax` | `Пошлина` |

### Sheet `basis`

Stored as a 6-column parameter row:

| Col | Key | Meaning |
|---|---|---|
| A | `usd` | USD exchange rate |
| B | `rmb` | RMB exchange rate |
| C | `log` | Base logistics |
| D | `com` | Base commission |
| E | `logFl` | Logistics for flakons |
| F | `comFl` | Commission for flakons |

### Sheet `Флаконы`

Current stored schema:

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
| K | `totalFl` | Total flakon cost |

### Result sheets `Результаты ...`

Saved result schema:

| Col | Key | Meaning |
|---|---|---|
| A | `name` | Item name |
| B | `type` | Item type |
| C | `raw` | Raw material cost |
| D | `taxDuty` | NDS + duty |
| E | `delivery` | Delivery |
| F | `rawFl` | Flakon raw cost |
| G | `deliveryFl` | Flakon delivery |
| H | `taxDutyFl` | Flakon NDS + duty |
| I | `label` | Label cost |
| J | `totalFl` | Flakon total |
| K | `total` | Final total cost |
| L | `cost1C` | Cost from 1C |
| M | `diff` | Difference vs 1C |
| N | `diffPct` | Difference percent |

## Current WebApp tabs

### 1. `Импорт данных`

Purpose:
- build base `1cData` from a local Excel file
- import current 1C cost into `Себестоимость 1С`
- import supplier prices into `Цена поставщика`
- manually edit `Цена поставщика`, `НДС`, `Пошлина`

Current behavior:
- file import is local through `input[type=file]` + SheetJS
- user selects file, sheet, header row, and column mapping
- import mappings are stored in `__meta` and reused later
- for value imports matching is:
  - first by normalized item name
  - fallback by article
- `НДС` and `Пошлина` can be set as defaults on first nomenclature import
- `НДС`, `Пошлина`, `Цена поставщика` can be edited manually in the table
- table supports:
  - search
  - slicers above the table
  - column sorting
  - sticky first four displayed columns

### 2. `Параметры`

Purpose:
- edit calculation parameters used by item and flakon formulas

Current fields:
- `usd`
- `rmb`
- `log`
- `com`
- `log_fl`
- `com_fl`

### 3. `Флаконы`

Purpose:
- maintain flakon list
- calculate flakon cost blocks used in item costing

Current behavior:
- flakon list is built from unique values of `Тара (флакон)` in `1cData`
- data is merged with the saved `Флаконы` sheet
- `Объём`, `Вес`, `Цена поставщика` are mapped by rule:
  - `Флакон` equals `Наименование` in `1cData`
- manual values are preserved by flakon name:
  - `Цена поставщика`
  - `НДС`
  - `Пошлина`
  - `Этикетка`
- buttons:
  - `Обновить` rebuilds the list from `1cData`
  - `Пересчитать` recalculates only computed columns without page reload
  - `Сохранить` writes the current table to sheet `Флаконы`
- tab also contains an instruction block for the user

### 4. `Себестоимость`

Purpose:
- run full cost calculation for all rows from `1cData`

Current behavior:
- button `Рассчитать` runs `calculateAll(params)`
- table supports:
  - search
  - slicers
  - sorting by clicking headers
  - zero-value highlighting for numeric cells
- button `Выгрузить в Excel` exports the full result set to `.xlsx`
- button `Записать результаты на лист` creates a new `Результаты ...` sheet

### 5. `Калькулятор`

Purpose:
- manual one-item calculation

Current behavior:
- user can select an existing item or enter values manually
- uses the same calculation engine as batch calculation

### 6. `Проверка`

Purpose:
- show step-by-step formula breakdown for a selected item

Current behavior:
- uses `getVerification(index, params)`
- shows formulas, substitutions and intermediate results for both item and flakon parts

## Server-side API

### Entry points

Defined in `Код.js`:
- `doGet()`
- `include(filename)`
- `onOpen()`
- `openWebApp()`

### Main public functions in `DataService.js`

| Function | Purpose |
|---|---|
| `getData()` | Returns `{ headers, rows }` from `1cData` |
| `getDataHeaders()` | Returns project header schema |
| `getImportSettings()` | Returns saved import mappings and default rates from `__meta` |
| `importBaseNomenclature(payload)` | Rebuilds `1cData` from mapped nomenclature import |
| `importCurrentCost(payload)` | Updates `Себестоимость 1С` in `1cData` |
| `importSupplierPriceData(payload)` | Updates `Цена поставщика` in `1cData` |
| `updateImportRows(updates)` | Batch-save manual edits in import table |
| `updateNdsTax(updates)` | Alias of `updateImportRows` |
| `updateAllNdsTax(nds, tax)` | Applies the same NDS and duty to all rows |
| `getParams()` | Loads parameters from `basis` |
| `saveParams(params)` | Saves parameters to `basis` |
| `getFlakonList()` | Builds/returns normalized flakon list |
| `recalculateFlakonData(flakons)` | Recalculates computed flakon columns only |
| `saveFlakonData(flakons)` | Saves the flakon table to sheet `Флаконы` |
| `saveResults(results)` | Creates a dated `Результаты ...` sheet |

## Calculation engine

### Item type detection

Defined in `CalcEngine.js`:
- `Наборы` if `Это набор = Да`
- `Готовый товар` if there is no raw material
- `Сырьё` otherwise

### Main public functions in `CalcEngine.js`

| Function | Purpose |
|---|---|
| `determineType(row)` | Detects item type |
| `calculateAll(params)` | Calculates full `1cData` |
| `calculateByIndex(index, params)` | Calculates one row by index |
| `calculateManual(input)` | Calculates one manual item |
| `getVerification(index, params)` | Returns a formula breakdown |

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

Final:

- `total = raw + taxDuty + delivery + totalFl`
- `diff = total - cost1C`
- `diffPct = (total - cost1C) / cost1C`

### Flakon formulas

Flakon metrics are calculated from flakon row data plus `log_fl` and `com_fl`:

- `rawFl = supplierPrice * rmb * comFl`
- `deliveryFl = logFl * weight * 1.15 * usd * comFl`
- `taxDutyFl = (rawFl + deliveryFl + rawFl * tax) * nds + rawFl * tax`
- `totalFl = rawFl + deliveryFl + taxDutyFl + label`

## Client-side notes

Main state in `Scripts.html` includes:

- imported table data
- calculated results
- flakon list
- current params
- import mapping state
- slicer and sorting state for tables

Important client capabilities:

- local Excel import through SheetJS
- local Excel export for the cost tab
- table search, slicers, sorting
- manual batch-saving of editable table fields
- user instruction blocks on key tabs

## Sample input files

Folder: `.sample data`

| File | Meaning |
|---|---|
| `1.xlsx` | Base nomenclature import example |
| `2.xls` | Current cost import example |
| `3.xlsx` | Supplier price import example |

## `clasp` notes

- project config is in `.clasp.json`
- `rootDir` is the repository root
- deploy from project root with `clasp push`
- after push, if the WebApp URL still shows old behavior, update the Apps Script deployment version in GAS

