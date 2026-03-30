# Performance Optimization Plan

## Context
Операции в WebApp (расчёт себестоимости, загрузка вкладок, сохранение правок) выполняются медленно. Основная причина — многократное чтение одних и тех же Google Sheets в рамках одного серверного вызова и поячеечная запись вместо batch-операций.

---

## Phase 1: Quick wins (наибольший эффект)

### 1.1 Execution-scoped кеш чтения листов
**Файл:** `DataService.js`
**Суть:** Добавить `_sheetCache` + `_cachedSheetRead(sheetName)` + `_invalidateCache(sheetName)` в начало файла. Заменить `getDataRange().getValues()` на кеш в:
- `getData()` (строка ~247)
- `getFlakonList()` (строки ~359, ~373 — два чтения)
- `getParams()` (строка ~296)
- `getStoredCostResults_()` (строка ~878)
- `getStoredBundleRows_()` (строка ~609)
- `getMetaValue_()` (строка ~1285)

Все write-функции должны вызывать `_invalidateCache()` после записи.
**Эффект:** -2–8 сек на каждый расчёт (устраняет 3-5 лишних чтений листов).

### 1.2 Batch-запись в updateImportRows
**Файл:** `DataService.js`, функция `updateImportRows` (строка ~181)
**Суть:** Вместо `setValue()` в цикле — один `getRange().getValues()` для колонок P-R, модификация массива в памяти, один `setValues()`.
**Эффект:** -10–30 сек при массовых правках (100 строк = 300 вызовов → 2 вызова).

### 1.3 Один read __meta в getImportSettings
**Файл:** `DataService.js`
**Суть:** Создать `getAllMetaValues_()` с кешем, заменить 4 вызова `getMetaValue_()` в `getImportSettings()` (строки ~115-118).
**Эффект:** -1–3 сек на загрузку настроек импорта.

### 1.4 Убрать read-after-write в recalculateAndStoreCostResults
**Файл:** `DataService.js`, функция `recalculateAndStoreCostResults` (строка ~799)
**Суть:** После `writeStoredCostResults_()` возвращать нормализованные результаты из памяти, не перечитывая лист.
**Эффект:** -0.5–2 сек на пересчёт.

---

## Phase 2: Средний приоритет

### 2.1 Lazy load данных + объединение серверных вызовов
**Файлы:** `Scripts.html`, `DataService.js`
**Суть:**
- Убрать `loadExportData()` из `window.load` (строка ~2196), перенести в `showTab('export')`
- Удалить дубликат `loadExportData()` (строка ~554)
- Создать `getExportPageData()` в DataService.js — возвращает `getData()` + `getBundleStats()` одним вызовом
- Обновить клиентский `loadExportData()` для одного вызова
**Эффект:** -2–5 сек на первую загрузку страницы.

### 2.2 Tabulator replaceData вместо destroy/create
**Файл:** `Scripts.html`, функция `renderExportTable` (строка ~692)
**Суть:** Если `STATE.tables.export` уже существует и структура колонок не изменилась — вызывать `replaceData(tableData)` вместо `destroy()` + `createTable()`.
**Эффект:** -200–500 мс на каждое обновление таблицы.

### 2.3 Асинхронная загрузка некритичных CDN-библиотек
**Файл:** `LibScripts.html`
**Суть:** Загружать Popper, Tippy, Toastify, SweetAlert2, Lucide динамически через `loadScript()` в `window.load`. Tabulator оставить синхронным.
**Эффект:** -0.5–2 сек на первую отрисовку.

---

## Phase 3: Крупные рефакторинги

### 3.1 Разделить getData() на raw + enrich
**Файлы:** `DataService.js`, `CalcEngine.js`
**Суть:** `getDataRaw_()` — чистые строки без обогащения типом. `enrichDataWithTypes(rows, flakonMap)` — добавление типа. `calculateAll` и `PlanService` используют `getDataRaw_()` когда тип не нужен, избегая вызова `getFlakonList()` внутри `getData()`.

### 3.2 Оптимизация saveBundleData — один write
**Файл:** `DataService.js`
**Суть:** Объединить `writeBundleSheetRows_()` + `refreshBundleSheetComputed_()` в один write. Передавать вычисленное состояние параметром, а не пересчитывать из листа.

### 3.3 Условные вызовы Protection API
**Файл:** `DataService.js`
**Суть:** В `removeSheetProtections_()` проверять наличие защит перед удалением. В `updateImportRows()` убрать защиту/переустановку для мелких правок.

---

## Сводка приоритетов

| # | Изменение | Файл | Экономия |
|---|-----------|------|----------|
| P0 | Кеш чтения листов | DataService.js | 2–8 сек |
| P0 | Batch-запись updateImportRows | DataService.js | 10–30 сек |
| P1 | Один read __meta | DataService.js | 1–3 сек |
| P1 | Убрать read-after-write | DataService.js | 0.5–2 сек |
| P1 | Lazy load + combined call | Scripts.html + DataService.js | 2–5 сек |
| P2 | replaceData для Tabulator | Scripts.html | 0.2–0.5 сек |
| P2 | Async CDN | LibScripts.html | 0.5–2 сек |
| P3 | Split getData | DataService.js + CalcEngine.js | архитектура |
| P3 | Single-write bundles | DataService.js | 2–5 сек |
| P3 | Conditional protection | DataService.js | 0.3–1 сек |

## Верификация
1. После Phase 1: открыть WebApp, перейти на вкладку «Себестоимость», нажать «Пересчитать» — должно быть заметно быстрее
2. На вкладке «Импорт данных» отредактировать несколько строк, нажать «Сохранить» — должно сохраняться быстро
3. Проверить что все расчёты дают те же числа что и до оптимизации
4. `clasp push` + обновить деплой, проверить в бою
