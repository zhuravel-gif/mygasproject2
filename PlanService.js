/**
 * PLAN SERVICE — импорт и расчёт затрат по плану продаж.
 */

var PLAN_COST_CFG = {
  SHEET: 'План затрат',
  META_KEY: 'plan.importConfig',
  PREVIEW_ROWS: 120,
  MONTH_KEYS: ['month1', 'month2', 'month3'],
  HEADERS: [
    'План.Наименование',
    'План.Артикул ВБ',
    'Совпадение найдено',
    'Тип совпадения',
    'Индекс строки 1cData',
    'Номенклатура',
    'Артикул',
    'Артикул ВБ',
    'Категория товара',
    'Товарная группа 1',
    'Цена поставщика',
    'Вес',
    'Объём тары',
    'Тара',
    'Основное сырьё',
    'Тип',
    'План 1',
    'План 2',
    'План 3',
    'Примечание'
  ]
};

var PLAN_COL = {
  PLAN_NAME: 0,
  PLAN_ARTICLE_WB: 1,
  MATCHED: 2,
  MATCH_BY: 3,
  SOURCE_ROW_INDEX: 4,
  NAME: 5,
  ARTICLE: 6,
  ARTICLE_WB: 7,
  CATEGORY: 8,
  GROUP1: 9,
  SUPPLIER_PRICE: 10,
  WEIGHT: 11,
  VOLUME: 12,
  FLAKON: 13,
  RAW: 14,
  TYPE: 15,
  MONTH1: 16,
  MONTH2: 17,
  MONTH3: 18,
  NOTE: 19
};

function buildPlanDiagnostic_(err) {
  if (!err) return '';
  if (typeof err === 'string') return err;
  if (err.stack) return String(err.stack);
  if (err.message) return String(err.message);
  try {
    return JSON.stringify(err);
  } catch (jsonErr) {
    return String(err);
  }
}

function planErrorResponse_(message, diagnostic, extra) {
  var payload = {
    success: false,
    message: String(message || 'Ошибка операции плана.')
  };
  var diagText = buildPlanDiagnostic_(diagnostic);
  if (diagText) payload.diagnostic = diagText;
  if (diagnostic && diagnostic.name) payload.errorName = String(diagnostic.name);
  if (extra) {
    for (var key in extra) {
      if (Object.prototype.hasOwnProperty.call(extra, key)) payload[key] = extra[key];
    }
  }
  return payload;
}

function extractPlanSpreadsheetId_(value) {
  var input = String(value || '').trim();
  if (!input) return '';

  var pathMatch = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  if (pathMatch && pathMatch[1]) return pathMatch[1];

  var queryMatch = input.match(/[?&]id=([a-zA-Z0-9-_]+)/i);
  if (queryMatch && queryMatch[1]) return queryMatch[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(input)) return input;
  return '';
}

function openPlanSpreadsheet_(value) {
  var input = String(value || '').trim();
  if (!input) throw new Error('Укажите ссылку или ID Google Sheets.');

  var id = extractPlanSpreadsheetId_(input);
  var byIdError = null;
  var byUrlError = null;

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      byIdError = err;
    }
  }

  if (/^https?:\/\//i.test(input)) {
    try {
      return SpreadsheetApp.openByUrl(input);
    } catch (err2) {
      byUrlError = err2;
    }
  }

  var details = [];
  if (byIdError) details.push('openById: ' + byIdError.message);
  if (byUrlError) details.push('openByUrl: ' + byUrlError.message);
  if (!details.length) details.push('Не удалось извлечь корректный ID таблицы из введённого значения.');

  throw new Error(details.join(' | '));
}

function getPlanSheetType_(sheet) {
  try {
    var type = sheet && sheet.getType ? sheet.getType() : '';
    return type ? String(type) : 'GRID';
  } catch (err) {
    return 'GRID';
  }
}

function getPlanDataSourceColumnNames_(dataSourceSheet, limitCols) {
  var dataSource = dataSourceSheet.getDataSource();
  var columns = dataSource && dataSource.getColumns ? dataSource.getColumns() : [];
  var names = [];

  for (var i = 0; i < columns.length; i++) {
    if (limitCols && names.length >= limitCols) break;
    var name = '';
    try {
      name = String(columns[i].getName ? columns[i].getName() : '').trim();
    } catch (err) {
      name = '';
    }
    names.push(name || ('Колонка ' + (i + 1)));
  }

  return names;
}

function buildPlanRowsFromColumnArrays_(headers, columnArrays) {
  var maxRows = 0;
  for (var i = 0; i < columnArrays.length; i++) {
    if (columnArrays[i] && columnArrays[i].length > maxRows) maxRows = columnArrays[i].length;
  }

  var rows = [headers.slice()];
  for (var r = 0; r < maxRows; r++) {
    var row = [];
    for (var c = 0; c < headers.length; c++) {
      row.push(columnArrays[c] && r < columnArrays[c].length ? columnArrays[c][r] : '');
    }
    rows.push(row);
  }
  return rows;
}

function readPlanDataSourceSheet_(sheet, rowLimit, colLimit) {
  var dataSourceSheet = sheet && sheet.asDataSourceSheet ? sheet.asDataSourceSheet() : null;
  if (!dataSourceSheet) throw new Error('Не удалось получить DataSourceSheet для выбранного листа.');

  var headers = getPlanDataSourceColumnNames_(dataSourceSheet, colLimit);
  if (!headers.length) throw new Error('В DATASOURCE-листе не найдено доступных колонок.');

  var columnArrays = [];
  for (var i = 0; i < headers.length; i++) {
    var values = rowLimit
      ? dataSourceSheet.getSheetValues(headers[i], 1, rowLimit)
      : dataSourceSheet.getSheetValues(headers[i]);
    columnArrays.push(values || []);
  }

  var rows = buildPlanRowsFromColumnArrays_(headers, columnArrays);
  return {
    rows: rows,
    headerRowLocked: true,
    rowCount: rows.length,
    colCount: headers.length
  };
}

function readPlanSheetSourceRows_(sheet, rowLimit, colLimit) {
  var type = getPlanSheetType_(sheet);
  if (type === 'DATASOURCE') {
    return readPlanDataSourceSheet_(sheet, rowLimit, colLimit);
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var rows = [];
  if (lastRow > 0 && lastCol > 0) {
    rows = sheet
      .getRange(1, 1, Math.min(lastRow, rowLimit || lastRow), Math.min(lastCol, colLimit || lastCol))
      .getDisplayValues();
  }

  return {
    rows: rows,
    headerRowLocked: false,
    rowCount: lastRow,
    colCount: lastCol
  };
}

function buildPlanSourceSheetInfo_(sheet) {
  var info = {
    name: sheet.getName(),
    type: getPlanSheetType_(sheet),
    rowCount: 0,
    colCount: 0,
    previewRows: [],
    previewError: '',
    selectable: true,
    headerRowLocked: false
  };

  try {
    var sourceData = readPlanSheetSourceRows_(sheet, PLAN_COST_CFG.PREVIEW_ROWS - 1, null);
    info.rowCount = sourceData.rowCount;
    info.colCount = sourceData.colCount;
    info.previewRows = sourceData.rows;
    info.headerRowLocked = !!sourceData.headerRowLocked;
  } catch (err) {
    info.selectable = false;
    info.previewError = 'Не удалось прочитать лист: ' + err.message;
  }

  return info;
}

function getPlanImportSource(url) {
  var cleanUrl = String(url || '').trim();
  if (!cleanUrl) return planErrorResponse_('Укажите ссылку на Google Sheets.');

  try {
    var ss = openPlanSpreadsheet_(cleanUrl);
    var sheets = ss.getSheets().map(buildPlanSourceSheetInfo_);

    return {
      success: true,
      spreadsheetName: ss.getName(),
      sheets: sheets
    };
  } catch (err) {
    return planErrorResponse_('Не удалось открыть Google Sheets по ссылке: ' + err.message, err);
  }
}

function importPlanCosts(payload) {
  try {
    payload = payload || {};
    var config = normalizePlanImportConfig_(payload);
    if (!config.spreadsheetUrl) return planErrorResponse_('Не указана ссылка на Google Sheets.');
    if (!config.sheetName) return planErrorResponse_('Не выбран лист источника.');
    if (config.headerRow < 1) return planErrorResponse_('Некорректная строка заголовков.');
    if (!isValidPlanImportPayload_(config)) {
      return planErrorResponse_('Проверьте маппинг колонок и месяцы плана.');
    }

    var externalSheet;
    try {
      externalSheet = openPlanSpreadsheet_(config.spreadsheetUrl).getSheetByName(config.sheetName);
    } catch (err) {
      return planErrorResponse_('Не удалось открыть источник плана: ' + err.message, err);
    }
    if (!externalSheet) return planErrorResponse_('Лист источника не найден.');

    var sourceData;
    try {
      sourceData = readPlanSheetSourceRows_(externalSheet, null, null);
    } catch (err2) {
      return planErrorResponse_('Не удалось прочитать строки источника плана: ' + err2.message, err2);
    }

    if (sourceData.headerRowLocked) config.headerRow = 1;
    var sourceRows = sourceData.rows || [];
    var lastRow = sourceRows.length;
    var lastCol = sourceRows[0] ? sourceRows[0].length : 0;
    if (lastRow < config.headerRow || lastCol < 1) {
      return planErrorResponse_('В выбранном листе нет данных для импорта.');
    }

    var headers = buildNormalizedHeaders_(sourceRows[config.headerRow - 1] || []);
    var dataRows = sourceRows.slice(config.headerRow).filter(isRowWithValues_);
    if (!dataRows.length) {
      return planErrorResponse_('После строки заголовков не найдено строк плана.');
    }

    var dataObj = getData();
    var dataRowsCurrent = dataObj.rows || [];
    var matchMaps = buildPlanMatchMaps_(dataRowsCurrent);
    var flakonMap = buildFlakonMap(getFlakonList());
    var aggregated = {};
    var processed = 0;
    var matchedByArticle = 0;
    var matchedByName = 0;

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      var planName = getPlanCellString_(row, config.mapping.name);
      var planArticleWb = getPlanCellString_(row, config.mapping.articleWb);
      var monthQtys = config.months.map(function(month) {
        return toPlanNumber_(isValidIndex_(month.column, headers.length) ? row[month.column] : '');
      });

      if (!planName && !planArticleWb) continue;
      processed++;

      var match = matchPlanRow_(planName, planArticleWb, matchMaps);
      if (match.matchBy === 'articleWb') matchedByArticle++;
      if (match.matchBy === 'name') matchedByName++;

      var aggregateKey = match.rowIndex >= 0
        ? ('matched:' + match.rowIndex)
        : ('unmatched:' + normalizeMatchKey_(planArticleWb) + '||' + normalizeMatchKey_(planName));

      if (!aggregated[aggregateKey]) {
        aggregated[aggregateKey] = buildInitialPlanItem_(planName, planArticleWb, monthQtys, match, dataRowsCurrent, flakonMap);
      } else {
        for (var q = 0; q < monthQtys.length; q++) {
          aggregated[aggregateKey].monthQtys[q] = round2_(toNumber_(aggregated[aggregateKey].monthQtys[q], 0) + monthQtys[q]);
        }
      }
    }

    var items = mapValues_(aggregated);
    items.sort(function(a, b) {
      if (!!a.matched !== !!b.matched) return a.matched ? -1 : 1;
      return String(a.name || a.planName || '').localeCompare(String(b.name || b.planName || ''), 'ru', { sensitivity: 'base', numeric: true });
    });

    var deduped = items.length;
    var activeMonths = getActivePlanMonths_(config.months);
    var importSummary = buildPlanImportSummary_(items, activeMonths, {
      processed: processed,
      matchedByArticle: matchedByArticle,
      matchedByName: matchedByName,
      deduped: deduped
    });

    config.mapping.headers = headers;
    config.importSummary = importSummary;
    config.importedAt = new Date().toISOString();

    writePlanSheetRows_(items);
    savePlanImportConfig_(config);
    var state = null;
    var stateLoadError = '';
    try {
      state = getPlanCostState();
    } catch (err3) {
      stateLoadError = err3 && err3.message ? String(err3.message) : 'Не удалось обновить состояние вкладки после импорта.';
    }

    return {
      success: true,
      message: 'План импортирован: ' + importSummary.matchedItems + ' сопоставлено, ' + importSummary.unmatchedItems + ' не сопоставлено.',
      importSummary: importSummary,
      state: state,
      stateLoadError: stateLoadError
    };
  } catch (err4) {
    return planErrorResponse_('Непредвиденная ошибка импорта плана: ' + err4.message, err4);
  }
}

function buildPlanCostState_(includeItems) {
  var config = getPlanImportConfig_();
  var items = getStoredPlanRows_();
  var months = getActivePlanMonths_(config.months || buildDefaultPlanMonths_());
  var payload = {
    success: true,
    months: months,
    stats: buildPlanStateStats_(items, months),
    importSummary: config.importSummary || buildPlanImportSummary_(items, months, {}),
    importConfig: config,
    itemCount: items.length
  };

  if (includeItems) payload.items = items;
  return payload;
}

function getPlanCostState() {
  try {
    return buildPlanCostState_(false);
  } catch (err) {
    return planErrorResponse_('Не удалось загрузить состояние плана: ' + err.message, err);
  }
}

function calculatePlanCosts(monthKey) {
  try {
    var state = buildPlanCostState_(true);
    if (!state || !state.success) {
      return planErrorResponse_(
        (state && state.message) || 'Не удалось загрузить данные плана перед расчётом.',
        state && state.diagnostic ? state.diagnostic : ''
      );
    }

    var items = state.items || [];
    var months = state.months || [];
    if (!months.length) {
      return planErrorResponse_('Сначала импортируйте план.');
    }

    var month = months[0];
    for (var i = 0; i < months.length; i++) {
      if (months[i].key === monthKey) {
        month = months[i];
        break;
      }
    }

    var dataObj = getData();
    var dataRows = dataObj.rows || [];
    if (!dataRows.length) return planErrorResponse_('Сначала загрузите данные на вкладке Импорт данных.');

    var params = getParams();
    var flakonMap = buildFlakonMap(getFlakonList());
    var bundleContext = buildBundleContext_(params, flakonMap, dataRows);
    var results = [];
    var unresolvedItems = [];

    for (i = 0; i < items.length; i++) {
      var item = items[i];
      var planQty = round2_(toNumber_(item.monthQtys[month.slotIndex], 0));
      if (planQty <= 0) continue;

      var matchedRow = resolvePlanDataRow_(item, dataRows);
      if (!matchedRow) {
        unresolvedItems.push(item);
        continue;
      }

      var calcResult = calculateOne(matchedRow, params, flakonMap, null, null, bundleContext, {});
      var rowResult = buildPlanResultRow_(item, matchedRow, calcResult, planQty, false, calcResult.type === 'Наборы' ? 'Комплектация набора' : 'Расчёт');

      if (calcResult.type === 'Наборы') {
        rowResult._children = buildPlanBundleChildren_(item, matchedRow, planQty, params, flakonMap, bundleContext);
        applyPlanChildrenTotals_(rowResult);
      }

      results.push(rowResult);
    }

    results.sort(function(a, b) {
      if (a.type !== b.type) return String(a.type || '').localeCompare(String(b.type || ''), 'ru', { sensitivity: 'base', numeric: true });
      return String(a.name || '').localeCompare(String(b.name || ''), 'ru', { sensitivity: 'base', numeric: true });
    });

    return {
      success: true,
      monthKey: month.key,
      monthLabel: month.label,
      results: results,
      stats: buildPlanCalculationStats_(results, unresolvedItems)
    };
  } catch (err) {
    return planErrorResponse_('Ошибка расчёта плана: ' + err.message, err);
  }
}

function isValidPlanImportPayload_(config) {
  if (!config.mapping) return false;
  if (!isValidIndex_(config.mapping.name, 9999) && !isValidIndex_(config.mapping.articleWb, 9999)) return false;
  if (!config.months || !config.months.length) return false;

  var usedColumns = {};
  var activeMonths = 0;
  for (var i = 0; i < config.months.length; i++) {
    var month = config.months[i];
    var hasLabel = !!String(month.label || '').trim();
    var hasColumn = isValidIndex_(month.column, 9999);
    if (!hasLabel && !hasColumn) continue;
    if (!hasLabel || !hasColumn) return false;
    if (usedColumns[month.column]) return false;
    usedColumns[month.column] = true;
    activeMonths++;
  }
  return activeMonths > 0;
}

function normalizePlanImportConfig_(payload) {
  var months = [];
  var inputMonths = payload.months || [];
  for (var i = 0; i < PLAN_COST_CFG.MONTH_KEYS.length; i++) {
    var src = inputMonths[i] || {};
    months.push({
      key: PLAN_COST_CFG.MONTH_KEYS[i],
      label: String(src.label || '').trim(),
      column: parseInt(src.column, 10),
      header: String(src.header || '').trim()
    });
  }

  return {
    spreadsheetUrl: String(payload.spreadsheetUrl || '').trim(),
    spreadsheetName: String(payload.spreadsheetName || '').trim(),
    sheetName: String(payload.sheetName || '').trim(),
    headerRow: Math.max(parseInt(payload.headerRow, 10) || 1, 1),
    mapping: {
      name: parseInt(payload.mapping && payload.mapping.name, 10),
      articleWb: parseInt(payload.mapping && payload.mapping.articleWb, 10),
      nameHeader: String(payload.mapping && payload.mapping.nameHeader || '').trim(),
      articleWbHeader: String(payload.mapping && payload.mapping.articleWbHeader || '').trim(),
      headers: payload.mapping && payload.mapping.headers ? payload.mapping.headers : []
    },
    months: months
  };
}

function getPlanImportConfig_() {
  var raw = getMetaValue_(PLAN_COST_CFG.META_KEY);
  if (!raw) {
    return {
      spreadsheetUrl: '',
      spreadsheetName: '',
      sheetName: '',
      headerRow: 1,
      mapping: { name: -1, articleWb: -1, nameHeader: '', articleWbHeader: '', headers: [] },
      months: buildDefaultPlanMonths_(),
      importSummary: null
    };
  }

  try {
    var parsed = JSON.parse(raw);
    parsed.mapping = parsed.mapping || { name: -1, articleWb: -1, nameHeader: '', articleWbHeader: '', headers: [] };
    parsed.mapping.name = parseInt(parsed.mapping.name, 10);
    parsed.mapping.articleWb = parseInt(parsed.mapping.articleWb, 10);
    parsed.mapping.nameHeader = String(parsed.mapping.nameHeader || '').trim();
    parsed.mapping.articleWbHeader = String(parsed.mapping.articleWbHeader || '').trim();
    parsed.mapping.headers = parsed.mapping.headers || [];
    parsed.months = parsed.months && parsed.months.length === 3 ? parsed.months : buildDefaultPlanMonths_();
    for (var i = 0; i < parsed.months.length; i++) {
      parsed.months[i].key = parsed.months[i].key || PLAN_COST_CFG.MONTH_KEYS[i];
      parsed.months[i].label = String(parsed.months[i].label || '').trim();
      parsed.months[i].column = parseInt(parsed.months[i].column, 10);
      parsed.months[i].header = String(parsed.months[i].header || '').trim();
    }
    parsed.headerRow = Math.max(parseInt(parsed.headerRow, 10) || 1, 1);
    return parsed;
  } catch (err) {
    return {
      spreadsheetUrl: '',
      spreadsheetName: '',
      sheetName: '',
      headerRow: 1,
      mapping: { name: -1, articleWb: -1, nameHeader: '', articleWbHeader: '', headers: [] },
      months: buildDefaultPlanMonths_(),
      importSummary: null
    };
  }
}

function savePlanImportConfig_(config) {
  setMetaValue_(PLAN_COST_CFG.META_KEY, JSON.stringify(config || {}));
}

function buildDefaultPlanMonths_() {
  var months = [];
  for (var i = 0; i < PLAN_COST_CFG.MONTH_KEYS.length; i++) {
    months.push({
      key: PLAN_COST_CFG.MONTH_KEYS[i],
      label: '',
      column: -1,
      header: ''
    });
  }
  return months;
}

function getActivePlanMonths_(months) {
  var source = months && months.length ? months : buildDefaultPlanMonths_();
  var active = [];

  for (var i = 0; i < source.length; i++) {
    var month = source[i] || {};
    var label = String(month.label || '').trim();
    var column = parseInt(month.column, 10);
    if (!label || !isValidIndex_(column, 9999)) continue;
    active.push({
      key: month.key || PLAN_COST_CFG.MONTH_KEYS[i] || ('month' + (i + 1)),
      label: label,
      column: column,
      header: String(month.header || '').trim(),
      slotIndex: i
    });
  }

  return active;
}

function buildNormalizedHeaders_(row) {
  var headers = [];
  for (var i = 0; i < row.length; i++) {
    headers.push(String(row[i] || '').trim() || ('Колонка ' + (i + 1)));
  }
  return headers;
}

function isRowWithValues_(row) {
  if (!row) return false;
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return true;
  }
  return false;
}

function isValidIndex_(index, length) {
  return typeof index === 'number' && !isNaN(index) && index >= 0 && index < length;
}

function getPlanCellString_(row, index) {
  return isValidIndex_(index, row.length) ? String(row[index] || '').trim() : '';
}

function toPlanNumber_(value) {
  var parsed = coerceNumber_(value);
  if (parsed === '' || parsed === null || parsed === undefined) return 0;
  return toNumber_(parsed, 0);
}

function buildPlanMatchMaps_(rows) {
  var byArticleWb = {};
  var byName = {};
  for (var i = 0; i < rows.length; i++) {
    var articleWbKey = normalizeMatchKey_(rows[i][2]);
    var nameKey = normalizeMatchKey_(rows[i][0]);
    if (articleWbKey && !Object.prototype.hasOwnProperty.call(byArticleWb, articleWbKey)) byArticleWb[articleWbKey] = i;
    if (nameKey && !Object.prototype.hasOwnProperty.call(byName, nameKey)) byName[nameKey] = i;
  }
  return {
    byArticleWb: byArticleWb,
    byName: byName
  };
}

function matchPlanRow_(planName, planArticleWb, maps) {
  var articleKey = normalizeMatchKey_(planArticleWb);
  var nameKey = normalizeMatchKey_(planName);
  if (articleKey && Object.prototype.hasOwnProperty.call(maps.byArticleWb, articleKey)) {
    return { rowIndex: maps.byArticleWb[articleKey], matchBy: 'articleWb' };
  }
  if (nameKey && Object.prototype.hasOwnProperty.call(maps.byName, nameKey)) {
    return { rowIndex: maps.byName[nameKey], matchBy: 'name' };
  }
  return { rowIndex: -1, matchBy: '' };
}

function buildInitialPlanItem_(planName, planArticleWb, monthQtys, match, dataRows, flakonMap) {
  var matchedRow = match.rowIndex >= 0 ? dataRows[match.rowIndex] : null;
  var item = {
    planName: planName,
    planArticleWb: planArticleWb,
    matched: !!matchedRow,
    matchBy: match.matchBy || '',
    sourceRowIndex: match.rowIndex,
    name: matchedRow ? String(matchedRow[0] || '').trim() : '',
    article: matchedRow ? String(matchedRow[1] || '').trim() : '',
    articleWb: matchedRow ? String(matchedRow[2] || '').trim() : '',
    category: matchedRow ? String(matchedRow[3] || '').trim() : '',
    group1: matchedRow ? String(matchedRow[13] || '').trim() : '',
    supplierPrice: matchedRow ? toNumber_(matchedRow[15], 0) : 0,
    weight: matchedRow ? toNumber_(matchedRow[12], 0) : 0,
    volume: matchedRow ? toNumber_(matchedRow[4], 0) : 0,
    flakon: matchedRow ? String(matchedRow[8] || '').trim() : '',
    rawName: matchedRow ? String(matchedRow[7] || '').trim() : '',
    type: matchedRow ? determineType(matchedRow, flakonMap) : '',
    monthQtys: monthQtys.slice(),
    note: matchedRow ? '' : 'Не найдено совпадение в 1cData'
  };
  return item;
}

function buildPlanImportSummary_(items, months, extra) {
  extra = extra || {};
  var matchedItems = 0;
  var unmatchedItems = 0;
  var monthCounts = {};
  for (var i = 0; i < months.length; i++) monthCounts[months[i].key] = 0;

  for (i = 0; i < items.length; i++) {
    if (items[i].matched) matchedItems++;
    else unmatchedItems++;
    for (var m = 0; m < months.length; m++) {
      if (toNumber_(items[i].monthQtys[months[m].slotIndex], 0) > 0) monthCounts[months[m].key]++;
    }
  }

  return {
    processedRows: extra.processed || items.length,
    dedupedRows: extra.deduped || items.length,
    matchedItems: matchedItems,
    unmatchedItems: unmatchedItems,
    matchedByArticle: extra.matchedByArticle || 0,
    matchedByName: extra.matchedByName || 0,
    monthCounts: monthCounts
  };
}

function getStoredPlanRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PLAN_COST_CFG.SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var item = normalizeStoredPlanRow_(values[i]);
    if (!item.planName && !item.planArticleWb && !item.name) continue;
    rows.push(item);
  }
  return rows;
}

function normalizeStoredPlanRow_(row) {
  return {
    planName: String(row[PLAN_COL.PLAN_NAME] || '').trim(),
    planArticleWb: String(row[PLAN_COL.PLAN_ARTICLE_WB] || '').trim(),
    matched: String(row[PLAN_COL.MATCHED] || '').trim() === 'Да',
    matchBy: String(row[PLAN_COL.MATCH_BY] || '').trim(),
    sourceRowIndex: parseInt(row[PLAN_COL.SOURCE_ROW_INDEX], 10),
    name: String(row[PLAN_COL.NAME] || '').trim(),
    article: String(row[PLAN_COL.ARTICLE] || '').trim(),
    articleWb: String(row[PLAN_COL.ARTICLE_WB] || '').trim(),
    category: String(row[PLAN_COL.CATEGORY] || '').trim(),
    group1: String(row[PLAN_COL.GROUP1] || '').trim(),
    supplierPrice: toNumber_(row[PLAN_COL.SUPPLIER_PRICE], 0),
    weight: toNumber_(row[PLAN_COL.WEIGHT], 0),
    volume: toNumber_(row[PLAN_COL.VOLUME], 0),
    flakon: String(row[PLAN_COL.FLAKON] || '').trim(),
    rawName: String(row[PLAN_COL.RAW] || '').trim(),
    type: String(row[PLAN_COL.TYPE] || '').trim(),
    monthQtys: [
      toNumber_(row[PLAN_COL.MONTH1], 0),
      toNumber_(row[PLAN_COL.MONTH2], 0),
      toNumber_(row[PLAN_COL.MONTH3], 0)
    ],
    note: String(row[PLAN_COL.NOTE] || '').trim()
  };
}

function writePlanSheetRows_(items) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PLAN_COST_CFG.SHEET) || ss.insertSheet(PLAN_COST_CFG.SHEET);
  var rows = [PLAN_COST_CFG.HEADERS];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    rows.push([
      item.planName || '',
      item.planArticleWb || '',
      item.matched ? 'Да' : '',
      item.matchBy || '',
      item.sourceRowIndex >= 0 ? item.sourceRowIndex : '',
      item.name || '',
      item.article || '',
      item.articleWb || '',
      item.category || '',
      item.group1 || '',
      item.supplierPrice || 0,
      item.weight || 0,
      item.volume || 0,
      item.flakon || '',
      item.rawName || '',
      item.type || '',
      toNumber_(item.monthQtys[0], 0),
      toNumber_(item.monthQtys[1], 0),
      toNumber_(item.monthQtys[2], 0),
      item.note || ''
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, rows.length, PLAN_COST_CFG.HEADERS.length).setValues(rows);
  sheet.getRange(1, 1, 1, PLAN_COST_CFG.HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rows.length > 1) {
    sheet.getRange(2, PLAN_COL.SUPPLIER_PRICE + 1, rows.length - 1, 3).setNumberFormat('#,##0.00');
    sheet.getRange(2, PLAN_COL.MONTH1 + 1, rows.length - 1, 3).setNumberFormat('#,##0.00');
  }
}

function buildPlanStateStats_(items, months) {
  var matched = 0;
  var unmatched = 0;
  var monthCounts = {};
  for (var i = 0; i < months.length; i++) monthCounts[months[i].key] = 0;

  for (i = 0; i < items.length; i++) {
    if (items[i].matched) matched++;
    else unmatched++;
    for (var m = 0; m < months.length; m++) {
      if (toNumber_(items[i].monthQtys[months[m].slotIndex], 0) > 0) monthCounts[months[m].key]++;
    }
  }

  return {
    totalItems: items.length,
    matchedItems: matched,
    unmatchedItems: unmatched,
    monthCounts: monthCounts
  };
}

function resolvePlanDataRow_(item, dataRows) {
  if (!item) return null;
  var idx = parseInt(item.sourceRowIndex, 10);
  if (!isNaN(idx) && idx >= 0 && idx < dataRows.length) {
    var row = dataRows[idx];
    if (String(row[0] || '').trim() === String(item.name || '').trim()) return row;
  }

  var articleWbKey = normalizeMatchKey_(item.articleWb || item.planArticleWb);
  if (articleWbKey) {
    for (var i = 0; i < dataRows.length; i++) {
      if (normalizeMatchKey_(dataRows[i][2]) === articleWbKey) return dataRows[i];
    }
  }

  var nameKey = normalizeMatchKey_(item.name || item.planName);
  if (nameKey) {
    for (i = 0; i < dataRows.length; i++) {
      if (normalizeMatchKey_(dataRows[i][0]) === nameKey) return dataRows[i];
    }
  }

  return null;
}

function buildPlanResultRow_(planItem, matchedRow, calcResult, planQty, isChild, sourceLabel) {
  return {
    name: String((matchedRow && matchedRow[0]) || planItem.name || planItem.planName || '').trim(),
    article: String((matchedRow && matchedRow[1]) || planItem.article || '').trim(),
    articleWb: String((matchedRow && matchedRow[2]) || planItem.articleWb || planItem.planArticleWb || '').trim(),
    category: String((matchedRow && matchedRow[3]) || planItem.category || '').trim(),
    group1: String((matchedRow && matchedRow[13]) || planItem.group1 || '').trim(),
    rawName: calcResult.rawName || String((matchedRow && matchedRow[7]) || planItem.rawName || '').trim(),
    flakonName: calcResult.flakonName || String((matchedRow && matchedRow[8]) || planItem.flakon || '').trim(),
    type: calcResult.type || planItem.type || '',
    matched: true,
    source: sourceLabel || 'Расчёт',
    isComponent: !!isChild,
    planQty: round2_(planQty),
    unitTotal: round2_(calcResult.total),
    rawPlan: round2_(calcResult.raw * planQty),
    taxDutyPlan: round2_(calcResult.taxDuty * planQty),
    deliveryPlan: round2_(calcResult.delivery * planQty),
    rawFlPlan: round2_(calcResult.rawFl * planQty),
    deliveryFlPlan: round2_(calcResult.deliveryFl * planQty),
    taxDutyFlPlan: round2_(calcResult.taxDutyFl * planQty),
    labelPlan: round2_(calcResult.label * planQty),
    totalFlPlan: round2_(calcResult.totalFl * planQty),
    totalPlan: round2_(calcResult.total * planQty)
  };
}

function buildPlanBundleChildren_(planItem, matchedRow, planQty, params, flakonMap, bundleContext) {
  var bundleName = String(matchedRow[0] || planItem.name || '').trim();
  var activeRows = getBundleActiveRows_(bundleName, bundleContext);
  var children = [];

  for (var i = 0; i < activeRows.length; i++) {
    var componentDef = activeRows[i];
    var componentName = String(componentDef.component || '').trim();
    var componentQty = toNumber_(componentDef.quantity, 1);
    var childPlanQty = round2_(planQty * componentQty);
    if (!componentName || childPlanQty <= 0) continue;

    var componentRow = bundleContext.dataRowMap[normalizeMatchKey_(componentName)] || null;
    if (componentRow) {
      var childCalc = calculateOne(componentRow, params, flakonMap, null, null, bundleContext, {});
      var childRow = buildPlanResultRow_({
        planName: componentName,
        articleWb: String(componentRow[2] || '').trim()
      }, componentRow, childCalc, childPlanQty, true, 'Компонент набора');
      childRow.name = componentName;
      childRow.source = 'Компонент набора';
      children.push(childRow);
      continue;
    }

    var manualCost = componentDef.manualCost === '' ? 0 : round2_(toNumber_(componentDef.manualCost, 0));
    var manualTotal = round2_(manualCost * childPlanQty);
    children.push({
      name: componentName,
      article: '',
      articleWb: '',
      category: '',
      group1: '',
      rawName: '',
      flakonName: '',
      type: 'Сырьё',
      matched: false,
      source: manualCost > 0 ? 'Ручная стоимость' : 'Нет данных',
      isComponent: true,
      planQty: childPlanQty,
      unitTotal: manualCost,
      rawPlan: manualTotal,
      taxDutyPlan: 0,
      deliveryPlan: 0,
      rawFlPlan: 0,
      deliveryFlPlan: 0,
      taxDutyFlPlan: 0,
      labelPlan: 0,
      totalFlPlan: 0,
      totalPlan: manualTotal
    });
  }

  return children;
}

function applyPlanChildrenTotals_(row) {
  var children = row && row._children ? row._children : [];
  if (!children.length) return row;

  var totals = {
    rawPlan: 0,
    taxDutyPlan: 0,
    deliveryPlan: 0,
    rawFlPlan: 0,
    deliveryFlPlan: 0,
    taxDutyFlPlan: 0,
    labelPlan: 0,
    totalFlPlan: 0,
    totalPlan: 0
  };

  for (var i = 0; i < children.length; i++) {
    totals.rawPlan += toNumber_(children[i].rawPlan, 0);
    totals.taxDutyPlan += toNumber_(children[i].taxDutyPlan, 0);
    totals.deliveryPlan += toNumber_(children[i].deliveryPlan, 0);
    totals.rawFlPlan += toNumber_(children[i].rawFlPlan, 0);
    totals.deliveryFlPlan += toNumber_(children[i].deliveryFlPlan, 0);
    totals.taxDutyFlPlan += toNumber_(children[i].taxDutyFlPlan, 0);
    totals.labelPlan += toNumber_(children[i].labelPlan, 0);
    totals.totalFlPlan += toNumber_(children[i].totalFlPlan, 0);
    totals.totalPlan += toNumber_(children[i].totalPlan, 0);
  }

  row.rawPlan = round2_(totals.rawPlan);
  row.taxDutyPlan = round2_(totals.taxDutyPlan);
  row.deliveryPlan = round2_(totals.deliveryPlan);
  row.rawFlPlan = round2_(totals.rawFlPlan);
  row.deliveryFlPlan = round2_(totals.deliveryFlPlan);
  row.taxDutyFlPlan = round2_(totals.taxDutyFlPlan);
  row.labelPlan = round2_(totals.labelPlan);
  row.totalFlPlan = round2_(totals.totalFlPlan);
  row.totalPlan = round2_(totals.totalPlan);
  row.unitTotal = row.planQty > 0 ? round2_(row.totalPlan / row.planQty) : row.unitTotal;
  return row;
}

function buildPlanCalculationStats_(results, unresolvedItems) {
  var stats = {
    totalBudget: 0,
    totalPlanQty: 0,
    unmatchedItems: 0,
    unmatchedList: [],
    types: { raw: 0, finished: 0, flakons: 0, sets: 0 },
    rawPlan: 0,
    taxDutyPlan: 0,
    deliveryPlan: 0,
    rawFlPlan: 0,
    deliveryFlPlan: 0,
    taxDutyFlPlan: 0,
    labelPlan: 0,
    totalFlPlan: 0
  };

  for (var i = 0; i < results.length; i++) {
    var row = results[i];
    stats.totalBudget += toNumber_(row.totalPlan, 0);
    stats.totalPlanQty += toNumber_(row.planQty, 0);
    stats.rawPlan += toNumber_(row.rawPlan, 0);
    stats.taxDutyPlan += toNumber_(row.taxDutyPlan, 0);
    stats.deliveryPlan += toNumber_(row.deliveryPlan, 0);
    stats.rawFlPlan += toNumber_(row.rawFlPlan, 0);
    stats.deliveryFlPlan += toNumber_(row.deliveryFlPlan, 0);
    stats.taxDutyFlPlan += toNumber_(row.taxDutyFlPlan, 0);
    stats.labelPlan += toNumber_(row.labelPlan, 0);
    stats.totalFlPlan += toNumber_(row.totalFlPlan, 0);

    if (row.type === 'Сырьё') stats.types.raw++;
    else if (row.type === 'Готовый товар') stats.types.finished++;
    else if (row.type === 'Флакон') stats.types.flakons++;
    else stats.types.sets++;
  }

  unresolvedItems = unresolvedItems || [];
  for (i = 0; i < unresolvedItems.length; i++) {
    stats.unmatchedList.push(formatPlanUnmatchedItemLabel_(unresolvedItems[i]));
  }
  stats.unmatchedItems = stats.unmatchedList.length;

  stats.totalBudget = round2_(stats.totalBudget);
  stats.totalPlanQty = round2_(stats.totalPlanQty);
  stats.rawPlan = round2_(stats.rawPlan);
  stats.taxDutyPlan = round2_(stats.taxDutyPlan);
  stats.deliveryPlan = round2_(stats.deliveryPlan);
  stats.rawFlPlan = round2_(stats.rawFlPlan);
  stats.deliveryFlPlan = round2_(stats.deliveryFlPlan);
  stats.taxDutyFlPlan = round2_(stats.taxDutyFlPlan);
  stats.labelPlan = round2_(stats.labelPlan);
  stats.totalFlPlan = round2_(stats.totalFlPlan);

  return stats;
}

function formatPlanUnmatchedItemLabel_(item) {
  item = item || {};
  var name = String(item.planName || item.name || '').trim();
  var articleWb = String(item.planArticleWb || item.articleWb || '').trim();
  if (name && articleWb) return name + ' (' + articleWb + ')';
  return name || articleWb || 'Без названия';
}
