/**
 * DATA SERVICE — чтение и запись данных Google Sheets.
 */

var CFG = {
  DATA: '1cData',
  FLAKONS: 'Флаконы',
  BUNDLES: 'Наборы',
  BASIS: 'basis',
  RESULTS: 'Результаты',
  RESULTS_SNAPSHOT_PREFIX: 'Результаты',
  META: '__meta',
  BASE_FIELDS: [
    { key: 'name', label: 'Номенклатура.Наименование' },
    { key: 'article', label: 'Артикул' },
    { key: 'articleWb', label: 'Артикул ВБ' },
    { key: 'category', label: 'Категория товаров' },
    { key: 'volume', label: 'Объем тары' },
    { key: 'group2', label: 'Номенклатура.Товарная группа 2 (Общие)' },
    { key: 'group3', label: 'Номенклатура.Товарная группа 3 (Общие)' },
    { key: 'raw', label: 'Номенклатура.Основное сырье (Общие)' },
    { key: 'flakon', label: 'Номенклатура.Тара (флакон) (Общие)' },
    { key: 'setQty', label: 'Номенклатура.Количество лаков в наборе (Общие)' },
    { key: 'articleMp', label: 'Номенклатура.Артикул МП' },
    { key: 'isSet', label: 'Номенклатура.Это набор (RockNail)' },
    { key: 'weight', label: 'Номенклатура.Вес (числитель)' },
    { key: 'group1', label: 'Номенклатура.Товарная группа 1 (Общие)' }
  ],
  CALC_FIELDS: [
    { key: 'cost1C', label: 'Себестоимость 1С' },
    { key: 'supplierPrice', label: 'Цена поставщика' },
    { key: 'nds', label: 'НДС' },
    { key: 'tax', label: 'Пошлина' }
  ],
  BASE_COLS: 14,
  COST_1C_COL: 14,
  PRICE_COL: 15,
  NDS_COL: 16,
  TAX_COL: 17,
  TOTAL_COLS: 18
};

/* ── Execution-scoped sheet read cache ── */
var _sheetCache = {};

function _cachedSheetRead(sheetName) {
  if (_sheetCache[sheetName]) return _sheetCache[sheetName];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) {
    _sheetCache[sheetName] = null;
    return null;
  }
  var data = sheet.getDataRange().getValues();
  _sheetCache[sheetName] = data;
  return data;
}

function _invalidateCache(sheetName) {
  if (sheetName) {
    delete _sheetCache[sheetName];
  } else {
    _sheetCache = {};
  }
}
/* ── End cache ── */

var COST_RESULT_HEADERS = [
  'Ключ',
  'Номенклатура',
  'Артикул ВБ',
  'Артикул',
  'Флакон',
  'Категория',
  'Тип',
  'Товарная группа 1',
  'Товарная группа 2',
  'Товарная группа 3',
  'Основное сырьё',
  'Сырьё',
  'Нал+Пош',
  'Доставка',
  'Цена флакона в руб.',
  'Доставка флакона в руб.',
  'НДС+пошлина флакона',
  'Этикетка',
  'Себестоимость флакона',
  'Итого расчёт',
  'Ручной расчёт',
  'Итого ручной',
  'ИТОГО',
  'Себес 1С',
  'Разница',
  'Разница %',
  'Обновлено'
];

var COST_RESULT_COL = {
  ROW_KEY: 0,
  NAME: 1,
  ARTICLE_WB: 2,
  ARTICLE: 3,
  FLAKON: 4,
  CATEGORY: 5,
  TYPE: 6,
  GROUP1: 7,
  GROUP2: 8,
  GROUP3: 9,
  RAW_NAME: 10,
  RAW: 11,
  TAX_DUTY: 12,
  DELIVERY: 13,
  RAW_FL: 14,
  DELIVERY_FL: 15,
  TAX_DUTY_FL: 16,
  LABEL: 17,
  TOTAL_FL: 18,
  CALC_TOTAL: 19,
  MANUAL_MODE: 20,
  MANUAL_TOTAL: 21,
  TOTAL: 22,
  COST_1C: 23,
  DIFF: 24,
  DIFF_PCT: 25,
  UPDATED_AT: 26
};

var IMPORT_META_KEYS = {
  nomenclature: 'mapping.nomenclature',
  cost: 'mapping.cost',
  supplier: 'mapping.supplier',
  bundles: 'mapping.bundles',
  rateDefaults: 'defaults.rate'
};

function getImportSettings() {
  var params = getParams();
  return {
    mappings: {
      nomenclature: getSavedImportMapping_('nomenclature'),
      cost: getSavedImportMapping_('cost'),
      supplier: getSavedImportMapping_('supplier'),
      bundles: getSavedImportMapping_('bundles')
    },
    rateDefaults: {
      nds: params.importNds,
      tax: params.importTax
    }
  };
}

function importBaseNomenclature(payload) {
  if (!payload || !payload.rows || payload.rows.length === 0) {
    return { success: false, message: 'Нет строк для импорта номенклатуры.' };
  }

  var defaults = payload.defaults || {};
  var params = getParams();
  var ndsDefault = coerceNumber_(defaults.nds);
  var taxDefault = coerceNumber_(defaults.tax);
  if (ndsDefault === '') ndsDefault = params.importNds;
  if (taxDefault === '') taxDefault = params.importTax;

  var rows = [];
  for (var i = 0; i < payload.rows.length; i++) {
    var item = payload.rows[i];
    if (!hasMeaningfulValue_(item)) continue;
    rows.push(buildBaseRow_(item, ndsDefault, taxDefault));
  }

  if (rows.length === 0) {
    return { success: false, message: 'После фильтрации не осталось данных для импорта.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (sheet) {
    removeSheetProtections_(sheet);
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CFG.DATA);
  }

  var allData = [getProjectHeaders_()].concat(rows);
  sheet.getRange(1, 1, allData.length, CFG.TOTAL_COLS).setValues(allData);
  formatDataSheet_(sheet, rows.length);
  applyWarningProtection_(sheet, '1cData — данные импортированы');

  _invalidateCache(CFG.DATA);
  saveImportMapping_('nomenclature', payload.mapping || null);

  return {
    success: true,
    message: 'Базовая номенклатура импортирована: ' + rows.length + ' строк.',
    count: rows.length
  };
}

function importCurrentCost(payload) {
  return importMappedValues_(payload, CFG.COST_1C_COL, 'Себестоимость 1С', 'cost');
}

function importSupplierPriceData(payload) {
  return importMappedValues_(payload, CFG.PRICE_COL, 'Цена поставщика', 'supplier');
}

function updateImportRows(updates) {
  if (!updates || updates.length === 0) {
    return { success: false, message: 'Нет изменений для сохранения.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Лист 1cData не заполнен.' };
  }

  removeSheetProtections_(sheet);

  var numRows = sheet.getLastRow() - 1;
  var colStart = CFG.PRICE_COL + 1;
  var colCount = CFG.TAX_COL - CFG.PRICE_COL + 1;
  var range = sheet.getRange(2, colStart, numRows, colCount);
  var values = range.getValues();

  for (var i = 0; i < updates.length; i++) {
    var item = updates[i];
    var rowIdx = Number(item.row);
    if (rowIdx < 0 || rowIdx >= numRows) continue;

    if (item.supplierPrice !== undefined) {
      values[rowIdx][0] = coerceNumber_(item.supplierPrice);
    }
    if (item.nds !== undefined) {
      values[rowIdx][1] = normalizeRateValue_(item.nds, 0.22);
    }
    if (item.tax !== undefined) {
      values[rowIdx][2] = normalizeRateValue_(item.tax, 0.065);
    }
  }

  range.setValues(values);
  _invalidateCache(CFG.DATA);
  SpreadsheetApp.flush();
  formatDataSheet_(sheet, numRows);
  applyWarningProtection_(sheet, '1cData — данные импортированы');

  return { success: true, count: updates.length };
}

function updateNdsTax(updates) {
  return updateImportRows(updates);
}

function updateAllNdsTax(nds, tax) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Лист 1cData не заполнен.' };
  }

  removeSheetProtections_(sheet);

  var numRows = sheet.getLastRow() - 1;
  var ndsValue = normalizeRateValue_(nds, 0.22);
  var taxValue = normalizeRateValue_(tax, 0.065);
  var ndsVals = [];
  var taxVals = [];
  for (var i = 0; i < numRows; i++) {
    ndsVals.push([ndsValue]);
    taxVals.push([taxValue]);
  }

  sheet.getRange(2, CFG.NDS_COL + 1, numRows, 1).setValues(ndsVals);
  sheet.getRange(2, CFG.TAX_COL + 1, numRows, 1).setValues(taxVals);
  _invalidateCache(CFG.DATA);
  formatDataSheet_(sheet, numRows);
  applyWarningProtection_(sheet, '1cData — данные импортированы');

  return { success: true, count: numRows };
}

function getData() {
  var data = _cachedSheetRead(CFG.DATA);
  if (!data || data.length < 1) {
    return { headers: getProjectHeaders_(), rows: [], typeColIndex: -1 };
  }
  if (data.length < 2) {
    return { headers: data[0] || getProjectHeaders_(), rows: [], typeColIndex: -1 };
  }

  var headers = data[0].slice();
  var rows = data.slice(1);
  var typeHeader = 'Тип товара';
  var typeColIndex = headers.length;
  headers.push(typeHeader);

  var flakonMap = null;
  if (typeof buildFlakonMap === 'function' && typeof getFlakonList === 'function') {
    flakonMap = buildFlakonMap(getFlakonList());
  }

  var enrichedRows = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i].slice();
    row.push(getDataRowType_(row, flakonMap || {}));
    enrichedRows.push(row);
  }

  return { headers: headers, rows: enrichedRows, typeColIndex: typeColIndex };
}

function getDataHeaders() {
  return getProjectHeaders_();
}

function getDataRowType_(row, flakonMap) {
  if (!row || !row.length) return '';
  if (typeof determineType === 'function') return determineType(row, flakonMap || {});

  var rawColIndex = (typeof COL !== 'undefined' && typeof COL.RAW === 'number') ? COL.RAW : 7;
  var setColIndex = (typeof COL !== 'undefined' && typeof COL.IS_SET === 'number') ? COL.IS_SET : 11;
  var isSet = String(row[setColIndex] || '').trim();
  var hasRaw = String(row[rawColIndex] || '').trim() !== '';
  if (isSet === 'Да') return 'Наборы';
  return hasRaw ? 'Сырьё' : 'Готовый товар';
}

function getParams() {
  var cached = _cachedSheetRead(CFG.BASIS);
  if (!cached || cached.length < 2) {
    return {
      usd: 92,
      rmb: 12.8,
      log: 4.5,
      com: 1.05,
      logFl: 4.5,
      comFl: 1.05,
      importNds: 0.22,
      importTax: 0.065,
      flakonNds: 0.22,
      flakonTax: 0.065
    };
  }

  var data = cached[1];
  return {
    usd: hasValue_(data[0]) ? data[0] : 92,
    rmb: hasValue_(data[1]) ? data[1] : 12.8,
    log: hasValue_(data[2]) ? data[2] : 4.5,
    com: hasValue_(data[3]) ? data[3] : 1.05,
    logFl: hasValue_(data[4]) ? data[4] : (hasValue_(data[2]) ? data[2] : 4.5),
    comFl: hasValue_(data[5]) ? data[5] : (hasValue_(data[3]) ? data[3] : 1.05),
    importNds: hasValue_(data[6]) ? data[6] : 0.22,
    importTax: hasValue_(data[7]) ? data[7] : 0.065,
    flakonNds: hasValue_(data[8]) ? data[8] : 0.22,
    flakonTax: hasValue_(data[9]) ? data[9] : 0.065
  };
}

function saveParams(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.BASIS) || ss.insertSheet(CFG.BASIS);
  sheet.clear();
  sheet.getRange(1, 1, 1, 10)
    .setValues([['USD', 'RMB', 'Логистика', 'Комиссия', 'Логистика_fl', 'Комиссия_fl', 'НДС_импорт', 'Пошлина_импорт', 'НДС_флаконы', 'Пошлина_флаконы']])
    .setFontWeight('bold');
  sheet.getRange(2, 1, 1, 10)
    .setValues([[
      hasValue_(coerceNumber_(p.usd)) ? coerceNumber_(p.usd) : 92,
      hasValue_(coerceNumber_(p.rmb)) ? coerceNumber_(p.rmb) : 12.8,
      hasValue_(coerceNumber_(p.log)) ? coerceNumber_(p.log) : 4.5,
      hasValue_(coerceNumber_(p.com)) ? coerceNumber_(p.com) : 1.05,
      hasValue_(coerceNumber_(p.logFl)) ? coerceNumber_(p.logFl) : (hasValue_(coerceNumber_(p.log)) ? coerceNumber_(p.log) : 4.5),
      hasValue_(coerceNumber_(p.comFl)) ? coerceNumber_(p.comFl) : (hasValue_(coerceNumber_(p.com)) ? coerceNumber_(p.com) : 1.05),
      normalizeRateValue_(p.importNds, 0.22),
      normalizeRateValue_(p.importTax, 0.065),
      normalizeRateValue_(p.flakonNds, 0.22),
      normalizeRateValue_(p.flakonTax, 0.065)
    ]]);
  _invalidateCache(CFG.BASIS);
  return { success: true };
}

function getFlakonList() {
  var savedMap = {};
  var params = getParams();

  var flData = _cachedSheetRead(CFG.FLAKONS);
  if (flData && flData.length > 1) {
    var flHeaders = flData[0] || [];
    for (var i = 1; i < flData.length; i++) {
      var savedName = String(flData[i][0] || '').trim();
      if (!savedName) continue;
      savedMap[savedName] = normalizeStoredFlakonRow_(flData[i], flHeaders);
    }
  }

  var data = _cachedSheetRead(CFG.DATA);
  if (!data || data.length < 2) {
    return recalculateFlakonList_(mapValues_(savedMap), params);
  }
  var flakonMap = {};
  var productByName = {};

  for (var k = 1; k < data.length; k++) {
    var productNameKey = normalizeMatchKey_(data[k][0]);
    if (productNameKey && !productByName.hasOwnProperty(productNameKey)) {
      productByName[productNameKey] = data[k];
    }
  }

  for (var j = 1; j < data.length; j++) {
    var flakonName = String(data[j][8] || '').trim();
    if (!flakonName || flakonMap[flakonName]) continue;
    var matchedRow = productByName[normalizeMatchKey_(flakonName)] || null;

    var imported = {
      name: flakonName,
      volume: matchedRow ? matchedRow[4] : '',
      weight: matchedRow ? matchedRow[12] : '',
      supplierPrice: matchedRow ? matchedRow[15] : '',
      nds: params.flakonNds,
      tax: params.flakonTax,
      label: 0
    };
    flakonMap[flakonName] = mergeFlakonRows_(imported, savedMap[flakonName]);
  }

  return recalculateFlakonList_(mapValues_(flakonMap), params);
}

function recalculateFlakonData(flakons) {
  var normalized = recalculateFlakonList_(flakons || [], getParams());
  return { success: true, count: normalized.length, flakons: normalized };
}

function saveFlakonData(flakons) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.FLAKONS) || ss.insertSheet(CFG.FLAKONS);
  var normalized = recalculateFlakonList_(flakons || [], getParams());
  var headers = ['Флакон', 'Объём', 'Вес', 'Цена поставщика', 'НДС', 'Пошлина', 'Этикетка', 'Цена флакона в руб.', 'Доставка в руб.', 'НДС+пошлина', 'Себестоимость флакона'];
  var rows = [headers];

  for (var i = 0; i < normalized.length; i++) {
    var item = normalized[i];
    rows.push([
      item.name || '',
      coerceNumber_(item.volume) || 0,
      coerceNumber_(item.weight) || 0,
      coerceNumber_(item.supplierPrice) || 0,
      hasValue_(coerceNumber_(item.nds)) ? coerceNumber_(item.nds) : params.flakonNds,
      hasValue_(coerceNumber_(item.tax)) ? coerceNumber_(item.tax) : params.flakonTax,
      coerceNumber_(item.label) || 0,
      coerceNumber_(item.rawFl) || 0,
      coerceNumber_(item.deliveryFl) || 0,
      coerceNumber_(item.taxDutyFl) || 0,
      coerceNumber_(item.totalFl) || 0
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rows.length > 1) {
    sheet.getRange(2, 2, rows.length - 1, 10).setNumberFormat('#,##0.00');
    sheet.getRange(2, 5, rows.length - 1, 2).setNumberFormat('0.0%');
  }

  _invalidateCache(CFG.FLAKONS);
  return { success: true, count: normalized.length, flakons: normalized };
}

function importBundleCompositions(payload) {
  if (!payload || !payload.rows || payload.rows.length === 0) {
    return { success: false, message: 'Нет строк для импорта комплектаций наборов.' };
  }

  var existingRows = getStoredBundleRows_();
  var manualCostMap = {};
  var activeSpecMap = {};
  var i;

  for (i = 0; i < existingRows.length; i++) {
    var existing = existingRows[i];
    if (hasValue_(existing.manualCost)) {
      manualCostMap[normalizeMatchKey_(existing.component)] = toNumber_(existing.manualCost, 0);
    }
    if (existing.active && existing.bundle && existing.specification) {
      activeSpecMap[existing.bundle] = existing.specification;
    }
  }

  var dedupeMap = {};
  var specOrder = {};
  for (i = 0; i < payload.rows.length; i++) {
    var item = payload.rows[i] || {};
    var component = String(item.component || '').trim();
    var bundle = String(item.bundle || '').trim();
    var specification = String(item.specification || '').trim();
    if (!component || !bundle || !specification) continue;

    var quantity = toNumber_(item.quantity, 1);
    if (!hasValue_(quantity)) quantity = 1;

    var dedupeKey = [bundle, specification, component].join('||');
    if (!dedupeMap[dedupeKey]) {
      dedupeMap[dedupeKey] = {
        component: component,
        bundle: bundle,
        specification: specification,
        quantity: 0,
        active: false,
        manualCost: hasValue_(manualCostMap[normalizeMatchKey_(component)]) ? manualCostMap[normalizeMatchKey_(component)] : ''
      };
      if (!specOrder[bundle]) specOrder[bundle] = [];
      if (specOrder[bundle].indexOf(specification) === -1) specOrder[bundle].push(specification);
    }
    dedupeMap[dedupeKey].quantity += toNumber_(quantity, 0);
  }

  var rows = mapValues_(dedupeMap);
  for (i = 0; i < rows.length; i++) {
    var row = rows[i];
    var activeSpec = activeSpecMap[row.bundle] || (specOrder[row.bundle] && specOrder[row.bundle][0]) || '';
    row.active = row.specification === activeSpec;
  }

  writeBundleSheetRows_(rows);
  refreshBundleSheetComputed_();
  saveImportMapping_('bundles', payload.mapping || null);

  var bundleData = getBundleData();
  return {
    success: true,
    message: 'Комплектации наборов импортированы: ' + rows.length + ' строк.',
    count: rows.length,
    stats: bundleData.stats
  };
}

function getBundleData() {
  if (typeof buildBundleUiState_ === 'function') {
    return buildBundleUiState_();
  }

  return {
    rows: getStoredBundleRows_(),
    bundles: [],
    manualItems: [],
    stats: getBundleStats()
  };
}

function getBundleStats() {
  var totalSets = 0;
  var linked = {};
  var knownSetNames = {};
  var data = getData();
  var i;

  for (i = 0; i < data.rows.length; i++) {
    if (String(data.rows[i][11] || '').trim() === 'Да') {
      totalSets++;
      knownSetNames[String(data.rows[i][0] || '').trim()] = true;
    }
  }

  var bundleRows = getStoredBundleRows_();
  for (i = 0; i < bundleRows.length; i++) {
    if (bundleRows[i].bundle && knownSetNames[bundleRows[i].bundle]) linked[bundleRows[i].bundle] = true;
  }

  return {
    totalSets: totalSets,
    loadedBundles: Object.keys(linked).length,
    compositionRows: bundleRows.length
  };
}

function saveBundleData(payload) {
  payload = payload || {};
  var existingRows = getStoredBundleRows_();
  if (!existingRows.length) {
    return { success: false, message: 'Сначала импортируйте комплектации наборов.' };
  }

  var activeSpecs = payload.activeSpecs || {};
  var manualCostsRaw = payload.manualCosts || {};
  var manualCosts = {};
  var i;

  for (var manualKey in manualCostsRaw) {
    if (!manualCostsRaw.hasOwnProperty(manualKey)) continue;
    manualCosts[normalizeMatchKey_(manualKey)] = manualCostsRaw[manualKey];
  }

  for (i = 0; i < existingRows.length; i++) {
    var row = existingRows[i];
    var activeSpec = activeSpecs.hasOwnProperty(row.bundle) ? String(activeSpecs[row.bundle] || '').trim() : '';
    if (activeSpec) {
      row.active = row.specification === activeSpec;
    }

    var componentKey = normalizeMatchKey_(row.component);
    if (manualCosts.hasOwnProperty(componentKey)) {
      row.manualCost = manualCosts[componentKey] === '' ? '' : toNumber_(manualCosts[componentKey], 0);
    }
  }

  writeBundleSheetRows_(existingRows);
  refreshBundleSheetComputed_();
  var data = getBundleData();
  data.success = true;
  data.message = 'Наборы сохранены.';
  return data;
}

function getBundleHeaders_() {
  return [
    'Компонент',
    'Набор',
    'Спецификация',
    'Количество',
    'Активная',
    'Себес 1С',
    'Себестоимость расчёт',
    'Ручная стоимость',
    'Используемая стоимость',
    'Источник стоимости'
  ];
}

function getStoredBundleRows_() {
  var data = _cachedSheetRead(CFG.BUNDLES);
  if (!data || data.length < 2) return [];
  var headers = data[0] || [];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var item = normalizeStoredBundleRow_(data[i], headers);
    if (!item.bundle || !item.component || !item.specification) continue;
    rows.push(item);
  }
  return rows;
}

function normalizeStoredBundleRow_(row) {
  return {
    component: String(row[0] || '').trim(),
    bundle: String(row[1] || '').trim(),
    specification: String(row[2] || '').trim(),
    quantity: toNumber_(row[3], 1),
    active: String(row[4] || '').trim() === 'Да',
    cost1C: toNumber_(row[5], 0),
    calcCost: toNumber_(row[6], 0),
    manualCost: row[7] === '' || row[7] === null || row[7] === undefined ? '' : toNumber_(row[7], 0),
    usedCost: toNumber_(row[8], 0),
    source: String(row[9] || '').trim()
  };
}

function writeBundleSheetRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.BUNDLES) || ss.insertSheet(CFG.BUNDLES);
  var headers = getBundleHeaders_();
  var values = [headers];
  var i;

  for (i = 0; i < rows.length; i++) {
    values.push([
      rows[i].component || '',
      rows[i].bundle || '',
      rows[i].specification || '',
      toNumber_(rows[i].quantity, 1),
      rows[i].active ? 'Да' : '',
      hasValue_(rows[i].cost1C) ? toNumber_(rows[i].cost1C, 0) : '',
      hasValue_(rows[i].calcCost) ? toNumber_(rows[i].calcCost, 0) : '',
      rows[i].manualCost === '' ? '' : toNumber_(rows[i].manualCost, 0),
      hasValue_(rows[i].usedCost) ? toNumber_(rows[i].usedCost, 0) : '',
      rows[i].source || ''
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (values.length > 1) {
    sheet.getRange(2, 4, values.length - 1, 5).setNumberFormat('#,##0.00');
  }
  _invalidateCache(CFG.BUNDLES);
}

function refreshBundleSheetComputed_() {
  if (typeof buildBundleUiState_ !== 'function') return;

  var data = buildBundleUiState_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.BUNDLES) || ss.insertSheet(CFG.BUNDLES);
  var headers = getBundleHeaders_();
  var rows = [headers];

  for (var i = 0; i < data.rows.length; i++) {
    var item = data.rows[i];
    rows.push([
      item.component || '',
      item.bundle || '',
      item.specification || '',
      toNumber_(item.quantity, 1),
      item.active ? 'Да' : '',
      toNumber_(item.cost1C, 0),
      toNumber_(item.calcCost, 0),
      item.manualCost === '' ? '' : toNumber_(item.manualCost, 0),
      toNumber_(item.usedCost, 0),
      item.source || ''
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rows.length > 1) {
    sheet.getRange(2, 4, rows.length - 1, 5).setNumberFormat('#,##0.00');
  }
  _invalidateCache(CFG.BUNDLES);
}

function saveResults(results) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var sheetName = CFG.RESULTS_SNAPSHOT_PREFIX + ' ' + Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
  var sheet = ss.insertSheet(sheetName);

  var headers = [
    'Номенклатура',
    'Флакон',
    'Категория',
    'Тип',
    'Сырьё',
    'Нал+Пош',
    'Доставка',
    'Цена флакона в руб.',
    'Доставка флакона в руб.',
    'НДС+пошлина флакона',
    'Этикетка',
    'Себестоимость флакона',
    'ИТОГО',
    'Себес 1С',
    'Разница',
    'Разница %'
  ];
  var rows = [headers];

  for (var k = 0; k < results.length; k++) {
    var resultItem = results[k];
    rows.push([
      resultItem.name,
      resultItem.flakonName || '',
      resultItem.category || '',
      resultItem.type,
      resultItem.raw,
      resultItem.taxDuty,
      resultItem.delivery,
      resultItem.rawFl,
      resultItem.deliveryFl,
      resultItem.taxDutyFl,
      resultItem.label,
      resultItem.totalFl,
      resultItem.total,
      resultItem.cost1C,
      resultItem.diff,
      resultItem.diffPct
    ]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rows.length > 1) {
    sheet.getRange(2, 5, rows.length - 1, 11).setNumberFormat('#,##0.00');
    sheet.getRange(2, 16, rows.length - 1, 1).setNumberFormat('0.0%');
  }

  return { success: true, sheetName: sheetName, count: results.length };
}

function getCostState() {
  try {
    var rows = getStoredCostResults_();
    if (!rows.length) {
      var data = getData();
      if (!data.rows || !data.rows.length) {
        return { success: true, results: [], count: 0, stats: { raw: 0, finished: 0, flakons: 0, sets: 0, totalCost: 0, total1C: 0 } };
      }
      return recalculateAndStoreCostResults(getParams());
    }

    return {
      success: true,
      results: rows,
      count: rows.length,
      stats: buildStoredCostStats_(rows),
      updatedAt: rows[0] && rows[0].updatedAt ? rows[0].updatedAt : ''
    };
  } catch (err) {
    return { success: false, message: 'Не удалось загрузить расчёт себестоимости: ' + err.message, diagnostic: err.stack || '' };
  }
}

function recalculateAndStoreCostResults(params) {
  try {
    var calc = calculateAll(params || getParams());
    if (!calc || !calc.success) return calc || { success: false, message: 'Не удалось пересчитать себестоимость.' };

    writeStoredCostResults_(calc.results || []);
    var rows = getStoredCostResults_();
    return {
      success: true,
      results: rows,
      count: rows.length,
      stats: buildStoredCostStats_(rows),
      updatedAt: rows[0] && rows[0].updatedAt ? rows[0].updatedAt : ''
    };
  } catch (err) {
    return { success: false, message: 'Не удалось пересчитать себестоимость: ' + err.message, diagnostic: err.stack || '' };
  }
}

function recalculateCostResults(payload) {
  try {
    payload = payload || {};
    var calc = calculateAll(payload.params || getParams(), payload.currentRows || []);
    if (!calc || !calc.success) return calc || { success: false, message: 'Не удалось пересчитать себестоимость.' };

    return {
      success: true,
      results: calc.results || [],
      count: (calc.results || []).length,
      stats: buildStoredCostStats_(calc.results || [])
    };
  } catch (err) {
    return { success: false, message: 'Не удалось пересчитать себестоимость: ' + err.message, diagnostic: err.stack || '' };
  }
}

function saveCostState(payload) {
  try {
    payload = payload || {};
    var rows = payload.results || [];
    writeStoredCostResults_(rows, { preserveExistingManuals: false });
    return getCostState();
  } catch (err) {
    return { success: false, message: 'Не удалось сохранить изменения себестоимости: ' + err.message, diagnostic: err.stack || '' };
  }
}

function saveCostManualOverride(payload) {
  try {
    payload = payload || {};
    var rowKey = String(payload.rowKey || '').trim();
    if (!rowKey) return { success: false, message: 'Не указан ключ строки для ручного расчёта.' };

    var rows = getStoredCostResults_();
    var targetIndex = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].rowKey || '') === rowKey) {
        targetIndex = i;
        break;
      }
    }
    if (targetIndex < 0) return { success: false, message: 'Строка себестоимости не найдена.' };

    var row = rows[targetIndex];
    var manualMode = !!payload.manualMode;
    row.manualMode = manualMode;
    row.manualTotal = manualMode ? round2_(toNumber_(payload.manualTotal, row.calcTotal || row.total || 0)) : '';
    row.total = manualMode ? row.manualTotal : round2_(toNumber_(row.calcTotal, row.total || 0));
    row.diff = row.cost1C > 0 ? round2_(row.total - row.cost1C) : 0;
    row.diffPct = row.cost1C > 0 ? round4_((row.total - row.cost1C) / row.cost1C) : 0;
    row.updatedAt = new Date().toISOString();

    writeStoredCostResults_(rows);
    return { success: true, row: row };
  } catch (err) {
    return { success: false, message: 'Не удалось сохранить ручную себестоимость: ' + err.message, diagnostic: err.stack || '' };
  }
}

function getStoredCostResults_() {
  var values = _cachedSheetRead(CFG.RESULTS);
  if (!values || values.length < 2) return [];
  if (!values.length || !isCurrentCostResultsSchema_(values[0])) return [];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var item = normalizeStoredCostResultRow_(values[i]);
    if (!item.name) continue;
    rows.push(item);
  }
  return rows;
}

function getCostOverrideMap_() {
  var rows = getStoredCostResults_();
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row.manualMode || !hasValue_(row.manualTotal)) continue;
    var keys = getCostLookupKeys_(row.name, row.articleWb, row.article);
    for (var j = 0; j < keys.length; j++) map[keys[j]] = { manualMode: true, manualTotal: row.manualTotal };
  }
  return map;
}

function writeStoredCostResults_(results, options) {
  options = options || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.RESULTS);
  if (!sheet) sheet = ss.insertSheet(CFG.RESULTS);

  var existingManualMap = options.preserveExistingManuals === false ? {} : getStoredCostManualMap_();
  var nowIso = new Date().toISOString();
  var rows = [COST_RESULT_HEADERS];

  for (var i = 0; i < results.length; i++) {
    var item = results[i] || {};
    var rowKey = buildCostResultKey_(item.name, item.articleWb, item.article);
    var existingManual = existingManualMap[rowKey] || null;
    var manualMode = item.manualMode === true || (existingManual && existingManual.manualMode);
    var calcTotal = round2_(toNumber_(item.calcTotal, item.total));
    var manualTotal = manualMode
      ? round2_(toNumber_(item.manualTotal, existingManual ? existingManual.manualTotal : calcTotal))
      : '';
    var total = manualMode ? manualTotal : round2_(toNumber_(item.total, calcTotal));
    var cost1C = round2_(toNumber_(item.cost1C, 0));
    var diff = cost1C > 0 ? round2_(total - cost1C) : 0;
    var diffPct = cost1C > 0 ? round4_((total - cost1C) / cost1C) : 0;

    rows.push([
      rowKey,
      item.name || '',
      item.articleWb || '',
      item.article || '',
      item.flakonName || '',
      item.category || '',
      item.type || '',
      item.group1 || '',
      item.group2 || '',
      item.group3 || '',
      item.rawName || '',
      round2_(toNumber_(item.raw, 0)),
      round2_(toNumber_(item.taxDuty, 0)),
      round2_(toNumber_(item.delivery, 0)),
      round2_(toNumber_(item.rawFl, 0)),
      round2_(toNumber_(item.deliveryFl, 0)),
      round2_(toNumber_(item.taxDutyFl, 0)),
      round2_(toNumber_(item.label, 0)),
      round2_(toNumber_(item.totalFl, 0)),
      calcTotal,
      manualMode ? 'Да' : '',
      manualMode ? manualTotal : '',
      total,
      cost1C,
      diff,
      diffPct,
      nowIso
    ]);
  }

  sheet.clear();
  sheet.getRange(1, 1, rows.length, COST_RESULT_HEADERS.length).setValues(rows);
  formatStoredCostSheet_(sheet, rows.length - 1);
  _invalidateCache(CFG.RESULTS);
}

function getStoredCostManualMap_() {
  var rows = getStoredCostResults_();
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    map[rows[i].rowKey] = { manualMode: !!rows[i].manualMode, manualTotal: rows[i].manualTotal };
  }
  return map;
}

function normalizeStoredCostResultRow_(row) {
  return {
    rowKey: String(row[COST_RESULT_COL.ROW_KEY] || '').trim(),
    name: String(row[COST_RESULT_COL.NAME] || '').trim(),
    articleWb: String(row[COST_RESULT_COL.ARTICLE_WB] || '').trim(),
    article: String(row[COST_RESULT_COL.ARTICLE] || '').trim(),
    flakonName: String(row[COST_RESULT_COL.FLAKON] || '').trim(),
    category: String(row[COST_RESULT_COL.CATEGORY] || '').trim(),
    type: String(row[COST_RESULT_COL.TYPE] || '').trim(),
    group1: String(row[COST_RESULT_COL.GROUP1] || '').trim(),
    group2: String(row[COST_RESULT_COL.GROUP2] || '').trim(),
    group3: String(row[COST_RESULT_COL.GROUP3] || '').trim(),
    rawName: String(row[COST_RESULT_COL.RAW_NAME] || '').trim(),
    raw: round2_(toNumber_(row[COST_RESULT_COL.RAW], 0)),
    taxDuty: round2_(toNumber_(row[COST_RESULT_COL.TAX_DUTY], 0)),
    delivery: round2_(toNumber_(row[COST_RESULT_COL.DELIVERY], 0)),
    rawFl: round2_(toNumber_(row[COST_RESULT_COL.RAW_FL], 0)),
    deliveryFl: round2_(toNumber_(row[COST_RESULT_COL.DELIVERY_FL], 0)),
    taxDutyFl: round2_(toNumber_(row[COST_RESULT_COL.TAX_DUTY_FL], 0)),
    label: round2_(toNumber_(row[COST_RESULT_COL.LABEL], 0)),
    totalFl: round2_(toNumber_(row[COST_RESULT_COL.TOTAL_FL], 0)),
    calcTotal: round2_(toNumber_(row[COST_RESULT_COL.CALC_TOTAL], 0)),
    manualMode: String(row[COST_RESULT_COL.MANUAL_MODE] || '').trim() === 'Да',
    manualTotal: row[COST_RESULT_COL.MANUAL_TOTAL] === '' || row[COST_RESULT_COL.MANUAL_TOTAL] === null || row[COST_RESULT_COL.MANUAL_TOTAL] === undefined
      ? ''
      : round2_(toNumber_(row[COST_RESULT_COL.MANUAL_TOTAL], 0)),
    total: round2_(toNumber_(row[COST_RESULT_COL.TOTAL], 0)),
    cost1C: round2_(toNumber_(row[COST_RESULT_COL.COST_1C], 0)),
    diff: round2_(toNumber_(row[COST_RESULT_COL.DIFF], 0)),
    diffPct: round4_(toNumber_(row[COST_RESULT_COL.DIFF_PCT], 0)),
    updatedAt: String(row[COST_RESULT_COL.UPDATED_AT] || '').trim()
  };
}

function buildStoredCostStats_(results) {
  var stats = { raw: 0, finished: 0, flakons: 0, sets: 0, totalCost: 0, total1C: 0 };
  for (var i = 0; i < results.length; i++) {
    var item = results[i];
    if (item.type === 'Сырьё') stats.raw++;
    else if (item.type === 'Готовый товар') stats.finished++;
    else if (item.type === 'Флакон') stats.flakons++;
    else if (item.type === 'Наборы') stats.sets++;
    stats.totalCost += toNumber_(item.total, 0);
    stats.total1C += toNumber_(item.cost1C, 0);
  }
  stats.totalCost = round2_(stats.totalCost);
  stats.total1C = round2_(stats.total1C);
  return stats;
}

function formatStoredCostSheet_(sheet, rowCount) {
  sheet.getRange(1, 1, 1, COST_RESULT_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rowCount > 0) {
    sheet.getRange(2, COST_RESULT_COL.RAW + 1, rowCount, 14).setNumberFormat('#,##0.00');
    sheet.getRange(2, COST_RESULT_COL.DIFF_PCT + 1, rowCount, 1).setNumberFormat('0.0%');
  }
}

function getCostLookupKeys_(name, articleWb, article) {
  var keys = [];
  var wbKey = normalizeMatchKey_(articleWb);
  var articleKey = normalizeMatchKey_(article);
  var nameKey = normalizeMatchKey_(name);
  if (wbKey) keys.push('wb:' + wbKey);
  if (articleKey) keys.push('art:' + articleKey);
  if (nameKey) keys.push('name:' + nameKey);
  return keys;
}

function buildCostResultKey_(name, articleWb, article) {
  return [
    normalizeMatchKey_(name),
    normalizeMatchKey_(articleWb),
    normalizeMatchKey_(article)
  ].join('|');
}

function isCurrentCostResultsSchema_(headers) {
  if (!headers || headers.length < COST_RESULT_HEADERS.length) return false;
  return String(headers[0] || '').trim() === COST_RESULT_HEADERS[0] &&
    String(headers[COST_RESULT_COL.TOTAL] || '').trim() === COST_RESULT_HEADERS[COST_RESULT_COL.TOTAL];
}

function importNomenclature(data, ndsDefault, taxDefault) {
  if (!data || data.length < 2) {
    return { success: false, message: 'Файл пустой или в нем нет строк.' };
  }

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push(arrayToNamedRow_(data[i]));
  }

  return importBaseNomenclature({
    rows: rows,
    defaults: { nds: ndsDefault, tax: taxDefault }
  });
}

function importCost1C(mappedData) {
  return importCurrentCost({
    rows: mappedData || []
  });
}

function importSupplierPrice(mappedData) {
  return importSupplierPriceData({
    rows: mappedData || []
  });
}

function importMappedValues_(payload, targetCol, label, mappingType) {
  if (!payload || !payload.rows || payload.rows.length === 0) {
    return { success: false, message: 'Нет строк для импорта.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Сначала загрузите базовую номенклатуру.' };
  }

  removeSheetProtections_(sheet);

  var data = sheet.getDataRange().getValues();
  var nameMap = {};
  var articleMap = {};
  var supplierNameMap = {};
  var rawColIndex = (typeof COL !== 'undefined' && typeof COL.RAW === 'number') ? COL.RAW : 7;
  var nameColIndex = (typeof COL !== 'undefined' && typeof COL.NAME === 'number') ? COL.NAME : 0;
  var flakonMap = null;
  if (mappingType === 'supplier' && typeof buildFlakonMap === 'function' && typeof getFlakonList === 'function') {
    flakonMap = buildFlakonMap(getFlakonList());
  }
  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var nameKey = normalizeMatchKey_(data[rowIndex][nameColIndex]);
    var articleKey = normalizeMatchKey_(data[rowIndex][1]);
    var rowType = '';
    if (mappingType === 'supplier') {
      if (typeof determineType === 'function') {
        rowType = determineType(data[rowIndex], flakonMap || {});
      } else {
        rowType = String(data[rowIndex][rawColIndex] || '').trim() ? 'Сырьё' : '';
      }
    }
    var supplierKey = rowType === 'Сырьё'
      ? normalizeMatchKey_(data[rowIndex][rawColIndex])
      : nameKey;
    if (nameKey && !nameMap.hasOwnProperty(nameKey)) nameMap[nameKey] = rowIndex;
    if (articleKey && !articleMap.hasOwnProperty(articleKey)) articleMap[articleKey] = rowIndex;
    if (mappingType === 'supplier' && supplierKey) {
      if (!supplierNameMap.hasOwnProperty(supplierKey)) supplierNameMap[supplierKey] = [];
      supplierNameMap[supplierKey].push(rowIndex);
    }
  }

  var matchedByName = 0;
  var matchedByArticle = 0;
  var unmatched = [];

  for (var i = 0; i < payload.rows.length; i++) {
    var item = payload.rows[i] || {};
    var nameLookup = normalizeMatchKey_(item.name);
    var articleLookup = normalizeMatchKey_(item.article);
    var targetRows = [];

    if (mappingType === 'supplier' && nameLookup && supplierNameMap.hasOwnProperty(nameLookup)) {
      targetRows = supplierNameMap[nameLookup].slice();
      matchedByName++;
    } else if (nameLookup && nameMap.hasOwnProperty(nameLookup)) {
      targetRows = [nameMap[nameLookup]];
      matchedByName++;
    } else if (articleLookup && articleMap.hasOwnProperty(articleLookup)) {
      targetRows = [articleMap[articleLookup]];
      matchedByArticle++;
    } else if (unmatched.length < 25) {
      unmatched.push(item.name || item.article || ('Строка ' + (i + 1)));
    }

    if (targetRows.length) {
      for (var targetIndex = 0; targetIndex < targetRows.length; targetIndex++) {
        data[targetRows[targetIndex]][targetCol] = coerceNumber_(item.value);
      }
    }
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  _invalidateCache(CFG.DATA);
  formatDataSheet_(sheet, data.length - 1);
  applyWarningProtection_(sheet, '1cData — данные импортированы');
  saveImportMapping_(mappingType, payload.mapping || null);

  var matchedTotal = matchedByName + matchedByArticle;
  return {
    success: true,
    message: label + ': сопоставлено ' + matchedTotal + ' из ' + payload.rows.length + ' строк.',
    matched: matchedTotal,
    matchedByName: matchedByName,
    matchedByArticle: matchedByArticle,
    unmatched: unmatched,
    total: payload.rows.length
  };
}

function buildBaseRow_(item, ndsDefault, taxDefault) {
  var row = [];
  for (var i = 0; i < CFG.BASE_FIELDS.length; i++) {
    row.push(sanitizeCell_(item[CFG.BASE_FIELDS[i].key]));
  }
  row.push('');
  row.push('');
  row.push(normalizeRateValue_(item.nds, ndsDefault));
  row.push(normalizeRateValue_(item.tax, taxDefault));
  return row;
}

function getProjectHeaders_() {
  var headers = [];
  for (var i = 0; i < CFG.BASE_FIELDS.length; i++) {
    headers.push(CFG.BASE_FIELDS[i].label);
  }
  for (var j = 0; j < CFG.CALC_FIELDS.length; j++) {
    headers.push(CFG.CALC_FIELDS[j].label);
  }
  return headers;
}

function sanitizeCell_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  return value;
}

function hasMeaningfulValue_(item) {
  if (!item) return false;
  var keys = Object.keys(item);
  for (var i = 0; i < keys.length; i++) {
    var value = item[keys[i]];
    if (value !== '' && value !== null && value !== undefined) return true;
  }
  return false;
}

function normalizeMatchKey_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function coerceNumber_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') return value;

  var normalized = String(value)
    .replace(/\s+/g, '')
    .replace(',', '.');
  var parsed = parseFloat(normalized);
  return isNaN(parsed) ? sanitizeCell_(value) : parsed;
}

function normalizeRateValue_(value, fallback) {
  var parsed = coerceNumber_(value);
  if (parsed === '') return fallback;
  return parsed;
}

function formatDataSheet_(sheet, rowCount) {
  sheet.getRange(1, 1, 1, CFG.TOTAL_COLS)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rowCount > 0) {
    sheet.getRange(2, CFG.COST_1C_COL + 1, rowCount, 2).setNumberFormat('#,##0.00');
    sheet.getRange(2, CFG.NDS_COL + 1, rowCount, 2).setNumberFormat('0.0%');
  }
}

function removeSheetProtections_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(protection) {
    protection.remove();
  });
}

function applyWarningProtection_(sheet, description) {
  var protection = sheet.protect().setDescription(description);
  protection.setWarningOnly(true);
}

function getMetaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.META);
  if (!sheet) {
    sheet = ss.insertSheet(CFG.META);
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sheet.hideSheet();
  }
  return sheet;
}

function getMetaValue_(key) {
  getMetaSheet_();
  var data = _cachedSheetRead(CFG.META);
  if (!data || data.length < 2) return '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

function setMetaValue_(key, value) {
  var sheet = getMetaSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, 2).setValues([[key, value]]);
    _invalidateCache(CFG.META);
    return;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      _invalidateCache(CFG.META);
      return;
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
  _invalidateCache(CFG.META);
}

function getSavedImportMapping_(type) {
  var key = IMPORT_META_KEYS[type];
  if (!key) return null;

  var raw = getMetaValue_(key);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function saveImportMapping_(type, mapping) {
  var key = IMPORT_META_KEYS[type];
  if (!key || !mapping) return;
  setMetaValue_(key, JSON.stringify(mapping));
}

function getRateDefaults_() {
  var raw = getMetaValue_(IMPORT_META_KEYS.rateDefaults);
  if (!raw) return { nds: 0.22, tax: 0.065 };

  try {
    var parsed = JSON.parse(raw);
    return {
      nds: normalizeRateValue_(parsed.nds, 0.22),
      tax: normalizeRateValue_(parsed.tax, 0.065)
    };
  } catch (err) {
    return { nds: 0.22, tax: 0.065 };
  }
}

function setRateDefaults_(defaults) {
  setMetaValue_(IMPORT_META_KEYS.rateDefaults, JSON.stringify({
    nds: normalizeRateValue_(defaults.nds, 0.22),
    tax: normalizeRateValue_(defaults.tax, 0.065)
  }));
}

function arrayToNamedRow_(row) {
  var result = {};
  for (var i = 0; i < CFG.BASE_FIELDS.length; i++) {
    result[CFG.BASE_FIELDS[i].key] = row[i];
  }
  if (row.length > CFG.BASE_COLS + 2) result.nds = row[CFG.NDS_COL];
  if (row.length > CFG.BASE_COLS + 3) result.tax = row[CFG.TAX_COL];
  return result;
}

function normalizeStoredFlakonRow_(row, headers) {
  var hasNewSchema = row.length >= 11;
  var isCurrentOrder = !headers || String(headers[6] || '').trim() === 'Этикетка';
  return {
    name: row[0] || '',
    volume: row[1] || 0,
    weight: hasNewSchema ? row[2] || 0 : 0,
    supplierPrice: hasNewSchema ? row[3] || 0 : 0,
    nds: hasNewSchema ? (hasValue_(row[4]) ? row[4] : '') : (hasValue_(row[3]) ? row[3] : ''),
    tax: hasNewSchema ? (hasValue_(row[5]) ? row[5] : '') : '',
    label: hasNewSchema ? (isCurrentOrder ? row[6] || 0 : row[9] || 0) : row[5] || 0
  };
}

function mergeFlakonRows_(imported, saved) {
  if (!saved) return imported;
  return {
    name: imported.name || saved.name || '',
    volume: hasValue_(imported.volume) ? imported.volume : saved.volume,
    weight: hasValue_(imported.weight) ? imported.weight : saved.weight,
    supplierPrice: hasValue_(saved.supplierPrice) ? saved.supplierPrice : imported.supplierPrice,
    nds: hasValue_(saved.nds) ? saved.nds : imported.nds,
    tax: hasValue_(saved.tax) ? saved.tax : imported.tax,
    label: hasValue_(saved.label) ? saved.label : imported.label
  };
}

function recalculateFlakonList_(flakons, params) {
  params = params || getParams();
  var result = [];
  for (var i = 0; i < flakons.length; i++) {
    var item = flakons[i] || {};
    var name = String(item.name || '').trim();
    if (!name) continue;

    var normalized = {
      name: name,
      volume: toNumber_(item.volume, 0),
      weight: toNumber_(item.weight, 0),
      supplierPrice: toNumber_(item.supplierPrice, 0),
      nds: toNumber_(item.nds, params.flakonNds),
      tax: toNumber_(item.tax, params.flakonTax),
      label: toNumber_(item.label, 0)
    };
    var metrics = calculateFlakonMetrics_(normalized, params || {});
    normalized.rawFl = metrics.rawFl;
    normalized.deliveryFl = metrics.deliveryFl;
    normalized.taxDutyFl = metrics.taxDutyFl;
    normalized.totalFl = metrics.totalFl;
    normalized.label = metrics.label;
    result.push(normalized);
  }
  return result;
}

function hasValue_(value) {
  return value !== '' && value !== null && value !== undefined;
}

function mapValues_(obj) {
  var result = [];
  var keys = Object.keys(obj || {});
  for (var i = 0; i < keys.length; i++) {
    result.push(obj[keys[i]]);
  }
  return result;
}
