/**
 * DATA SERVICE — чтение и запись данных Google Sheets.
 */

var CFG = {
  DATA: '1cData',
  FLAKONS: 'Флаконы',
  BASIS: 'basis',
  RESULTS: 'Результаты',
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

var IMPORT_META_KEYS = {
  nomenclature: 'mapping.nomenclature',
  cost: 'mapping.cost',
  supplier: 'mapping.supplier',
  rateDefaults: 'defaults.rate'
};

function getImportSettings() {
  var params = getParams();
  return {
    mappings: {
      nomenclature: getSavedImportMapping_('nomenclature'),
      cost: getSavedImportMapping_('cost'),
      supplier: getSavedImportMapping_('supplier')
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

  for (var i = 0; i < updates.length; i++) {
    var item = updates[i];
    var rowNumber = Number(item.row) + 2;

    if (item.supplierPrice !== undefined) {
      sheet.getRange(rowNumber, CFG.PRICE_COL + 1).setValue(coerceNumber_(item.supplierPrice));
    }
    if (item.nds !== undefined) {
      sheet.getRange(rowNumber, CFG.NDS_COL + 1).setValue(normalizeRateValue_(item.nds, 0.22));
    }
    if (item.tax !== undefined) {
      sheet.getRange(rowNumber, CFG.TAX_COL + 1).setValue(normalizeRateValue_(item.tax, 0.065));
    }
  }

  formatDataSheet_(sheet, sheet.getLastRow() - 1);
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
  formatDataSheet_(sheet, numRows);
  applyWarningProtection_(sheet, '1cData — данные импортированы');

  return { success: true, count: numRows };
}

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 1) {
    return { headers: getProjectHeaders_(), rows: [] };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { headers: data[0] || getProjectHeaders_(), rows: [] };
  }

  return { headers: data[0], rows: data.slice(1) };
}

function getDataHeaders() {
  return getProjectHeaders_();
}

function getParams() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.BASIS);
  if (!sheet || sheet.getLastRow() < 2) {
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

  var data = sheet.getRange(2, 1, 1, 10).getValues()[0];
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
  return { success: true };
}

function getFlakonList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var flSheet = ss.getSheetByName(CFG.FLAKONS);
  var savedMap = {};
  var params = getParams();

  if (flSheet && flSheet.getLastRow() > 1) {
    var flData = flSheet.getDataRange().getValues();
    var flHeaders = flData[0] || [];
    for (var i = 1; i < flData.length; i++) {
      var savedName = String(flData[i][0] || '').trim();
      if (!savedName) continue;
      savedMap[savedName] = normalizeStoredFlakonRow_(flData[i], flHeaders);
    }
  }

  var dataSheet = ss.getSheetByName(CFG.DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) {
    return recalculateFlakonList_(mapValues_(savedMap), params);
  }

  var data = dataSheet.getDataRange().getValues();
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

  return { success: true, count: normalized.length, flakons: normalized };
}

function saveResults(results) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var sheetName = CFG.RESULTS + ' ' + Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
  var sheet = ss.insertSheet(sheetName);

  var headers = [
    'Номенклатура',
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
    sheet.getRange(2, 3, rows.length - 1, 10).setNumberFormat('#,##0.00');
    sheet.getRange(2, 14, rows.length - 1, 1).setNumberFormat('0.0%');
  }

  return { success: true, sheetName: sheetName, count: results.length };
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
  for (var rowIndex = 1; rowIndex < data.length; rowIndex++) {
    var nameKey = normalizeMatchKey_(data[rowIndex][0]);
    var articleKey = normalizeMatchKey_(data[rowIndex][1]);
    if (nameKey && !nameMap.hasOwnProperty(nameKey)) nameMap[nameKey] = rowIndex;
    if (articleKey && !articleMap.hasOwnProperty(articleKey)) articleMap[articleKey] = rowIndex;
  }

  var matchedByName = 0;
  var matchedByArticle = 0;
  var unmatched = [];

  for (var i = 0; i < payload.rows.length; i++) {
    var item = payload.rows[i] || {};
    var nameLookup = normalizeMatchKey_(item.name);
    var articleLookup = normalizeMatchKey_(item.article);
    var targetRow = null;

    if (nameLookup && nameMap.hasOwnProperty(nameLookup)) {
      targetRow = nameMap[nameLookup];
      matchedByName++;
    } else if (articleLookup && articleMap.hasOwnProperty(articleLookup)) {
      targetRow = articleMap[articleLookup];
      matchedByArticle++;
    } else if (unmatched.length < 25) {
      unmatched.push(item.name || item.article || ('Строка ' + (i + 1)));
    }

    if (targetRow !== null) {
      data[targetRow][targetCol] = coerceNumber_(item.value);
    }
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
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
  var sheet = getMetaSheet_();
  if (sheet.getLastRow() < 2) return '';

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return '';
}

function setMetaValue_(key, value) {
  var sheet = getMetaSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, 2).setValues([[key, value]]);
    return;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
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
