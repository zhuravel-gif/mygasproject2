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
  return {
    mappings: {
      nomenclature: getSavedImportMapping_('nomenclature'),
      cost: getSavedImportMapping_('cost'),
      supplier: getSavedImportMapping_('supplier')
    },
    rateDefaults: getRateDefaults_()
  };
}

function importBaseNomenclature(payload) {
  if (!payload || !payload.rows || payload.rows.length === 0) {
    return { success: false, message: 'Нет строк для импорта номенклатуры.' };
  }

  var defaults = payload.defaults || {};
  var ndsDefault = coerceNumber_(defaults.nds);
  var taxDefault = coerceNumber_(defaults.tax);
  if (ndsDefault === '') ndsDefault = 0.22;
  if (taxDefault === '') taxDefault = 0.065;

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
  setRateDefaults_({ nds: ndsDefault, tax: taxDefault });

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

  setRateDefaults_({ nds: ndsValue, tax: taxValue });

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
    return { usd: 92, rmb: 12.8, log: 4.5, com: 1.05 };
  }

  var data = sheet.getRange(2, 1, 1, 4).getValues()[0];
  return {
    usd: data[0] || 92,
    rmb: data[1] || 12.8,
    log: data[2] || 4.5,
    com: data[3] || 1.05
  };
}

function saveParams(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.BASIS) || ss.insertSheet(CFG.BASIS);
  sheet.clear();
  sheet.getRange(1, 1, 1, 4)
    .setValues([['USD', 'RMB', 'Логистика', 'Комиссия']])
    .setFontWeight('bold');
  sheet.getRange(2, 1, 1, 4)
    .setValues([[
      coerceNumber_(p.usd) || 92,
      coerceNumber_(p.rmb) || 12.8,
      coerceNumber_(p.log) || 4.5,
      coerceNumber_(p.com) || 1.05
    ]]);
  return { success: true };
}

function getFlakonList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var flSheet = ss.getSheetByName(CFG.FLAKONS);
  var savedMap = {};
  var savedList = [];

  if (flSheet && flSheet.getLastRow() > 1) {
    var flData = flSheet.getDataRange().getValues();
    for (var i = 1; i < flData.length; i++) {
      var savedItem = {
        name: String(flData[i][0] || '').trim(),
        volume: flData[i][1] || 0,
        price: flData[i][2] || 0,
        nds: flData[i][3] || 0,
        delivery: flData[i][4] || 0,
        label: flData[i][5] || 0
      };
      if (!savedItem.name) continue;
      savedMap[savedItem.name] = savedItem;
      savedList.push(savedItem);
    }
  }

  var dataSheet = ss.getSheetByName(CFG.DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) return savedList;

  var data = dataSheet.getDataRange().getValues();
  var result = [];
  var seen = {};

  for (var j = 1; j < data.length; j++) {
    var flakonName = String(data[j][8] || '').trim();
    if (flakonName && !seen[flakonName]) {
      seen[flakonName] = true;
      if (savedMap[flakonName]) {
        result.push({
          name: flakonName,
          volume: savedMap[flakonName].volume || 0,
          price: savedMap[flakonName].price || 0,
          nds: savedMap[flakonName].nds || 0,
          delivery: savedMap[flakonName].delivery || 0,
          label: savedMap[flakonName].label || 0
        });
      } else {
        result.push({
          name: flakonName,
          volume: data[j][4] || 0,
          price: 0,
          nds: 0,
          delivery: 0,
          label: 0
        });
      }
    }
  }

  return result;
}

function saveFlakonData(flakons) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.FLAKONS) || ss.insertSheet(CFG.FLAKONS);
  sheet.clear();

  var headers = ['Флакон', 'Объем', 'Цена', 'НДС', 'Доставка', 'Этикетка'];
  var rows = [headers];

  for (var i = 0; i < flakons.length; i++) {
    var item = flakons[i];
    rows.push([
      item.name || '',
      coerceNumber_(item.volume) || 0,
      coerceNumber_(item.price) || 0,
      coerceNumber_(item.nds) || 0,
      coerceNumber_(item.delivery) || 0,
      coerceNumber_(item.label) || 0
    ]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return { success: true, count: flakons.length };
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
    'Флакон',
    'Нал.Фл',
    'Дост.Фл',
    'Этикетка',
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
      resultItem.flakon,
      resultItem.flakonTax,
      resultItem.flakonDelivery,
      resultItem.label,
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
    sheet.getRange(2, 3, rows.length - 1, 9).setNumberFormat('#,##0.00');
    sheet.getRange(2, 13, rows.length - 1, 1).setNumberFormat('0.0%');
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
