/**
 * DATA SERVICE — чтение/запись данных из Google Sheets
 */

var CFG = {
  DATA: '1cData',
  FLAKONS: 'Флаконы',
  BASIS: 'basis',
  RESULTS: 'Результаты',
  // Структура 1cData: A-N (14 колонок из файла) + O,P,Q,R (расчётные)
  BASE_COLS: 14,      // A-N
  COST_1C_COL: 14,    // O — Себестоимость 1С
  PRICE_COL: 15,      // P — Цена поставщика
  NDS_COL: 16,        // Q — НДС
  TAX_COL: 17,        // R — Пошлина
  TOTAL_COLS: 18      // A-R всего
};

// Жёсткие заголовки расчётных колонок
var CALC_HEADERS = ['Себестоимость 1С', 'Цена поставщика', 'НДС', 'Пошлина'];

// ============================================================
// ИМПОРТ НОМЕНКЛАТУРЫ (базовый файл)
// ============================================================

/**
 * Принимает массив данных (заголовки + строки) из клиента,
 * записывает на лист 1cData, добавляет колонки НДС и Пошлина.
 * @param {Array[]} data — [headers[], row1[], row2[], ...]
 * @param {number} ndsDefault — значение НДС по умолчанию (0.22)
 * @param {number} taxDefault — значение Пошлины по умолчанию (0.065)
 */
function importNomenclature(data, ndsDefault, taxDefault) {
  if (!data || data.length < 2) return { success: false, message: 'Файл пустой или нет данных' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);

  // Снять защиту если есть
  if (sheet) {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CFG.DATA);
  }

  var headers = data[0].slice(0, CFG.BASE_COLS);
  // Дополнить до 14 если меньше
  while (headers.length < CFG.BASE_COLS) headers.push('');
  // Добавить расчётные заголовки
  headers = headers.concat(CALC_HEADERS);

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i].slice(0, CFG.BASE_COLS);
    while (row.length < CFG.BASE_COLS) row.push('');
    row.push('', '', ndsDefault || 0.22, taxDefault || 0.065); // O,P,Q,R
    rows.push(row);
  }

  var allData = [headers].concat(rows);
  sheet.getRange(1, 1, allData.length, CFG.TOTAL_COLS).setValues(allData);

  // Форматирование
  sheet.getRange(1, 1, 1, CFG.TOTAL_COLS).setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  // Формат НДС и Пошлины как проценты
  if (rows.length > 0) {
    sheet.getRange(2, CFG.NDS_COL + 1, rows.length, 2).setNumberFormat('0.0%');
  }

  // Защита (warning only)
  var protection = sheet.protect().setDescription('1cData — данные импортированы');
  protection.setWarningOnly(true);

  return { success: true, message: 'Импортировано ' + rows.length + ' позиций', count: rows.length };
}

// ============================================================
// ИМПОРТ СЕБЕСТОИМОСТИ 1С (маппинг по номенклатуре)
// ============================================================

/**
 * @param {Object[]} mappedData — [{ name: "...", value: 123.45 }, ...]
 */
function importCost1C(mappedData) {
  return _importByName(mappedData, CFG.COST_1C_COL, 'Себестоимость 1С');
}

// ============================================================
// ИМПОРТ ЦЕНЫ ПОСТАВЩИКА (маппинг по номенклатуре)
// ============================================================

/**
 * @param {Object[]} mappedData — [{ name: "...", value: 123.45 }, ...]
 */
function importSupplierPrice(mappedData) {
  return _importByName(mappedData, CFG.PRICE_COL, 'Цена поставщика');
}

/**
 * Общая функция: заполняет колонку targetCol по совпадению номенклатуры
 */
function _importByName(mappedData, targetCol, label) {
  if (!mappedData || mappedData.length === 0) return { success: false, message: 'Нет данных для импорта' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Лист 1cData пуст. Сначала импортируйте номенклатуру.' };

  // Снять защиту
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

  var data = sheet.getDataRange().getValues();

  // Построить map из импортируемых данных
  var importMap = {};
  for (var i = 0; i < mappedData.length; i++) {
    var key = String(mappedData[i].name || '').trim().toLowerCase();
    if (key) importMap[key] = mappedData[i].value;
  }

  // Найти и заполнить
  var matched = 0;
  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][0] || '').trim().toLowerCase();
    if (name && importMap.hasOwnProperty(name)) {
      data[r][targetCol] = importMap[name];
      matched++;
    }
  }

  // Записать обратно
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Восстановить защиту
  var protection = sheet.protect().setDescription('1cData — данные импортированы');
  protection.setWarningOnly(true);

  return {
    success: true,
    message: label + ': найдено ' + matched + ' совпадений из ' + mappedData.length + ' строк файла',
    matched: matched,
    total: mappedData.length
  };
}

// ============================================================
// ОБНОВЛЕНИЕ НДС/ПОШЛИНЫ ПО СТРОКАМ
// ============================================================

function updateNdsTax(updates) {
  // updates = [{ row: 0-based index, nds: 0.22, tax: 0.065 }, ...]
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet) return { success: false };

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

  for (var i = 0; i < updates.length; i++) {
    var u = updates[i];
    var sheetRow = u.row + 2; // +1 для header, +1 для 1-based
    if (u.nds !== undefined) sheet.getRange(sheetRow, CFG.NDS_COL + 1).setValue(u.nds);
    if (u.tax !== undefined) sheet.getRange(sheetRow, CFG.TAX_COL + 1).setValue(u.tax);
  }

  var protection = sheet.protect().setDescription('1cData — данные импортированы');
  protection.setWarningOnly(true);

  return { success: true };
}

// Массовое обновление НДС/Пошлины для всех строк
function updateAllNdsTax(nds, tax) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 2) return { success: false };

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

  var numRows = sheet.getLastRow() - 1;
  var ndsRange = sheet.getRange(2, CFG.NDS_COL + 1, numRows, 1);
  var taxRange = sheet.getRange(2, CFG.TAX_COL + 1, numRows, 1);

  var ndsVals = [];
  var taxVals = [];
  for (var i = 0; i < numRows; i++) {
    ndsVals.push([nds]);
    taxVals.push([tax]);
  }
  ndsRange.setValues(ndsVals);
  taxRange.setValues(taxVals);

  var protection = sheet.protect().setDescription('1cData — данные импортированы');
  protection.setWarningOnly(true);

  return { success: true, count: numRows };
}

// ============================================================
// ЧТЕНИЕ ДАННЫХ
// ============================================================

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 1) return { headers: [], rows: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };

  return { headers: data[0], rows: data.slice(1) };
}

function getDataHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.DATA);
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ============================================================
// ПАРАМЕТРЫ
// ============================================================

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
  sheet.getRange(1, 1, 1, 4).setValues([['USD', 'RMB', 'Логистика', 'Комиссия']]).setFontWeight('bold');
  sheet.getRange(2, 1, 1, 4).setValues([[p.usd, p.rmb, p.log, p.com]]);
  return { success: true };
}

// ============================================================
// ФЛАКОНЫ
// ============================================================

function getFlakonList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var flSheet = ss.getSheetByName(CFG.FLAKONS);

  if (flSheet && flSheet.getLastRow() > 1) {
    var flData = flSheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < flData.length; i++) {
      result.push({
        name: flData[i][0] || '',
        volume: flData[i][1] || 0,
        price: flData[i][2] || 0,
        nds: flData[i][3] || 0,
        delivery: flData[i][4] || 0,
        label: flData[i][5] || 0
      });
    }
    return result;
  }

  // Собрать уникальные из 1cData
  var dataSheet = ss.getSheetByName(CFG.DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) return [];

  var data = dataSheet.getDataRange().getValues();
  var flakonMap = {};
  for (var i = 1; i < data.length; i++) {
    var fName = String(data[i][8]).trim(); // Колонка I — Тара (флакон)
    if (fName && fName !== '' && fName !== 'undefined' && !flakonMap[fName]) {
      flakonMap[fName] = {
        name: fName,
        volume: data[i][4] || 0,
        price: 0, nds: 0, delivery: 0, label: 0
      };
    }
  }
  return Object.values(flakonMap);
}

function saveFlakonData(flakons) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.FLAKONS) || ss.insertSheet(CFG.FLAKONS);
  sheet.clear();

  var headers = ['Флакон', 'Объём', 'Цена', 'НДС', 'Доставка', 'Этикетка'];
  var rows = [headers];

  for (var i = 0; i < flakons.length; i++) {
    var f = flakons[i];
    rows.push([f.name, f.volume, f.price, f.nds, f.delivery, f.label]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return { success: true, count: flakons.length };
}

// ============================================================
// РЕЗУЛЬТАТЫ
// ============================================================

function saveResults(results) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CFG.RESULTS + ' ' + Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
  var sheet = ss.insertSheet(sheetName);

  var headers = ['Номенклатура', 'Тип', 'Сырьё', 'Нал+Пош', 'Доставка', 'Флакон', 'Нал.Фл', 'Дост.Фл', 'Этикетка', 'ИТОГО', 'Себес 1С', 'Разница', 'Разница %'];
  var rows = [headers];

  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    rows.push([r.name, r.type, r.raw, r.taxDuty, r.delivery, r.flakon, r.flakonTax, r.flakonDelivery, r.label, r.total, r.cost1C, r.diff, r.diffPct]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86c8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  if (rows.length > 1) {
    sheet.getRange(2, 3, rows.length - 1, 9).setNumberFormat('#,##0.00');
    sheet.getRange(2, 13, rows.length - 1, 1).setNumberFormat('0.0%');
  }

  return { success: true, sheetName: sheetName, count: results.length };
}
