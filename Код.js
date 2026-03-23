/**
 * Main entry points for Apps Script WebApp.
 */

function doGet(e) {
  var debugMode = e && e.parameter ? String(e.parameter.debug || '').trim() : '';
  if (debugMode === 'plan-state') {
    return ContentService
      .createTextOutput(JSON.stringify(buildPlanStateDebugPayload_(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('Калькулятор себестоимости')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Калькулятор')
    .addItem('Открыть калькулятор', 'openWebApp')
    .addSeparator()
    .addItem('Сбросить рабочие листы проекта', 'resetProjectSheets')
    .addToUi();
}

function openWebApp() {
  var html = HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setWidth(1280)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Калькулятор себестоимости');
}

function resetProjectSheets() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Сброс рабочих листов',
    'Будут удалены все лишние листы, включая старые листы результатов. Рабочие листы проекта будут очищены и оставлены только с заголовками. Продолжить?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var specs = getProjectSheetResetSpecs_();
  var keepNames = Object.keys(specs);

  keepNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  ss.getSheets().forEach(function(sheet) {
    if (keepNames.indexOf(sheet.getName()) === -1) {
      ss.deleteSheet(sheet);
    }
  });

  var safeActiveSheet = ss.getSheetByName('basis') || ss.getSheetByName('1cData');
  if (safeActiveSheet) ss.setActiveSheet(safeActiveSheet);

  keepNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    resetSheetProtections_(sheet);
    setupProjectSheet_(sheet, specs[name]);
  });

  var metaSheet = ss.getSheetByName('__meta');
  if (metaSheet) metaSheet.hideSheet();

  var firstSheet = ss.getSheetByName('basis') || ss.getSheetByName('1cData');
  if (firstSheet) ss.setActiveSheet(firstSheet);

  ui.alert('Готово', 'Рабочие листы проекта пересозданы. Лишние листы удалены.', ui.ButtonSet.OK);
}

function getProjectSheetResetSpecs_() {
  return {
    basis: {
      headers: ['USD', 'RMB', 'Логистика', 'Комиссия', 'Логистика_fl', 'Комиссия_fl', 'НДС_импорт', 'Пошлина_импорт', 'НДС_флаконы', 'Пошлина_флаконы'],
      hidden: false
    },
    '1cData': {
      headers: [
        'Номенклатура.Наименование',
        'Артикул',
        'Артикул ВБ',
        'Категория товаров',
        'Объем тары',
        'Номенклатура.Товарная группа 2 (Общие)',
        'Номенклатура.Товарная группа 3 (Общие)',
        'Номенклатура.Основное сырье (Общие)',
        'Номенклатура.Тара (флакон) (Общие)',
        'Номенклатура.Количество лаков в наборе (Общие)',
        'Номенклатура.Артикул МП',
        'Номенклатура.Это набор (RockNail)',
        'Номенклатура.Вес (числитель)',
        'Номенклатура.Товарная группа 1 (Общие)',
        'Себестоимость 1С',
        'Цена поставщика',
        'НДС',
        'Пошлина'
      ],
      hidden: false
    },
    'Флаконы': {
      headers: [
        'Флакон',
        'Объём',
        'Вес',
        'Цена поставщика',
        'НДС',
        'Пошлина',
        'Этикетка',
        'Цена флакона в руб.',
        'Доставка в руб.',
        'НДС+пошлина',
        'Себестоимость флакона'
      ],
      hidden: false
    },
    'Наборы': {
      headers: [
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
      ],
      hidden: false
    },
    '__meta': {
      headers: ['key', 'value'],
      hidden: true
    }
  };
}

function setupProjectSheet_(sheet, spec) {
  sheet.clear();
  sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
  sheet.getRange(1, 1, 1, spec.headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  if (spec.hidden) sheet.hideSheet();
  else sheet.showSheet();
}

function resetSheetProtections_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(protection) {
    try {
      protection.remove();
    } catch (e) {}
  });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(protection) {
    try {
      protection.remove();
    } catch (e) {}
  });
}

function buildPlanStateDebugPayload_() {
  var payload = {
    timestamp: new Date().toISOString(),
    scriptVersion: 'plan-state-debug-endpoint'
  };

  try {
    payload.state = typeof getPlanCostState === 'function'
      ? getPlanCostState()
      : { success: false, message: 'getPlanCostState недоступна.' };
  } catch (stateErr) {
    payload.state = {
      success: false,
      message: stateErr && stateErr.message ? String(stateErr.message) : 'Ошибка вызова getPlanCostState.',
      stack: stateErr && stateErr.stack ? String(stateErr.stack) : ''
    };
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('План затрат');
    payload.planSheet = sheet ? {
      exists: true,
      type: sheet.getType ? String(sheet.getType()) : 'GRID',
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      header: sheet.getLastRow() >= 1 ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0] : []
    } : { exists: false };
  } catch (sheetErr) {
    payload.planSheet = {
      exists: false,
      error: sheetErr && sheetErr.message ? String(sheetErr.message) : 'Не удалось прочитать лист "План затрат".',
      stack: sheetErr && sheetErr.stack ? String(sheetErr.stack) : ''
    };
  }

  try {
    payload.metaPreview = typeof getMetaValue_ === 'function'
      ? String(getMetaValue_('plan.importConfig') || '').slice(0, 4000)
      : 'getMetaValue_ недоступна.';
  } catch (metaErr) {
    payload.metaPreview = 'META_ERROR: ' + (metaErr && metaErr.message ? String(metaErr.message) : metaErr);
  }

  return payload;
}
