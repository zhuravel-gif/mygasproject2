/**
 * CALC ENGINE — единый движок расчёта себестоимости
 */

// Индексы колонок в 1cData
var COL = {
  NAME: 0,     // A — Номенклатура
  ART: 1,      // B — Артикул
  ART_WB: 2,   // C — Артикул ВБ
  CAT: 3,      // D — Категория товаров
  VOL: 4,      // E — Объем тары
  GRP2: 5,     // F — Товарная группа 2
  GRP3: 6,     // G — Товарная группа 3
  RAW: 7,      // H — Основное сырье
  FLAKON: 8,   // I — Тара (флакон)
  SET_QTY: 9,  // J — Кол-во лаков в наборе
  ART_MP: 10,  // K — Артикул МП
  IS_SET: 11,  // L — Это набор
  WEIGHT: 12,  // M — Вес
  GRP1: 13,    // N — Товарная группа 1
  COST_1C: 14, // O — Себестоимость 1С
  SUPP: 15,    // P — Поставщик
  PRICE: 16,   // Q — Цена поставщика
  PRICE5: 17,  // R — Цена за 5 л
  PRICE10: 18, // S — Цена за 10 л и более
  NDS: 19,     // T — НДС
  TAX: 20      // U — Пошлина
};

// Маппинг буквы колонки → индекс
var PRICE_COL_MAP = { 'Q': 16, 'R': 17, 'S': 18 };

/**
 * Определить тип товара
 */
function determineType(row) {
  var isSet = String(row[COL.IS_SET] || '').trim();
  var hasRaw = row[COL.RAW] && String(row[COL.RAW]).trim() !== '';

  if (isSet === 'Да') return 'Наборы';
  if (!hasRaw) return 'Готовый товар';
  return 'Сырьё';
}

/**
 * Построить карту флаконов из массива
 */
function buildFlakonMap(flakons) {
  var map = {};
  for (var i = 0; i < flakons.length; i++) {
    var f = flakons[i];
    map[f.name] = { price: f.price || 0, nds: f.nds || 0, delivery: f.delivery || 0, label: f.label || 0 };
  }
  return map;
}

/**
 * Расчёт одной позиции
 * @param {Array} row — строка данных из 1cData
 * @param {Object} params — { usd, rmb, log, com, priceCol, ndsRate, taxRate }
 * @param {Object} flakonMap — карта флаконов
 * @returns {Object} результат расчёта
 */
function calculateOne(row, params, flakonMap) {
  var type = determineType(row);
  var name = String(row[COL.NAME] || '').trim();
  var cost1C = parseFloat(row[COL.COST_1C]) || 0;

  var result = {
    name: name,
    type: type,
    raw: 0,
    taxDuty: 0,
    delivery: 0,
    flakon: 0,
    flakonTax: 0,
    flakonDelivery: 0,
    label: 0,
    total: 0,
    cost1C: cost1C,
    diff: 0,
    diffPct: 0
  };

  if (type === 'Наборы') {
    result.diff = cost1C > 0 ? -cost1C : 0;
    result.diffPct = 0;
    return result;
  }

  var priceIdx = PRICE_COL_MAP[params.priceCol] || COL.PRICE;
  var priceVal = parseFloat(row[priceIdx]) || 0;
  var ndsRate = parseFloat(params.ndsRate) || parseFloat(row[COL.NDS]) || 0.22;
  var taxRate = parseFloat(params.taxRate) || parseFloat(row[COL.TAX]) || 0.065;
  var usd = parseFloat(params.usd) || 92;
  var rmb = parseFloat(params.rmb) || 12.8;
  var log = parseFloat(params.log) || 4.5;
  var com = parseFloat(params.com) || 1.05;

  if (type === 'Готовый товар') {
    var weight = parseFloat(row[COL.WEIGHT]) || 0;
    result.raw = priceVal * rmb * com;
    result.delivery = log * weight * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;
  } else {
    // Сырьё
    var vol = parseFloat(row[COL.VOL]) || 0;
    result.raw = (priceVal * rmb * com / 1000) * vol;
    result.delivery = log * vol / 1000 * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;

    // Флакон
    var flakonName = String(row[COL.FLAKON] || '').trim();
    var fl = flakonMap[flakonName] || { price: 0, nds: 0, delivery: 0, label: 0 };
    result.flakon = parseFloat(fl.price) || 0;
    result.flakonTax = parseFloat(fl.nds) || 0;
    result.flakonDelivery = parseFloat(fl.delivery) || 0;
    result.label = parseFloat(fl.label) || 0;
  }

  result.raw = Math.round(result.raw * 100) / 100;
  result.taxDuty = Math.round(result.taxDuty * 100) / 100;
  result.delivery = Math.round(result.delivery * 100) / 100;
  result.flakon = Math.round(result.flakon * 100) / 100;
  result.flakonTax = Math.round(result.flakonTax * 100) / 100;
  result.flakonDelivery = Math.round(result.flakonDelivery * 100) / 100;
  result.label = Math.round(result.label * 100) / 100;

  result.total = result.raw + result.taxDuty + result.delivery + result.flakon + result.flakonTax + result.flakonDelivery + result.label;
  result.total = Math.round(result.total * 100) / 100;

  result.diff = cost1C > 0 ? Math.round((result.total - cost1C) * 100) / 100 : 0;
  result.diffPct = cost1C > 0 ? Math.round((result.total - cost1C) / cost1C * 10000) / 10000 : 0;

  return result;
}

/**
 * Расчёт всех позиций
 */
function calculateAll(params) {
  var dataObj = getData();
  if (!dataObj.rows || dataObj.rows.length === 0) {
    return { success: false, message: 'Нет данных в 1cData. Сначала выполните экспорт.', results: [] };
  }

  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);

  var results = [];
  for (var i = 0; i < dataObj.rows.length; i++) {
    results.push(calculateOne(dataObj.rows[i], params, flakonMap));
  }

  // Сохранить параметры
  saveParams(params);

  return { success: true, results: results, count: results.length };
}

/**
 * Расчёт одной позиции по индексу
 */
function calculateByIndex(index, params) {
  var dataObj = getData();
  if (!dataObj.rows || index < 0 || index >= dataObj.rows.length) {
    return { success: false, message: 'Позиция не найдена' };
  }
  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  return calculateOne(dataObj.rows[index], params, flakonMap);
}

/**
 * Ручной расчёт с произвольными данными
 */
function calculateManual(input) {
  var params = {
    usd: input.usd, rmb: input.rmb, log: input.log, com: input.com,
    priceCol: 'Q', ndsRate: input.ndsRate, taxRate: input.taxRate
  };

  // Сформировать виртуальную строку
  var row = new Array(21).fill('');
  row[COL.NAME] = input.name || 'Ручной расчёт';
  row[COL.VOL] = input.volume || 0;
  row[COL.RAW] = input.type === 'Сырьё' ? 'manual' : '';
  row[COL.FLAKON] = input.flakonName || '';
  row[COL.IS_SET] = input.type === 'Наборы' ? 'Да' : 'Нет';
  row[COL.WEIGHT] = input.weight || 0;
  row[COL.PRICE] = input.price || 0;
  row[COL.NDS] = input.ndsRate || 0.22;
  row[COL.TAX] = input.taxRate || 0.065;
  row[COL.COST_1C] = input.cost1C || 0;

  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);

  return calculateOne(row, params, flakonMap);
}

/**
 * Детальная проверка расчёта — формулы с подстановкой
 */
function getVerification(index, params) {
  var dataObj = getData();
  if (!dataObj.rows || index < 0 || index >= dataObj.rows.length) {
    return { success: false, message: 'Позиция не найдена' };
  }

  var row = dataObj.rows[index];
  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  var result = calculateOne(row, params, flakonMap);

  var type = determineType(row);
  var priceIdx = PRICE_COL_MAP[params.priceCol] || COL.PRICE;
  var priceVal = parseFloat(row[priceIdx]) || 0;
  var ndsRate = parseFloat(params.ndsRate) || 0.22;
  var taxRate = parseFloat(params.taxRate) || 0.065;
  var usd = parseFloat(params.usd) || 92;
  var rmb = parseFloat(params.rmb) || 12.8;
  var log = parseFloat(params.log) || 4.5;
  var com = parseFloat(params.com) || 1.05;
  var vol = parseFloat(row[COL.VOL]) || 0;
  var weight = parseFloat(row[COL.WEIGHT]) || 0;
  var flakonName = String(row[COL.FLAKON] || '').trim();
  var fl = flakonMap[flakonName] || { price: 0, nds: 0, delivery: 0, label: 0 };
  var cost1C = parseFloat(row[COL.COST_1C]) || 0;

  var steps = [];

  steps.push({
    title: 'Определение типа',
    formula: 'L (Набор) = "' + String(row[COL.IS_SET] || '') + '", H (Сырьё) = "' + String(row[COL.RAW] || '') + '"',
    result: type
  });

  if (type === 'Наборы') {
    steps.push({ title: 'Наборы не рассчитываются', formula: '', result: '0.00' });
  } else if (type === 'Готовый товар') {
    steps.push({
      title: 'Сырьё',
      formula: 'Цена × RMB × Ком = ' + priceVal + ' × ' + rmb + ' × ' + com,
      result: result.raw.toFixed(2)
    });
    steps.push({
      title: 'Доставка',
      formula: 'Лог × Вес × 1.15 × USD × Ком = ' + log + ' × ' + weight + ' × 1.15 × ' + usd + ' × ' + com,
      result: result.delivery.toFixed(2)
    });
    steps.push({
      title: 'Нал+Пош',
      formula: '(Сырьё + Доставка + Сырьё×Пошлина) × НДС + Сырьё×Пошлина\n= (' + result.raw.toFixed(2) + ' + ' + result.delivery.toFixed(2) + ' + ' + result.raw.toFixed(2) + '×' + taxRate + ') × ' + ndsRate + ' + ' + result.raw.toFixed(2) + '×' + taxRate,
      result: result.taxDuty.toFixed(2)
    });
  } else {
    // Сырьё
    steps.push({
      title: 'Сырьё',
      formula: '(Цена × RMB × Ком / 1000) × Объём = (' + priceVal + ' × ' + rmb + ' × ' + com + ' / 1000) × ' + vol,
      result: result.raw.toFixed(2)
    });
    steps.push({
      title: 'Доставка',
      formula: 'Лог × Объём/1000 × 1.15 × USD × Ком = ' + log + ' × ' + (vol / 1000).toFixed(4) + ' × 1.15 × ' + usd + ' × ' + com,
      result: result.delivery.toFixed(2)
    });
    steps.push({
      title: 'Нал+Пош',
      formula: '(Сырьё + Доставка + Сырьё×Пошлина) × НДС + Сырьё×Пошлина\n= (' + result.raw.toFixed(2) + ' + ' + result.delivery.toFixed(2) + ' + ' + result.raw.toFixed(2) + '×' + taxRate + ') × ' + ndsRate + ' + ' + result.raw.toFixed(2) + '×' + taxRate,
      result: result.taxDuty.toFixed(2)
    });
    steps.push({
      title: 'Флакон',
      formula: 'Флакон "' + flakonName + '": Цена = ' + fl.price,
      result: result.flakon.toFixed(2)
    });
    steps.push({
      title: 'НДС флакона',
      formula: 'Из таблицы флаконов',
      result: result.flakonTax.toFixed(2)
    });
    steps.push({
      title: 'Доставка флакона',
      formula: 'Из таблицы флаконов',
      result: result.flakonDelivery.toFixed(2)
    });
    steps.push({
      title: 'Этикетка',
      formula: 'Из таблицы флаконов',
      result: result.label.toFixed(2)
    });
  }

  steps.push({
    title: 'ИТОГО',
    formula: 'Сумма всех компонентов',
    result: result.total.toFixed(2)
  });

  if (cost1C > 0) {
    steps.push({
      title: 'Сравнение с 1С',
      formula: 'Себестоимость 1С = ' + cost1C.toFixed(2) + ', Разница = ' + result.diff.toFixed(2) + ' (' + (result.diffPct * 100).toFixed(1) + '%)',
      result: result.diff > 0 ? '+' + result.diff.toFixed(2) : result.diff.toFixed(2)
    });
  }

  return {
    success: true,
    name: String(row[COL.NAME] || ''),
    type: type,
    steps: steps,
    result: result,
    params: { usd: usd, rmb: rmb, log: log, com: com, ndsRate: ndsRate, taxRate: taxRate, priceCol: params.priceCol }
  };
}
