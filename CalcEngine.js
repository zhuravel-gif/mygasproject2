/**
 * CALC ENGINE — единый движок расчета себестоимости.
 */

var COL = {
  NAME: 0,
  ART: 1,
  ART_WB: 2,
  CAT: 3,
  VOL: 4,
  GRP2: 5,
  GRP3: 6,
  RAW: 7,
  FLAKON: 8,
  SET_QTY: 9,
  ART_MP: 10,
  IS_SET: 11,
  WEIGHT: 12,
  GRP1: 13,
  COST_1C: 14,
  PRICE: 15,
  NDS: 16,
  TAX: 17
};

function determineType(row) {
  var isSet = String(row[COL.IS_SET] || '').trim();
  var hasRaw = String(row[COL.RAW] || '').trim() !== '';

  if (isSet === 'Да') return 'Наборы';
  if (!hasRaw) return 'Готовый товар';
  return 'Сырьё';
}

function calculateFlakonMetrics_(item, params) {
  var supplierPrice = toNumber_(item && item.supplierPrice, 0);
  var weight = toNumber_(item && item.weight, 0);
  var ndsRate = toNumber_(item && item.nds, 0.22);
  var taxRate = toNumber_(item && item.tax, 0.065);
  var label = toNumber_(item && item.label, 0);
  var usd = toNumber_(params && params.usd, 92);
  var rmb = toNumber_(params && params.rmb, 12.8);
  var logFl = toNumber_(params && params.logFl, 4.5);
  var comFl = toNumber_(params && params.comFl, 1.05);

  var rawFl = supplierPrice * rmb * comFl;
  var deliveryFl = logFl * weight * 1.15 * usd * comFl;
  var taxDutyFl = (rawFl + deliveryFl + rawFl * taxRate) * ndsRate + rawFl * taxRate;
  var totalFl = rawFl + deliveryFl + taxDutyFl + label;

  return {
    supplierPrice: supplierPrice,
    weight: weight,
    nds: ndsRate,
    tax: taxRate,
    label: round2_(label),
    rawFl: round2_(rawFl),
    deliveryFl: round2_(deliveryFl),
    taxDutyFl: round2_(taxDutyFl),
    totalFl: round2_(totalFl)
  };
}

function buildFlakonMap(flakons) {
  var map = {};
  for (var i = 0; i < flakons.length; i++) {
    var item = flakons[i] || {};
    var name = String(item.name || '').trim();
    if (!name) continue;
    var metrics = calculateFlakonMetrics_(item, {});
    map[name] = {
      name: name,
      volume: toNumber_(item.volume, 0),
      supplierPrice: toNumber_(item.supplierPrice, metrics.supplierPrice),
      weight: toNumber_(item.weight, metrics.weight),
      nds: toNumber_(item.nds, metrics.nds),
      tax: toNumber_(item.tax, metrics.tax),
      label: toNumber_(item.label, metrics.label),
      rawFl: toNumber_(item.rawFl, metrics.rawFl),
      deliveryFl: toNumber_(item.deliveryFl, metrics.deliveryFl),
      taxDutyFl: toNumber_(item.taxDutyFl, metrics.taxDutyFl),
      totalFl: toNumber_(item.totalFl, metrics.totalFl)
    };
  }
  return map;
}

function calculateOne(row, params, flakonMap) {
  var type = determineType(row);
  var name = String(row[COL.NAME] || '').trim();
  var cost1C = toNumber_(row[COL.COST_1C], 0);
  var priceVal = toNumber_(row[COL.PRICE], 0);
  var ndsRate = toNumber_(row[COL.NDS], 0.22);
  var taxRate = toNumber_(row[COL.TAX], 0.065);
  var usd = toNumber_(params.usd, 92);
  var rmb = toNumber_(params.rmb, 12.8);
  var log = toNumber_(params.log, 4.5);
  var com = toNumber_(params.com, 1.05);

  var result = {
    name: name,
    type: type,
    raw: 0,
    taxDuty: 0,
    delivery: 0,
    rawFl: 0,
    deliveryFl: 0,
    taxDutyFl: 0,
    label: 0,
    totalFl: 0,
    total: 0,
    cost1C: cost1C,
    diff: 0,
    diffPct: 0
  };

  if (type === 'Наборы') {
    result.diff = cost1C > 0 ? -cost1C : 0;
    return result;
  }

  if (type === 'Готовый товар') {
    var finishedWeight = toNumber_(row[COL.WEIGHT], 0);
    result.raw = priceVal * rmb * com;
    result.delivery = log * finishedWeight * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;
  } else {
    var volume = toNumber_(row[COL.VOL], 0);
    result.raw = (priceVal * rmb * com / 1000) * volume;
    result.delivery = log * volume / 1000 * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;

    var flakonName = String(row[COL.FLAKON] || '').trim();
    var flakon = flakonMap[flakonName] || {
      rawFl: 0,
      deliveryFl: 0,
      taxDutyFl: 0,
      label: 0,
      totalFl: 0
    };
    result.rawFl = toNumber_(flakon.rawFl, 0);
    result.deliveryFl = toNumber_(flakon.deliveryFl, 0);
    result.taxDutyFl = toNumber_(flakon.taxDutyFl, 0);
    result.label = toNumber_(flakon.label, 0);
    result.totalFl = toNumber_(flakon.totalFl, 0);
  }

  result.raw = round2_(result.raw);
  result.taxDuty = round2_(result.taxDuty);
  result.delivery = round2_(result.delivery);
  result.rawFl = round2_(result.rawFl);
  result.deliveryFl = round2_(result.deliveryFl);
  result.taxDutyFl = round2_(result.taxDutyFl);
  result.label = round2_(result.label);
  result.totalFl = round2_(result.totalFl);

  result.total = round2_(
    result.raw +
    result.taxDuty +
    result.delivery +
    result.totalFl
  );

  result.diff = cost1C > 0 ? round2_(result.total - cost1C) : 0;
  result.diffPct = cost1C > 0 ? round4_((result.total - cost1C) / cost1C) : 0;

  return result;
}

function calculateAll(params) {
  var dataObj = getData();
  if (!dataObj.rows || dataObj.rows.length === 0) {
    return {
      success: false,
      message: 'Нет данных в 1cData. Сначала загрузите базовую номенклатуру.',
      results: []
    };
  }

  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  var results = [];

  for (var i = 0; i < dataObj.rows.length; i++) {
    results.push(calculateOne(dataObj.rows[i], params || {}, flakonMap));
  }

  saveParams(params || {});

  return { success: true, results: results, count: results.length };
}

function calculateByIndex(index, params) {
  var dataObj = getData();
  if (!dataObj.rows || index < 0 || index >= dataObj.rows.length) {
    return { success: false, message: 'Позиция не найдена.' };
  }

  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  return calculateOne(dataObj.rows[index], params || {}, flakonMap);
}

function calculateManual(input) {
  var row = new Array(COL.TAX + 1).fill('');
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
  return calculateOne(row, input || {}, flakonMap);
}

function getVerification(index, params) {
  var dataObj = getData();
  if (!dataObj.rows || index < 0 || index >= dataObj.rows.length) {
    return { success: false, message: 'Позиция не найдена.' };
  }

  var row = dataObj.rows[index];
  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  var result = calculateOne(row, params || {}, flakonMap);
  var type = determineType(row);
  var priceVal = toNumber_(row[COL.PRICE], 0);
  var ndsRate = toNumber_(row[COL.NDS], 0.22);
  var taxRate = toNumber_(row[COL.TAX], 0.065);
  var usd = toNumber_(params.usd, 92);
  var rmb = toNumber_(params.rmb, 12.8);
  var log = toNumber_(params.log, 4.5);
  var com = toNumber_(params.com, 1.05);
  var logFl = toNumber_(params.logFl, 4.5);
  var comFl = toNumber_(params.comFl, 1.05);
  var volume = toNumber_(row[COL.VOL], 0);
  var weight = toNumber_(row[COL.WEIGHT], 0);
  var flakonName = String(row[COL.FLAKON] || '').trim();
  var flakon = flakonMap[flakonName] || {
    supplierPrice: 0,
    weight: 0,
    nds: 0.22,
    tax: 0.065,
    label: 0,
    rawFl: 0,
    deliveryFl: 0,
    taxDutyFl: 0,
    totalFl: 0
  };
  var cost1C = toNumber_(row[COL.COST_1C], 0);
  var steps = [];

  steps.push({
    title: 'Определение типа',
    formula: 'Набор = "' + String(row[COL.IS_SET] || '') + '", Сырьё = "' + String(row[COL.RAW] || '') + '"',
    result: type
  });

  if (type === 'Наборы') {
    steps.push({
      title: 'Наборы не рассчитываются',
      formula: '',
      result: '0.00'
    });
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
      formula: '(Сырьё + Доставка + Сырьё×Пошлина) × НДС + Сырьё×Пошлина',
      result: result.taxDuty.toFixed(2)
    });
  } else {
    steps.push({
      title: 'Сырьё',
      formula: '(Цена × RMB × Ком / 1000) × Объём = (' + priceVal + ' × ' + rmb + ' × ' + com + ' / 1000) × ' + volume,
      result: result.raw.toFixed(2)
    });
    steps.push({
      title: 'Доставка',
      formula: 'Лог × Объём/1000 × 1.15 × USD × Ком = ' + log + ' × ' + (volume / 1000).toFixed(4) + ' × 1.15 × ' + usd + ' × ' + com,
      result: result.delivery.toFixed(2)
    });
    steps.push({
      title: 'Нал+Пош',
      formula: '(Сырьё + Доставка + Сырьё×Пошлина) × НДС + Сырьё×Пошлина',
      result: result.taxDuty.toFixed(2)
    });
    steps.push({
      title: 'Цена флакона в руб.',
      formula: 'Цена поставщика × RMB × Ком_fl = ' + toNumber_(flakon.supplierPrice, 0) + ' × ' + rmb + ' × ' + comFl,
      result: result.rawFl.toFixed(2)
    });
    steps.push({
      title: 'Доставка флакона в руб.',
      formula: 'Лог_fl × Вес × 1.15 × USD × Ком_fl = ' + logFl + ' × ' + toNumber_(flakon.weight, 0) + ' × 1.15 × ' + usd + ' × ' + comFl,
      result: result.deliveryFl.toFixed(2)
    });
    steps.push({
      title: 'НДС+пошлина флакона',
      formula: '(Цена флакона + Доставка флакона + Цена флакона×Пошлина_fl) × НДС_fl + Цена флакона×Пошлина_fl',
      result: result.taxDutyFl.toFixed(2)
    });
    steps.push({
      title: 'Этикетка',
      formula: 'Из таблицы флаконов',
      result: result.label.toFixed(2)
    });
    steps.push({
      title: 'Себестоимость флакона',
      formula: 'Цена флакона + Доставка флакона + НДС+пошлина + Этикетка',
      result: result.totalFl.toFixed(2)
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
    params: {
      usd: usd,
      rmb: rmb,
      log: log,
      com: com,
      logFl: logFl,
      comFl: comFl,
      ndsRate: ndsRate,
      taxRate: taxRate
    }
  };
}

function toNumber_(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  if (typeof value === 'number') return value;

  var normalized = String(value).replace(/\s+/g, '').replace(',', '.');
  var parsed = parseFloat(normalized);
  return isNaN(parsed) ? fallback : parsed;
}

function round2_(value) {
  return Math.round(value * 100) / 100;
}

function round4_(value) {
  return Math.round(value * 10000) / 10000;
}
