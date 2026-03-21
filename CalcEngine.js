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

function getDirectFlakon_(name, flakonMap) {
  var key = String(name || '').trim();
  if (!key || !flakonMap) return null;
  return Object.prototype.hasOwnProperty.call(flakonMap, key) ? flakonMap[key] : null;
}

function determineType(row, flakonMap, forcedType) {
  if (forcedType) return forcedType;

  var directFlakon = getDirectFlakon_(row[COL.NAME], flakonMap);
  var isSet = String(row[COL.IS_SET] || '').trim();
  var hasRaw = String(row[COL.RAW] || '').trim() !== '';

  if (directFlakon) return 'Флакон';
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

function buildDataRowMap_(rows) {
  var map = {};
  for (var i = 0; i < (rows || []).length; i++) {
    var key = normalizeMatchKey_(rows[i][COL.NAME]);
    if (key && !map.hasOwnProperty(key)) map[key] = rows[i];
  }
  return map;
}

function buildBundleContext_(params, flakonMap, dataRows) {
  var bundleRows = typeof getStoredBundleRows_ === 'function' ? getStoredBundleRows_() : [];
  var bundles = {};
  var dataRowMap = buildDataRowMap_(dataRows || []);

  for (var i = 0; i < bundleRows.length; i++) {
    var row = bundleRows[i];
    var bundleName = String(row.bundle || '').trim();
    var specName = String(row.specification || '').trim();
    if (!bundleName || !specName) continue;

    if (!bundles[bundleName]) {
      bundles[bundleName] = { specs: {}, specList: [], activeSpec: '' };
    }
    if (!bundles[bundleName].specs[specName]) {
      bundles[bundleName].specs[specName] = [];
      bundles[bundleName].specList.push(specName);
    }

    bundles[bundleName].specs[specName].push(row);
    if (row.active) bundles[bundleName].activeSpec = specName;
  }

  var bundleNames = Object.keys(bundles);
  for (var j = 0; j < bundleNames.length; j++) {
    var bundle = bundles[bundleNames[j]];
    if (!bundle.activeSpec) bundle.activeSpec = bundle.specList[0] || '';
  }

  return {
    params: params || {},
    flakonMap: flakonMap || {},
    dataRows: dataRows || [],
    dataRowMap: dataRowMap,
    bundles: bundles,
    cache: {}
  };
}

function getBundleDefinition_(bundleName, bundleContext) {
  if (!bundleContext || !bundleContext.bundles) return null;
  return bundleContext.bundles[String(bundleName || '').trim()] || null;
}

function getBundleActiveRows_(bundleName, bundleContext) {
  var bundleDef = getBundleDefinition_(bundleName, bundleContext);
  if (!bundleDef) return [];
  return bundleDef.specs[bundleDef.activeSpec] || [];
}

function buildBundleItemSource_(calcCost, manualCost, detailsSource) {
  if (calcCost > 0) return detailsSource || 'Расчёт';
  if (manualCost > 0) return 'Ручная стоимость';
  return 'Нет данных';
}

function calculateNamedCostInfo_(name, bundleContext, stack) {
  var key = normalizeMatchKey_(name);
  if (!key) {
    return { total: 0, source: 'Нет данных', type: '', specification: '', items: [] };
  }
  if (!bundleContext) {
    return { total: 0, source: 'Нет данных', type: '', specification: '', items: [] };
  }
  if (bundleContext.cache.hasOwnProperty(key)) return bundleContext.cache[key];

  stack = stack || {};
  if (stack[key]) {
    return { total: 0, source: 'Циклическая ссылка', type: 'Наборы', specification: '', items: [] };
  }

  var nextStack = {};
  for (var prop in stack) {
    if (stack.hasOwnProperty(prop)) nextStack[prop] = true;
  }
  nextStack[key] = true;

  var row = bundleContext.dataRowMap[key];
  var result;

  if (row) {
    var type = determineType(row, bundleContext.flakonMap);
    if (type === 'Наборы') {
      result = calculateBundleCostByName_(String(row[COL.NAME] || name), row, bundleContext, nextStack);
    } else {
      var calcResult = calculateOne(row, bundleContext.params, bundleContext.flakonMap, type, null, bundleContext, nextStack);
      result = {
        total: round2_(calcResult.total),
        source: 'Расчёт',
        type: calcResult.type,
        specification: '',
        items: [],
        result: calcResult
      };
    }
  } else {
    result = { total: 0, source: 'Нет данных', type: '', specification: '', items: [] };
  }

  bundleContext.cache[key] = result;
  return result;
}

function calculateBundleCostByName_(bundleName, row, bundleContext, stack) {
  var bundleDef = getBundleDefinition_(bundleName, bundleContext);
  var cost1C = row ? toNumber_(row[COL.COST_1C], 0) : 0;
  var result = {
    total: 0,
    source: 'Нет комплектации',
    type: 'Наборы',
    specification: '',
    items: [],
    cost1C: cost1C
  };

  if (!bundleDef) return result;

  var activeSpec = bundleDef.activeSpec || '';
  var activeRows = bundleDef.specs[activeSpec] || [];
  var total = 0;
  var items = [];

  for (var i = 0; i < activeRows.length; i++) {
    var componentRow = activeRows[i];
    var componentName = String(componentRow.component || '').trim();
    var quantity = toNumber_(componentRow.quantity, 1);
    var costInfo = calculateNamedCostInfo_(componentName, bundleContext, stack);
    var calcCost = round2_(toNumber_(costInfo.total, 0));
    var manualCost = componentRow.manualCost === '' ? 0 : round2_(toNumber_(componentRow.manualCost, 0));
    var usedCost = calcCost > 0 ? calcCost : manualCost;
    var lineTotal = round2_(usedCost * quantity);
    total += lineTotal;

    items.push({
      component: componentName,
      quantity: quantity,
      cost1C: toNumber_(componentRow.cost1C, 0),
      calcCost: calcCost,
      manualCost: manualCost,
      usedCost: usedCost,
      total: lineTotal,
      source: buildBundleItemSource_(calcCost, manualCost, costInfo.source)
    });
  }

  result.total = round2_(total);
  result.source = activeRows.length ? 'Комплектация набора' : 'Нет строк активной спецификации';
  result.specification = activeSpec;
  result.items = items;
  return result;
}

function buildBundleUiState_() {
  var params = getParams();
  var dataObj = getData();
  var flakonMap = buildFlakonMap(getFlakonList());
  var bundleContext = buildBundleContext_(params, flakonMap, dataObj.rows || []);
  var storedRows = typeof getStoredBundleRows_ === 'function' ? getStoredBundleRows_() : [];
  var resolvedRows = [];
  var bundleNames = {};
  var knownSetNames = {};
  var manualMap = {};
  var i;

  for (i = 0; i < (dataObj.rows || []).length; i++) {
    if (String(dataObj.rows[i][COL.IS_SET] || '').trim() === 'Да') {
      knownSetNames[String(dataObj.rows[i][COL.NAME] || '').trim()] = true;
    }
  }

  for (i = 0; i < storedRows.length; i++) {
    var item = storedRows[i];
    var info = calculateNamedCostInfo_(item.component, bundleContext, {});
    var calcCost = round2_(toNumber_(info.total, 0));
    var manualCost = item.manualCost === '' ? '' : round2_(toNumber_(item.manualCost, 0));
    var usedCost = calcCost > 0 ? calcCost : (manualCost === '' ? 0 : manualCost);
    var componentRow = bundleContext.dataRowMap[normalizeMatchKey_(item.component)] || null;
    var cost1C = componentRow ? toNumber_(componentRow[COL.COST_1C], 0) : 0;
    var source = buildBundleItemSource_(calcCost, manualCost === '' ? 0 : manualCost, info.source);

    resolvedRows.push({
      component: item.component,
      bundle: item.bundle,
      specification: item.specification,
      quantity: toNumber_(item.quantity, 1),
      active: !!item.active,
      cost1C: round2_(cost1C),
      calcCost: calcCost,
      manualCost: manualCost,
      usedCost: round2_(usedCost),
      source: source
    });

    bundleNames[item.bundle] = true;

    var manualKey = normalizeMatchKey_(item.component);
    if (!manualMap[manualKey] && calcCost <= 0) {
      manualMap[manualKey] = {
        component: item.component,
        manualCost: manualCost === '' ? '' : manualCost,
        usageCount: 1
      };
    } else if (manualMap[manualKey] && calcCost <= 0) {
      manualMap[manualKey].usageCount++;
      if (manualMap[manualKey].manualCost === '' && manualCost !== '') {
        manualMap[manualKey].manualCost = manualCost;
      }
    }
  }

  resolvedRows.sort(function(a, b) {
    if (a.bundle !== b.bundle) return a.bundle.localeCompare(b.bundle, 'ru', { sensitivity: 'base', numeric: true });
    if (a.specification !== b.specification) return a.specification.localeCompare(b.specification, 'ru', { sensitivity: 'base', numeric: true });
    return a.component.localeCompare(b.component, 'ru', { sensitivity: 'base', numeric: true });
  });

  var bundleSummary = [];
  var bundleKeys = Object.keys(bundleContext.bundles);
  for (i = 0; i < bundleKeys.length; i++) {
    var bundleName = bundleKeys[i];
    bundleSummary.push({
      bundle: bundleName,
      activeSpec: bundleContext.bundles[bundleName].activeSpec || '',
      specs: bundleContext.bundles[bundleName].specList.slice()
    });
  }
  bundleSummary.sort(function(a, b) {
    return a.bundle.localeCompare(b.bundle, 'ru', { sensitivity: 'base', numeric: true });
  });

  return {
    rows: resolvedRows,
    bundles: bundleSummary,
    manualItems: mapValues_(manualMap).sort(function(a, b) {
      return a.component.localeCompare(b.component, 'ru', { sensitivity: 'base', numeric: true });
    }),
    stats: {
      totalSets: Object.keys(knownSetNames).length,
      loadedBundles: bundleSummary.filter(function(item) { return knownSetNames[item.bundle]; }).length,
      compositionRows: resolvedRows.length,
      unresolvedComponents: mapValues_(manualMap).length
    }
  };
}

function calculateOne(row, params, flakonMap, forcedType, flakonNameOverride, bundleContext, stack) {
  params = params || {};
  flakonMap = flakonMap || {};

  var type = determineType(row, flakonMap, forcedType);
  var name = String(row[COL.NAME] || '').trim();
  var category = String(row[COL.CAT] || '').trim();
  var flakonNameSource = String(row[COL.FLAKON] || '').trim();
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
    flakonName: flakonNameSource,
    category: category,
    group1: String(row[COL.GRP1] || '').trim(),
    group2: String(row[COL.GRP2] || '').trim(),
    group3: String(row[COL.GRP3] || '').trim(),
    rawName: String(row[COL.RAW] || '').trim(),
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
    diffPct: 0,
    bundleSpec: '',
    bundleItems: []
  };

  if (type === 'Наборы') {
    var runtimeContext = bundleContext || buildBundleContext_(params, flakonMap, getData().rows || []);
    var bundleCalc = calculateBundleCostByName_(name, row, runtimeContext, stack || {});
    result.total = round2_(bundleCalc.total);
    result.bundleSpec = bundleCalc.specification || '';
    result.bundleItems = bundleCalc.items || [];
    result.diff = cost1C > 0 ? round2_(result.total - cost1C) : 0;
    result.diffPct = cost1C > 0 ? round4_((result.total - cost1C) / cost1C) : 0;
    return result;
  }

  if (type === 'Флакон') {
    var directFlakon = getDirectFlakon_(flakonNameOverride || name, flakonMap);
    var directMetrics = directFlakon || calculateFlakonMetrics_({
      supplierPrice: priceVal,
      weight: toNumber_(row[COL.WEIGHT], 0),
      nds: ndsRate,
      tax: taxRate,
      label: 0
    }, params);

    result.rawFl = toNumber_(directMetrics.rawFl, 0);
    result.deliveryFl = toNumber_(directMetrics.deliveryFl, 0);
    result.taxDutyFl = toNumber_(directMetrics.taxDutyFl, 0);
    result.label = toNumber_(directMetrics.label, 0);
    result.totalFl = toNumber_(directMetrics.totalFl, 0);
  } else if (type === 'Готовый товар') {
    var finishedWeight = toNumber_(row[COL.WEIGHT], 0);
    result.raw = priceVal * rmb * com;
    result.delivery = log * finishedWeight * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;
  } else {
    var volume = toNumber_(row[COL.VOL], 0);
    result.raw = (priceVal * rmb * com / 1000) * volume;
    result.delivery = log * volume / 1000 * 1.15 * usd * com;
    result.taxDuty = (result.raw + result.delivery + result.raw * taxRate) * ndsRate + result.raw * taxRate;

    var flakonName = String(flakonNameOverride || row[COL.FLAKON] || '').trim();
    var flakon = getDirectFlakon_(flakonName, flakonMap) || {
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
  var bundleContext = buildBundleContext_(params || {}, flakonMap, dataObj.rows || []);
  var results = [];

  for (var i = 0; i < dataObj.rows.length; i++) {
    results.push(calculateOne(dataObj.rows[i], params || {}, flakonMap, null, null, bundleContext, {}));
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
  var bundleContext = buildBundleContext_(params || {}, flakonMap, dataObj.rows || []);
  return calculateOne(dataObj.rows[index], params || {}, flakonMap, null, null, bundleContext, {});
}

function hasManualFlakonInput_(manualFlakon) {
  if (!manualFlakon) return false;
  return manualFlakon.supplierPrice !== '' ||
    manualFlakon.weight !== '' ||
    manualFlakon.nds !== '' ||
    manualFlakon.tax !== '' ||
    manualFlakon.label !== '';
}

function applyManualFlakonOverride_(flakonMap, flakonName, manualFlakon, params) {
  if (!flakonName) return;

  var base = getDirectFlakon_(flakonName, flakonMap) || {};
  flakonMap[flakonName] = {
    name: flakonName,
    volume: toNumber_(base.volume, 0),
    supplierPrice: manualFlakon.supplierPrice !== '' ? toNumber_(manualFlakon.supplierPrice, 0) : toNumber_(base.supplierPrice, 0),
    weight: manualFlakon.weight !== '' ? toNumber_(manualFlakon.weight, 0) : toNumber_(base.weight, 0),
    nds: manualFlakon.nds !== '' ? toNumber_(manualFlakon.nds, toNumber_(params.flakonNds, 0.22)) : toNumber_(base.nds, toNumber_(params.flakonNds, 0.22)),
    tax: manualFlakon.tax !== '' ? toNumber_(manualFlakon.tax, toNumber_(params.flakonTax, 0.065)) : toNumber_(base.tax, toNumber_(params.flakonTax, 0.065)),
    label: manualFlakon.label !== '' ? toNumber_(manualFlakon.label, 0) : toNumber_(base.label, 0)
  };
}

function buildCalculationPayload_(row, params, flakonMap, forcedType, flakonNameOverride, sourceLabel, bundleContext) {
  params = params || {};
  flakonMap = flakonMap || {};

  var type = determineType(row, flakonMap, forcedType);
  var result = calculateOne(row, params, flakonMap, forcedType, flakonNameOverride, bundleContext, {});
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
  var flakonName = type === 'Флакон'
    ? String(flakonNameOverride || row[COL.NAME] || '').trim()
    : String(flakonNameOverride || row[COL.FLAKON] || '').trim();
  var flakon = getDirectFlakon_(flakonName, flakonMap) || {
    supplierPrice: 0,
    weight: 0,
    nds: toNumber_(params.flakonNds, 0.22),
    tax: toNumber_(params.flakonTax, 0.065),
    label: 0,
    rawFl: 0,
    deliveryFl: 0,
    taxDutyFl: 0,
    totalFl: 0
  };
  var cost1C = toNumber_(row[COL.COST_1C], 0);
  var steps = [];
  var directFlakon = getDirectFlakon_(row[COL.NAME], flakonMap);
  var flSourceText = sourceLabel || 'Позиция найдена на листе "Флаконы" по совпадению Наименование = Флакон';
  var labelSourceText = sourceLabel && sourceLabel.indexOf('руч') >= 0 ? 'Задано вручную' : 'Из таблицы флаконов';

  steps.push({
    title: 'Определение типа',
    formula: 'Флакон = "' + (directFlakon ? 'Да' : 'Нет') + '", Набор = "' + String(row[COL.IS_SET] || '') + '", Сырьё = "' + String(row[COL.RAW] || '') + '"',
    result: type
  });

  if (type === 'Наборы') {
    steps.push({
      title: 'Активная спецификация набора',
      formula: 'Для набора используется выбранная пользователем активная спецификация',
      result: result.bundleSpec || 'Не выбрана'
    });
    for (var itemIdx = 0; itemIdx < (result.bundleItems || []).length; itemIdx++) {
      var bundleItem = result.bundleItems[itemIdx];
      steps.push({
        title: 'Компонент: ' + bundleItem.component,
        formula: 'Количество ' + bundleItem.quantity + ' × стоимость ' + bundleItem.usedCost.toFixed(2) + ' [' + bundleItem.source + ']',
        result: bundleItem.total.toFixed(2)
      });
    }
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
  } else if (type === 'Флакон') {
    steps.push({
      title: 'Источник расчёта флакона',
      formula: flSourceText,
      result: flakonName || '—'
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
      formula: labelSourceText,
      result: result.label.toFixed(2)
    });
    steps.push({
      title: 'Себестоимость флакона',
      formula: 'Цена флакона + Доставка флакона + НДС+пошлина + Этикетка',
      result: result.totalFl.toFixed(2)
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
      formula: labelSourceText,
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
      ndsRate: type === 'Флакон' ? toNumber_(flakon.nds, ndsRate) : ndsRate,
      taxRate: type === 'Флакон' ? toNumber_(flakon.tax, taxRate) : taxRate
    }
  };
}

function calculateManual(input) {
  input = input || {};
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
  var dataObj = getData();
  var bundleContext = buildBundleContext_(input, flakonMap, dataObj.rows || []);
  var manualFlakon = input.manualFlakon || {};
  var flakonNameOverride = String(input.flakonName || '').trim();
  if (!flakonNameOverride && (input.type === 'Флакон' || hasManualFlakonInput_(manualFlakon))) {
    flakonNameOverride = '__manual_flakon__';
  }
  if (flakonNameOverride) {
    applyManualFlakonOverride_(flakonMap, flakonNameOverride, manualFlakon, input);
  }

  var sourceLabel = flakonNameOverride === '__manual_flakon__'
    ? 'Флакон и его параметры заданы вручную'
    : (flakonNameOverride ? 'Использован выбранный флакон с возможностью ручной корректировки' : '');

  return buildCalculationPayload_(row, input, flakonMap, input.type, flakonNameOverride, sourceLabel, bundleContext);
}

function getVerification(index, params) {
  var dataObj = getData();
  if (!dataObj.rows || index < 0 || index >= dataObj.rows.length) {
    return { success: false, message: 'Позиция не найдена.' };
  }

  var row = dataObj.rows[index];
  var flakons = getFlakonList();
  var flakonMap = buildFlakonMap(flakons);
  var bundleContext = buildBundleContext_(params || {}, flakonMap, dataObj.rows || []);
  return buildCalculationPayload_(row, params || {}, flakonMap, null, null, '', bundleContext);
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
