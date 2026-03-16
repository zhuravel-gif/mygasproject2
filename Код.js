/**
 * ГЛАВНЫЙ ФАЙЛ УПРАВЛЕНИЯ ПРОЕКТОМ
 * Содержит меню, триггеры и логику обновления данных Calc
 */

const CFG = {
  CALC: 'Calc',
  S_1C: 'Себес 1С',
  SUPPLIER: 'Артикул и цены поставщика',
  FLAKONS: 'Флаконы 20.02.2026',
  REPORT: 'Себестоимость 20.02.2026',
  COL_START_NEW: 15 // Столбец O
};

/**
 * Создание пользовательского меню
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ ОБРАБОТКА')
    .addItem('🚀 Запустить обновление данных', 'mainExecution')
    .addItem('💰 Расчет себестоимости (V1)', 'showCostDialog')
    .addItem('🔴 Расчет себестоимости (V2)', 'showCostDialog2')
    .addSeparator()
    .addItem('📊 Перенести данные в итоговую таблицу', 'transferDataWithVLOOKUP')
    .addToUi();
}

/**
 * Единая точка входа для автоматического пересчета при правках в таблице
 */
//function onEdit(e) {
  // Вызов onEdit из файла Расчёт.gs (V1)
 // if (typeof onEditV1 === 'function') { 
 //   onEditV1(e); 
  //}
  
  // Вызов onEdit из файла Расчёт2.gs (V2)
  //if (typeof onEditV2 === 'function') { 
   // onEditV2(e); 
  //}
//}

/**
 * МАКРОС: Перенос данных в итоговую таблицу через формулы ВПР (VLOOKUP)
 */
function transferDataWithVLOOKUP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CFG.REPORT);
  const sourceSheetName = CFG.CALC;
  
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("Лист '" + CFG.REPORT + "' не найден!");
    return;
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Лист '" + CFG.REPORT + "' пуст.");
    return;
  }

  // Маппинг: [Буква столбца в итоговом листе, Индекс столбца в листе Calc]
  const mapping = [
    ['F', 22], // V
    ['G', 17], // Q
    ['H', 23], // W
    ['I', 24], // X
    ['J', 25], // Y
    ['K', 26], // Z
    ['L', 27], // AA
    ['M', 28], // AB
    ['N', 29], // AC
    ['O', 30], // AD
    ['Q', 32], // AF
    ['R', 33], // AG
    ['S', 34], // AH
    ['T', 35], // AI
    ['U', 36], // AJ
    ['V', 37], // AK
    ['W', 38], // AL
    ['X', 39], // AM
    ['Y', 15]  // O
  ];

  const numRows = lastRow - 1;

  mapping.forEach(pair => {
    const targetColLetter = pair[0];
    const sourceColIndex = pair[1];
    
    // Формула VLOOKUP с IFERROR
    const formula = `=IFERROR(VLOOKUP($A2; '${sourceSheetName}'!$A:$AM; ${sourceColIndex}; 0); "")`;
    
    const range = targetSheet.getRange(targetColLetter + "2:" + targetColLetter + lastRow);
    range.setFormula(formula);
  });

  ss.toast("Данные связаны через ВПР", "Успех ✅");
}

/**
 * ФУНКЦИЯ: Запуск обновления данных (Маппинг из 1С и Поставщика)
 */
function mainExecution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCalc = ss.getSheetByName(CFG.CALC);
  const sheet1C = ss.getSheetByName(CFG.S_1C);
  const sheetSupp = ss.getSheetByName(CFG.SUPPLIER);

  if (!sheetCalc || !sheet1C || !sheetSupp) {
    SpreadsheetApp.getUi().alert("Ошибка: Проверьте названия листов.");
    return;
  }

  const dataCalc = sheetCalc.getDataRange().getValues();
  const data1C = sheet1C.getDataRange().getValues();
  const dataSupp = sheetSupp.getDataRange().getValues();

  // 1. ИНДЕКСАЦИЯ 1С
  const map1C = {};
  data1C.slice(1).forEach(r => { 
    if(r[4]) map1C[String(r[4]).trim()] = r[8]; 
  });

  // 2. ИНДЕКСАЦИЯ ПОСТАВЩИКА
  const mapSupp = {};
  const suppHeaders = dataSupp[0];
  
  const newHeaderNames = [
    "Себестоимость 1С",     // O (15)
    suppHeaders[1] || "Артикул", 
    suppHeaders[7] || "Цена",    
    suppHeaders[10] || "Срок",   
    suppHeaders[11] || "Условия",
    "НДС",                  // T (20)
    "Пошлина"               // U (21)
  ];

  dataSupp.slice(1).forEach(r => {
    if(r[0]) {
      const key = String(r[0]).trim();
      mapSupp[key] = [r[1], r[7], r[10], r[11]];
    }
  });

  // 3. ПОДГОТОВКА ДАННЫХ
  const calcNewColumnsData = dataCalc.map((row, index) => {
    if (index === 0) return newHeaderNames;
    
    const nameKey = String(row[0]).trim(); 
    const artKey = String(row[7]).trim();  
    const val1C = map1C[nameKey] || "";
    
    let valSupp = mapSupp[nameKey];
    if (!valSupp || valSupp.every(item => item === "")) {
      valSupp = mapSupp[artKey] || ["", "", "", ""];
    }
    
    return [
      val1C,       
      valSupp[0],  
      valSupp[1],  
      valSupp[2],  
      valSupp[3],  
      0.22,        
      0.065        
    ];
  });

  // 4. ЗАПИСЬ И ФОРМАТИРОВАНИЕ
  const maxRows = sheetCalc.getMaxRows();
  const maxCols = sheetCalc.getMaxColumns();
  
  if (maxCols >= CFG.COL_START_NEW) {
    sheetCalc.getRange(1, CFG.COL_START_NEW, maxRows, maxCols - CFG.COL_START_NEW + 1).clearContent().clearFormat();
  }

  const numRows = calcNewColumnsData.length;
  const numCols = newHeaderNames.length;
  const outputRange = sheetCalc.getRange(1, CFG.COL_START_NEW, numRows, numCols);
  outputRange.setValues(calcNewColumnsData);

  outputRange.setNumberFormat("0.00");
  sheetCalc.getRange(2, 20, numRows - 1, 2).setNumberFormat("0.0%");

  // 5. ОБНОВЛЕНИЕ ЛИСТА ФЛАКОНЫ
  _generateFlakonReport(ss, dataCalc, mapSupp, newHeaderNames);

  ss.toast("Маппинг обновлен!", "Готово ✅");
}

/**
 * ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ: Генерация отчета по флаконам
 */
function _generateFlakonReport(ss, dataCalc, mapSupp, allHeaders) {
  let sheetFlak = ss.getSheetByName(CFG.FLAKONS) || ss.insertSheet(CFG.FLAKONS);
  sheetFlak.clear(); 

  const headers = ["Флаконы", "Объём", ...allHeaders];
  const reportMap = new Map();

  dataCalc.slice(1).forEach(row => {
    const fName = String(row[8]).trim(); 
    if (fName && fName !== "" && fName !== "undefined") {
      if (!reportMap.has(fName)) {
        const nameKey = String(row[0]).trim();
        const artKey = String(row[7]).trim();
        
        let sVals = mapSupp[nameKey];
        if (!sVals || sVals.every(v => v === "")) {
          sVals = mapSupp[artKey] || ["", "", "", ""];
        }
        
        reportMap.set(fName, [
          fName, 
          row[4], 
          "",        
          sVals[0],  
          sVals[1],  
          sVals[2],  
          sVals[3],  
          0.22,      
          0.065      
        ]);
      }
    }
  });

  if (reportMap.size > 0) {
    const output = [headers, ...Array.from(reportMap.values())];
    const numRows = output.length;
    const numCols = headers.length;
    const rng = sheetFlak.getRange(1, 1, numRows, numCols);
    rng.setValues(output);
    
    rng.setNumberFormat("0.00");
    sheetFlak.getRange(2, 8, numRows - 1, 2).setNumberFormat("0.0%");
    
    sheetFlak.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#efefef");
    sheetFlak.setFrozenRows(1);
    sheetFlak.autoResizeColumns(1, numCols);
  }
}