/**
 * ВЕРСИЯ №2: ЖЕСТКАЯ ПРИВЯЗКА К СТОЛБЦАМ AE:AM (31-39)
 * ФЛАКОНЫ: E (этикетка), J (цена), K (налоги), L (доставка)
 */

const CFG_V2 = {
  SHEET_CALC: 'Calc',
  SHEET_FLAKON: 'Флаконы 20.02.2026',
  SHEET_BASIS: 'basis2',
  START_COL: 31, // Столбец AE
  NUM_COLS: 9    // Количество создаваемых колонок
};

function onEditV2(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== CFG_V2.SHEET_CALC || range.getRow() <= 1) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const basisSheet = ss.getSheetByName(CFG_V2.SHEET_BASIS);
  if (!basisSheet) return;

  const bData = basisSheet.getRange(2, 1, 1, 8).getValues()[0];
  const p = {
    usd: bData[0], rmb: bData[1], log: bData[2], com: bData[3],
    nds: bData[4], tax: bData[5], ves: bData[6], priceCol: bData[7]
  };

  const watchCols = [5, 9, 12, _letterToIdxV2(p.priceCol) + 1, _letterToIdxV2(p.nds) + 1, _letterToIdxV2(p.tax) + 1, _letterToIdxV2(p.ves) + 1];

  if (watchCols.includes(range.getColumn())) {
    runCalculationV2(p);
  }
}

function showCostDialog2() {
  const html = HtmlService.createHtmlOutputFromFile('Form2')
    .setWidth(400)
    .setHeight(620)
    .setTitle('Параметры расчета V2');
  SpreadsheetApp.getUi().showModalDialog(html, 'Введите данные V2');
}

function runCalculationV2(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _saveToBasisV2(ss, p);
  const sheet = ss.getSheetByName(CFG_V2.SHEET_CALC);
  const flSheet = ss.getSheetByName(CFG_V2.SHEET_FLAKON);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const flakonMap = _getFlakonMapV2(flSheet);

  const cNds = _letterToIdxV2(p.nds);
  const cTax = _letterToIdxV2(p.tax);
  const cVes = _letterToIdxV2(p.ves);
  const cPrice = _letterToIdxV2(p.priceCol);

  const finalData = rows.map((row) => {
    const L = String(row[11]).trim();
    const H = row[7];
    const E = parseFloat(row[4]) || 0;
    const I = String(row[8]).trim();
    const priceVal = parseFloat(row[cPrice]) || 0;
    const nds = parseFloat(row[cNds]) || 0;
    const tax = parseFloat(row[cTax]) || 0;
    const ves = parseFloat(row[cVes]) || 0;

    let res = { type: "Не определено", s:0, nps:0, ds:0, fl:0, npf:0, df:0, et:0 };

    if (L === "Да") { 
      res.type = "Наборы"; 
    } 
    else if (L !== "Да" && (!H || H == 0 || H == "")) {
      res.type = "Готовый товар";
      res.s = priceVal * p.rmb * p.com;
      res.ds = p.log * ves * 1.15 * p.usd * p.com;
      res.nps = ((res.s + res.ds + (res.s * tax)) * nds + (res.s * tax));
    } else if (L !== "Да" && H) {
      res.type = "Сырьё";
      const f = flakonMap[I] || { g: 0, h: 0, i: 0, et: 0 };
      res.s = (priceVal * p.rmb * p.com / 1000) * E;
      res.ds = p.log * E / 1000 * 1.15 * p.usd * p.com;
      res.nps = ((res.s + res.ds + (res.s * tax)) * nds + (res.s * tax));
      res.fl = f.g; 
      res.npf = f.h; 
      res.df = f.i;
      res.et = f.et; // Заполнение этикетки из маппинга (столбец E)
    }
    const total = res.s + res.nps + res.ds + res.fl + res.npf + res.df + res.et;
    return [res.type, Number(res.s.toFixed(2)), Number(res.nps.toFixed(2)), Number(res.ds.toFixed(2)), Number(res.fl.toFixed(2)), Number(res.npf.toFixed(2)), Number(res.df.toFixed(2)), Number(res.et.toFixed(2)), Number(total.toFixed(2))];
  });

  const outHeaders = ["V2 Тип", "V2 Сырьё", "V2 Нал+Пош", "V2 Доставка", "V2 Флакон", "V2 Нал Фл", "V2 Дост Фл", "V2 Этикетка", `V2 Итого (${p.usd})`];
  
  sheet.getRange(1, CFG_V2.START_COL, sheet.getLastRow(), CFG_V2.NUM_COLS).clearContent();
  const headerRange = sheet.getRange(1, CFG_V2.START_COL, 1, outHeaders.length);
  headerRange.setValues([outHeaders]).setBackground("#fce8e6").setFontWeight("bold");
  sheet.getRange(2, CFG_V2.START_COL, finalData.length, outHeaders.length).setValues(finalData);
  
  return "Расчет V2 завершен!";
}

function _getFlakonMapV2(flSheet) {
  const map = {};
  if (flSheet) {
    const data = flSheet.getDataRange().getValues();
    // V2 берет этикетку из E (индекс 4), а данные флакона из J, K, L (индексы 9, 10, 11)
    data.slice(1).forEach(r => { 
      map[String(r[0]).trim()] = { et: r[4] || 0, g: r[9] || 0, h: r[10] || 0, i: r[11] || 0 }; 
    });
  }
  return map;
}

function _saveToBasisV2(ss, p) {
  let bSheet = ss.getSheetByName(CFG_V2.SHEET_BASIS) || ss.insertSheet(CFG_V2.SHEET_BASIS);
  bSheet.clear();
  bSheet.getRange(1, 1, 1, 8).setValues([["USD", "RMB", "Logistics", "Commission", "Col_NDS", "Col_Tax", "Col_Weight", "Col_Price"]]).setFontWeight("bold");
  bSheet.getRange(2, 1, 1, 8).setValues([[p.usd, p.rmb, p.log, p.com, p.nds, p.tax, p.ves, p.priceCol]]);
}

function _letterToIdxV2(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) { col += (letter.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1); }
  return col - 1;
}