/**
 * ВЕРСИЯ №1: ЖЕСТКАЯ ПРИВЯЗКА К СТОЛБЦАМ V:AD (22-30)
 * ФЛАКОНЫ: E (этикетка), G (цена), H (налоги), I (доставка)
 */

const CFG_V1 = {
  SHEET_CALC: 'Calc',
  SHEET_FLAKON: 'Флаконы 20.02.2026',
  SHEET_BASIS: 'basis',
  START_COL: 22, // Столбец V
  NUM_COLS: 9    // Количество создаваемых колонок
};

function onEditV1(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== CFG_V1.SHEET_CALC || range.getRow() <= 1) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const basisSheet = ss.getSheetByName(CFG_V1.SHEET_BASIS);
  if (!basisSheet) return;

  const bData = basisSheet.getRange(2, 1, 1, 8).getValues()[0];
  const p = {
    usd: bData[0], rmb: bData[1], log: bData[2], com: bData[3],
    nds: bData[4], tax: bData[5], ves: bData[6], priceCol: bData[7]
  };

  const watchCols = [5, 9, 12, _letterToIdxV1(p.priceCol) + 1, _letterToIdxV1(p.nds) + 1, _letterToIdxV1(p.tax) + 1, _letterToIdxV1(p.ves) + 1];

  if (watchCols.includes(range.getColumn())) {
    runCalculation(p);
  }
}

function showCostDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(400)
    .setHeight(620)
    .setTitle('Параметры расчета V1');
  SpreadsheetApp.getUi().showModalDialog(html, 'Введите данные V1');
}

function runCalculation(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _saveToBasisV1(ss, p);
  const sheet = ss.getSheetByName(CFG_V1.SHEET_CALC);
  const flSheet = ss.getSheetByName(CFG_V1.SHEET_FLAKON);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const flakonMap = _getFlakonMapV1(flSheet);

  const cNds = _letterToIdxV1(p.nds);
  const cTax = _letterToIdxV1(p.tax);
  const cVes = _letterToIdxV1(p.ves);
  const cPrice = _letterToIdxV1(p.priceCol);

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

  const outHeaders = ["V1 Тип", "V1 Сырьё", "V1 Нал+Пош", "V1 Доставка", "V1 Флакон", "V1 Нал Фл", "V1 Дост Фл", "V1 Этикетка", `V1 Итого (${p.usd})`];
  
  sheet.getRange(1, CFG_V1.START_COL, sheet.getLastRow(), CFG_V1.NUM_COLS).clearContent();
  sheet.getRange(1, CFG_V1.START_COL, 1, outHeaders.length).setValues([outHeaders]).setFontWeight("bold");
  sheet.getRange(2, CFG_V1.START_COL, finalData.length, outHeaders.length).setValues(finalData);
  
  return "Расчет V1 завершен!";
}

function _getFlakonMapV1(flSheet) {
  const map = {};
  if (flSheet) {
    const data = flSheet.getDataRange().getValues();
    // V1 берет E (индекс 4), G (индекс 6), H (индекс 7), I (индекс 8)
    data.slice(1).forEach(r => { 
      map[String(r[0]).trim()] = { et: r[4] || 0, g: r[6] || 0, h: r[7] || 0, i: r[8] || 0 }; 
    });
  }
  return map;
}

function _saveToBasisV1(ss, p) {
  let bSheet = ss.getSheetByName(CFG_V1.SHEET_BASIS) || ss.insertSheet(CFG_V1.SHEET_BASIS);
  bSheet.clear();
  bSheet.getRange(1, 1, 1, 8).setValues([["USD", "RMB", "Logistics", "Commission", "Col_NDS", "Col_Tax", "Col_Weight", "Col_Price"]]).setFontWeight("bold");
  bSheet.getRange(2, 1, 1, 8).setValues([[p.usd, p.rmb, p.log, p.com, p.nds, p.tax, p.ves, p.priceCol]]);
}

function _letterToIdxV1(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) { col += (letter.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1); }
  return col - 1;
}