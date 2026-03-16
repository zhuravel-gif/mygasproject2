/**
 * ГЛАВНЫЙ ФАЙЛ — точка входа WebApp
 */

function doGet() {
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
    .addItem('Экспорт данных в 1cData', 'exportToDataSheet')
    .addToUi();
}

function openWebApp() {
  const html = HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Калькулятор себестоимости');
}
