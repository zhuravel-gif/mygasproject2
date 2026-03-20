/**
 * Main entry points for Apps Script WebApp.
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
    .addToUi();
}

function openWebApp() {
  var html = HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setWidth(1280)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Калькулятор себестоимости');
}
