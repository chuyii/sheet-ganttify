import { SheetGanttify } from './SheetGanttify';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('SheetGanttify');

  menu.addItem(createCalendar.name, createCalendar.name);
  menu.addItem(createGantt.name, createGantt.name);
  menu.addItem(downloadCsvForImport.name, downloadCsvForImport.name);
  menu.addItem(showImportDialog.name, showImportDialog.name);
  menu.addItem(showCsvImportDialog.name, showCsvImportDialog.name);
  menu.addToUi();
}

function createCalendar() {
  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.createCalendar();
}

function createGantt() {
  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.parseDummy();
  sheetGanttify.createGantt();
  sheetGanttify.createLink();
  sheetGanttify.write();
}

function downloadCsvForImport() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('downloadCsv').evaluate(),
    'インポート用 CSV 生成'
  );
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function generateCsvForImport() {
  const sheetGanttify = SheetGanttify.getInstance();
  return sheetGanttify.generateCsvForImport();
}

function showImportDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('importForm'),
    'インポート'
  );
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function processForm(formObject: { result: string }) {
  const result = formObject.result.trim();
  const ticketIds = result
    .split('\n')
    .map(line => line.match(/[^#]*#([0-9]*):/)![1]);
  const sheetGanttify = SheetGanttify.getInstance();

  sheetGanttify.importTicketIds(ticketIds);
  sheetGanttify.createLink();
  sheetGanttify.write();
}

function showCsvImportDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('uploadForm').evaluate(),
    'CSV インポート'
  );
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function stUploaderV8(fObject: { mimeType: string; bytes: number[] }) {
  const csvData = Utilities.newBlob(
    fObject.bytes,
    fObject.mimeType
  ).getDataAsString(); // Convert received bytes to CSV string
  const parsed = Utilities.parseCsv(csvData); // Convert to a 2D array

  const shintyokuIndex = parsed[0].findIndex(v => v === '進捗率');
  const kaishiIndex = parsed[0].findIndex(v => v === '開始日');
  const kijitsuIndex = parsed[0].findIndex(v => v === '期日');
  const tantouIndex = parsed[0].findIndex(v => v === '担当者');
  parsed.shift();

  const dict: Record<string, [string, string, string, string]> = {};
  parsed.forEach(
    r =>
      (dict[r[0]] = [
        r[shintyokuIndex],
        r[kaishiIndex],
        r[kijitsuIndex],
        r[tantouIndex],
      ])
  );

  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.importInfo(dict);
  sheetGanttify.createGantt();
  sheetGanttify.createLink();
  sheetGanttify.write();
}
