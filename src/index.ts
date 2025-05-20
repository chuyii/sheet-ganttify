import { SheetGanttify } from './SheetGanttify';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('SheetGanttify');

  menu.addItem(createCalendar.name, createCalendar.name);
  menu.addItem(createGanttChart.name, createGanttChart.name);
  menu.addItem(showCsvDownloadDialog.name, showCsvDownloadDialog.name);
  menu.addItem(showImportDialog.name, showImportDialog.name);
  menu.addItem(showCsvImportDialog.name, showCsvImportDialog.name);
  menu.addToUi();
}

function createCalendar() {
  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.createCalendar();
}

function createGanttChart() {
  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.parseDummy();
  sheetGanttify.createGantt();
  sheetGanttify.createLink();
  sheetGanttify.write();
}

function showCsvDownloadDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('downloadCsvForImport').evaluate(),
    'インポート用 CSV 生成'
  );
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function generateImportCsv() {
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
function processTicketIdForm(formObject: { result: string }) {
  const trimmedResult = formObject.result.trim();
  const ticketIds = trimmedResult
    .split('\n')
    .map(line => line.match(/[^#]*#([0-9]*):/)![1]);
  const sheetGanttify = SheetGanttify.getInstance();

  sheetGanttify.importTicketIds(ticketIds);
  sheetGanttify.createLink();
  sheetGanttify.write();
}

function showCsvImportDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createTemplateFromFile('csvImportForm').evaluate(),
    'CSV インポート'
  );
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function handleCsvUpload(fileObject: { mimeType: string; bytes: number[] }) {
  const csvString = Utilities.newBlob(
    fileObject.bytes,
    fileObject.mimeType
  ).getDataAsString(); // Convert received bytes to CSV string
  const parsedCsv = Utilities.parseCsv(csvString); // Convert to a 2D array

  const progressIndex = parsedCsv[0].findIndex(v => v === '進捗率');
  const startDateIndex = parsedCsv[0].findIndex(v => v === '開始日');
  const dueDateIndex = parsedCsv[0].findIndex(v => v === '期日');
  const assigneeIndex = parsedCsv[0].findIndex(v => v === '担当者');
  parsedCsv.shift();

  const infoMap: Record<string, [string, string, string, string]> = {};
  parsedCsv.forEach(
    row =>
      (infoMap[row[0]] = [
        row[progressIndex],
        row[startDateIndex],
        row[dueDateIndex],
        row[assigneeIndex],
      ])
  );

  const sheetGanttify = SheetGanttify.getInstance();
  sheetGanttify.importInfo(infoMap);
  sheetGanttify.createGantt();
  sheetGanttify.createLink();
  sheetGanttify.write();
}
