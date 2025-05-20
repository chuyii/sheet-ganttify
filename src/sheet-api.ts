export class SheetApi {
  static batchGet(spreadsheetId: string, params: object) {
    return Sheets.Spreadsheets?.Values?.batchGet(spreadsheetId, params);
  }

  static batchUpdate(
    resource: GoogleAppsScript.Sheets.Schema.BatchUpdateValuesRequest,
    spreadsheetId: string
  ) {
    return Sheets.Spreadsheets?.Values?.batchUpdate(resource, spreadsheetId);
  }

  static batchClear(
    resource: GoogleAppsScript.Sheets.Schema.BatchClearValuesRequest,
    spreadsheetId: string
  ) {
    return Sheets.Spreadsheets?.Values?.batchClear(resource, spreadsheetId);
  }
}
