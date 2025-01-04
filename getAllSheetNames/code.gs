function getAllSheetNames(spreadsheetId) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(function(sheet) {
    return sheet.getName();
  });
  return sheetNames;
}
