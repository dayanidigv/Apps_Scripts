
async function sheetToJson(spreadsheetId, sheetName) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  var headers = data[0];

  for (var i = 0; i < headers.length; i++) {
    headers[i] = headers[i]
      .replace(/[.\-/\\]/g, "_") // Replace '.', '-', '/', and '\' with '_'
      .toLowerCase();           // Convert to lowercase
  }

  var jsonArray = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var jsonObject = {};

    for (var j = 0; j < headers.length; j++) {
      jsonObject[headers[j]] = row[j];
    }

    jsonArray.push(jsonObject);
  }

  // Logger.log(JSON.stringify(jsonArray, null, 2));
  return jsonArray;
}
