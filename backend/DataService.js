function appendRow_(sheetName, headers, rowObject) {
  var sheet = getOrCreateSheet_(sheetName, headers);
  var headersNow = sheet.getRange(1,1,1,headers.length).getValues()[0];
  var row = headersNow.map(function(h){ return rowObject[h] !== undefined ? rowObject[h] : ''; });
  sheet.appendRow(row);
  return row;
}

function getRows_(sheetName) {
  var sheet = getCRMSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues();
}
