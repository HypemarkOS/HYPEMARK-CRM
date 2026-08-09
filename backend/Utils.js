function now_() { return new Date(); }
function asText_(value) { return value == null ? '' : String(value).trim(); }
function normalizeEmail_(value) { return asText_(value).toLowerCase(); }
function normalizeMobile_(value) { return asText_(value).replace(/\D/g, ''); }
function headerMap_(headers) { var map = {}; headers.forEach(function(h,i){ map[asText_(h)] = i; }); return map; }
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers && headers.length && sheet.getLastRow() === 0) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  return sheet;
}
