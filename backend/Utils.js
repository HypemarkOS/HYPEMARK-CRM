function now_() { return new Date(); }
function asText_(value) { return value == null ? '' : String(value).trim(); }
function normalizeEmail_(value) { return asText_(value).toLowerCase(); }
function normalizeMobile_(value) { return asText_(value).replace(/\D/g, ''); }
function headerMap_(headers) { var map = {}; headers.forEach(function(h,i){ map[asText_(h)] = i; }); return map; }

function getCRMSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CRM_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('CRM_SPREADSHEET_ID', active.getId());
    return active;
  }

  throw new Error('CRM spreadsheet is not initialized. Run initializeCRM once.');
}

function getOrCreateSheet_(name, headers) {
  var ss = getCRMSpreadsheet_();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers && headers.length && sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sheet;
}
