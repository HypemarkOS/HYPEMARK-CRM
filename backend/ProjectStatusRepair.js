function repairProjectStatusValidation() {
  var ss = getCRMSpreadsheet_();
  var sheet = ss.getSheetByName(APP.SHEETS.PROJECTS);
  if (!sheet) throw new Error('Projects sheet not found.');
  var statusIndex = headerIndex_(sheet, 'Status');
  if (statusIndex < 0) throw new Error('Project Status column not found.');
  var statuses = (CONFIG.STATUS_OPTIONS || []).slice();
  if (statuses.indexOf('Planned') < 0) statuses.unshift('Planned');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(false).build();
  sheet.getRange(2, statusIndex + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  return 'Project Status validation repaired. Allowed values: ' + statuses.join(', ');
}
