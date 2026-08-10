var AUDIT_SHEETS = {
  'Clients': 'Client',
  'Engagements': 'Engagement',
  'Projects': 'Project',
  'Deliverables': 'Deliverable',
  'ContentBank': 'Content',
  'ClientReceipts': 'Client Receipt',
  'EmployeePayments': 'Employee Payment',
  'BusinessExpenses': 'Business Expense',
  'Users': 'User'
};

function recordActivity_(entity, entityId, action) {
  appendRow_(APP.SHEETS.ACTIVITIES, ['Activity ID','Entity','Entity ID','Action','User','Date Time'], {
    'Activity ID': generateId_(CONFIG.ID_PREFIX.ACTIVITY),
    'Entity': entity,
    'Entity ID': entityId,
    'Action': action,
    'User': getCurrentUser_(),
    'Date Time': now_()
  });
}

function getCurrentUser_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'System';
  } catch (err) {
    return 'System';
  }
}

function handleCRMEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (!AUDIT_SHEETS[sheetName] || e.range.getRow() === 1) return;

  var row = e.range.getRow();
  var lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var rowValues = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  var entityId = rowValues[0] || row;
  var entity = AUDIT_SHEETS[sheetName];
  var startCol = e.range.getColumn();
  var numCols = e.range.getNumColumns();
  var action;

  if (numCols === 1) {
    var header = headers[startCol - 1] || ('Column ' + startCol);
    var oldValue = e.oldValue === undefined ? '(blank)' : String(e.oldValue);
    var newValue = e.value === undefined ? '(blank)' : String(e.value);
    if (oldValue === newValue) return;
    action = 'Updated ' + header + ': ' + oldValue + ' → ' + newValue;
  } else {
    action = 'Edited ' + sheetName + ' row ' + row + ' (' + numCols + ' cells)';
  }

  recordActivity_(entity, String(entityId), action);
}

function getActivitiesService_() {
  var s = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.ACTIVITIES);
  if (!s || s.getLastRow() < 2) return [];
  var values = s.getDataRange().getValues();
  var headers = values.shift();
  return values.filter(function(row) { return row.some(function(value) { return value !== ''; }); }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) {
      object[header] = row[index] instanceof Date ? Utilities.formatDate(row[index], CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm') : row[index];
    });
    return object;
  });
}
