function initializeCRMService_() {
  var props = PropertiesService.getScriptProperties();
  var ss, id = props.getProperty('CRM_SPREADSHEET_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create(APP.NAME + ' Data');
    props.setProperty('CRM_SPREADSHEET_ID', ss.getId());
  }

  var definitions = [
    [APP.SHEETS.CLIENTS, ['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On']],
    [APP.SHEETS.ENGAGEMENTS, ['Engagement ID','Client ID','Service','Start Date','End Date','Status']],
    [APP.SHEETS.PROJECTS, ['Project ID','Client ID','Engagement ID','Project Name','Status','Created On']],
    [APP.SHEETS.DELIVERABLES, ['Deliverable ID','Project ID','Type','Quantity','Status']],
    [APP.SHEETS.CONTENT_BANK, ['Content ID','Deliverable ID','Title','Status','Assigned To','Publish Date']],
    [APP.SHEETS.ACTIVITIES, ['Activity ID','Entity','Entity ID','Action','User','Date Time']],
    [APP.SHEETS.CLIENT_RECEIPTS, ['Receipt ID','Client ID','Amount','Payment Date','Received By','Payment Mode','Reference','Status','Notes']],
    [APP.SHEETS.EMPLOYEE_PAYMENTS, ['Employee Payment ID','Employee','Client ID','Payment Type','Amount','Payment Date','Commission %','Commission Base Amount','Paid From','Payment Mode','Reference','Status','Notes']],
    [APP.SHEETS.BUSINESS_EXPENSES, ['Expense ID','Client ID','Expense Category','Related To','Amount','Expense Date','Paid To','Payment Mode','Reference','Description','Status']],
    [APP.SHEETS.USERS, ['User ID','Name','Role','Email','Status']],
    [APP.SHEETS.SETTINGS, ['Key','Value']]
  ];

  definitions.forEach(function(def) {
    var s = ss.getSheetByName(def[0]) || ss.insertSheet(def[0]);
    ensureHeaders_(s, def[1]);
    styleSheet_(s);
  });

  migrateLegacyPayments_(ss);
  applyFinanceValidations_(ss);
  ensureCRMTriggers_(ss);

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0 && defaultSheet.getLastColumn() === 1) {
    ss.deleteSheet(defaultSheet);
  }

  console.log('HYPEMARK CRM spreadsheet: ' + ss.getUrl());
  return ss.getUrl();
}

function ensureHeaders_(sheet, headers) {
  var current = sheet.getLastColumn() || 0;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  var existing = sheet.getRange(1, 1, 1, current).getValues()[0];
  headers.forEach(function(h) {
    if (existing.indexOf(h) === -1) sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
  });
}

function styleSheet_(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastColumn);
}

function applyFinanceValidations_(ss) {
  var clients = ss.getSheetByName(APP.SHEETS.CLIENTS);
  var users = ss.getSheetByName(APP.SHEETS.USERS);
  var clientRule = clients ? SpreadsheetApp.newDataValidation().requireValueInRange(clients.getRange('A2:A'), true).setAllowInvalid(false).build() : null;
  var employeeRule = users ? SpreadsheetApp.newDataValidation().requireValueInRange(users.getRange('B2:B'), true).setAllowInvalid(false).build() : null;
  var accountRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PAYMENT_ACCOUNTS, true).setAllowInvalid(false).build();
  var modeRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PAYMENT_MODES, true).setAllowInvalid(false).build();
  var employeeTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EMPLOYEE_PAYMENT_TYPES, true).setAllowInvalid(false).build();
  var categoryRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EXPENSE_CATEGORIES, true).setAllowInvalid(false).build();
  var contextRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EXPENSE_CONTEXTS, true).setAllowInvalid(false).build();

  var receipts = ss.getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  if (receipts && clientRule) receipts.getRange('B2:B').setDataValidation(clientRule);
  if (receipts) { receipts.getRange('E2:E').setDataValidation(accountRule); receipts.getRange('F2:F').setDataValidation(modeRule); }

  var employees = ss.getSheetByName(APP.SHEETS.EMPLOYEE_PAYMENTS);
  if (employees) {
    if (employeeRule) employees.getRange('B2:B').setDataValidation(employeeRule);
    if (clientRule) employees.getRange('C2:C').setDataValidation(clientRule);
    employees.getRange('D2:D').setDataValidation(employeeTypeRule);
    employees.getRange('I2:I').setDataValidation(accountRule);
    employees.getRange('J2:J').setDataValidation(modeRule);
  }

  var expenses = ss.getSheetByName(APP.SHEETS.BUSINESS_EXPENSES);
  if (expenses) {
    if (clientRule) expenses.getRange('B2:B').setDataValidation(clientRule);
    expenses.getRange('C2:C').setDataValidation(categoryRule);
    expenses.getRange('D2:D').setDataValidation(contextRule);
    expenses.getRange('H2:H').setDataValidation(modeRule);
  }
}

function migrateLegacyPayments_(ss) {
  var legacy = ss.getSheetByName(APP.SHEETS.PAYMENTS);
  if (!legacy || legacy.getLastRow() < 2) {
    if (legacy && ss.getSheets().length > 1) ss.deleteSheet(legacy);
    return;
  }

  var values = legacy.getDataRange().getValues();
  var headers = values.shift();
  var index = function(name) { return headers.indexOf(name); };
  var receiptSheet = ss.getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  var employeeSheet = ss.getSheetByName(APP.SHEETS.EMPLOYEE_PAYMENTS);
  var expenseSheet = ss.getSheetByName(APP.SHEETS.BUSINESS_EXPENSES);

  values.forEach(function(r) {
    if (!r.some(function(x) { return x !== ''; })) return;
    var type = index('Type') >= 0 ? String(r[index('Type')] || '') : 'Client Receipt';
    var id = index('Payment ID') >= 0 ? r[index('Payment ID')] : '';
    var client = index('Client ID') >= 0 ? r[index('Client ID')] : '';
    var amount = index('Amount') >= 0 ? r[index('Amount')] : '';
    var date = index('Payment Date') >= 0 ? r[index('Payment Date')] : '';
    var paidTo = index('Paid To') >= 0 ? r[index('Paid To')] : '';
    var employee = index('Employee') >= 0 ? r[index('Employee')] : '';
    var category = index('Expense Category') >= 0 ? r[index('Expense Category')] : '';
    var related = index('Related To') >= 0 ? r[index('Related To')] : '';
    var description = index('Description') >= 0 ? r[index('Description')] : '';
    var status = index('Status') >= 0 ? r[index('Status')] : 'Completed';

    if (type === 'Employee Payment') {
      employeeSheet.appendRow([id || generateId_(CONFIG.ID_PREFIX.EMPLOYEE_PAYMENT), employee, client, 'Other', amount, date, '', '', paidTo, '', '', status, description]);
    } else if (type === 'Miscellaneous Expense') {
      expenseSheet.appendRow([id || generateId_(CONFIG.ID_PREFIX.BUSINESS_EXPENSE), client, category || 'Other', related || 'Other', amount, date, paidTo, '', '', description, status]);
    } else {
      receiptSheet.appendRow([id || generateId_(CONFIG.ID_PREFIX.CLIENT_RECEIPT), client, amount, date, paidTo, '', '', status || 'Received', description]);
    }
  });

  ss.deleteSheet(legacy);
}

function ensureCRMTriggers_(spreadsheet) {
  var triggers = ScriptApp.getProjectTriggers();
  var hasOpen = false, hasEdit = false;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onOpen') hasOpen = true;
    if (t.getHandlerFunction() === 'handleCRMEdit_') hasEdit = true;
  });
  if (!hasOpen) ScriptApp.newTrigger('onOpen').forSpreadsheet(spreadsheet).onOpen().create();
  if (!hasEdit) ScriptApp.newTrigger('handleCRMEdit_').forSpreadsheet(spreadsheet).onEdit().create();
}

function ensureCRMOpenTrigger_(spreadsheet) { ensureCRMTriggers_(spreadsheet); }
