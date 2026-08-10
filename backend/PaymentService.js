var PAYMENT_TYPES = {
  CLIENT_RECEIPT: 'Client Receipt',
  EMPLOYEE_PAYMENT: 'Employee Payment',
  BUSINESS_EXPENSE: 'Business Expense'
};

var RECEIPT_HEADERS = ['Receipt ID','Client ID','Amount','Payment Date','Received By','Payment Mode','Reference','Status','Notes'];
var EMPLOYEE_HEADERS = ['Employee Payment ID','Employee','Client ID','Payment Type','Amount','Payment Date','Commission %','Commission Base Amount','Paid From','Payment Mode','Reference','Status','Notes'];
var EXPENSE_HEADERS = ['Expense ID','Client ID','Expense Category','Related To','Amount','Expense Date','Paid To','Payment Mode','Reference','Description','Status'];

function ensurePaymentSheets_() {
  var ss = getCRMSpreadsheet_();
  ensurePaymentSheet_(ss, APP.SHEETS.CLIENT_RECEIPTS, RECEIPT_HEADERS);
  ensurePaymentSheet_(ss, APP.SHEETS.EMPLOYEE_PAYMENTS, EMPLOYEE_HEADERS);
  ensurePaymentSheet_(ss, APP.SHEETS.BUSINESS_EXPENSES, EXPENSE_HEADERS);
  applyFinanceValidations_(ss);
}

function ensurePaymentSheet_(ss, name, headers) {
  var s = ss.getSheetByName(name) || ss.insertSheet(name);
  if (s.getLastRow() === 0) s.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    var current = s.getRange(1,1,1,s.getLastColumn() || headers.length).getValues()[0];
    headers.forEach(function(h){if(current.indexOf(h)===-1)s.getRange(1,s.getLastColumn()+1).setValue(h);});
  }
  s.getRange(1,1,1,s.getLastColumn()).setFontWeight('bold');
  s.setFrozenRows(1);
  return s;
}

function createPaymentService_(data) {
  if (!data) throw new Error('Payment data is required.');
  ensurePaymentSheets_();
  var type = asText_(data.type) || PAYMENT_TYPES.CLIENT_RECEIPT;
  if (type === 'Miscellaneous Expense') type = PAYMENT_TYPES.BUSINESS_EXPENSE;

  if (type === PAYMENT_TYPES.CLIENT_RECEIPT) return createClientReceipt_(data);
  if (type === PAYMENT_TYPES.EMPLOYEE_PAYMENT) return createEmployeePayment_(data);
  if (type === PAYMENT_TYPES.BUSINESS_EXPENSE) return createBusinessExpense_(data);
  throw new Error('Invalid payment type.');
}

function createClientReceipt_(data) {
  var clientId = asText_(data.clientId);
  if (!clientId || !getClientService_(clientId)) throw new Error('Please select a valid client.');
  var amount = positiveAmount_(data.amount);
  var item = {
    'Receipt ID': generateId_(CONFIG.ID_PREFIX.CLIENT_RECEIPT),
    'Client ID': clientId,
    'Amount': amount,
    'Payment Date': data.paymentDate ? new Date(data.paymentDate) : now_(),
    'Received By': asText_(data.receivedBy || data.paidTo),
    'Payment Mode': asText_(data.paymentMode),
    'Reference': asText_(data.reference),
    'Status': asText_(data.status) || 'Received',
    'Notes': asText_(data.notes || data.description)
  };
  appendRow_(APP.SHEETS.CLIENT_RECEIPTS, RECEIPT_HEADERS, item);
  recordActivity_('Client Receipt', item['Receipt ID'], 'Recorded client receipt of ₹' + amount);
  return serializePayment_(item);
}

function createEmployeePayment_(data) {
  var employee = asText_(data.employee);
  if (!employee) throw new Error('Please select an employee.');
  var clientId = asText_(data.clientId);
  if (clientId && !getClientService_(clientId)) throw new Error('Client not found.');
  var paymentType = asText_(data.paymentType) || 'Other';
  if (CONFIG.EMPLOYEE_PAYMENT_TYPES.indexOf(paymentType) === -1) throw new Error('Invalid employee payment type.');
  var amount = positiveAmount_(data.amount);
  var commissionPercent = data.commissionPercent === '' || data.commissionPercent == null ? '' : Number(data.commissionPercent);
  var commissionBase = data.commissionBaseAmount === '' || data.commissionBaseAmount == null ? '' : Number(data.commissionBaseAmount);
  if (paymentType === 'Commission' && commissionPercent !== '' && commissionPercent < 0) throw new Error('Commission % cannot be negative.');
  var item = {
    'Employee Payment ID': generateId_(CONFIG.ID_PREFIX.EMPLOYEE_PAYMENT),
    'Employee': employee,
    'Client ID': clientId,
    'Payment Type': paymentType,
    'Amount': amount,
    'Payment Date': data.paymentDate ? new Date(data.paymentDate) : now_(),
    'Commission %': commissionPercent,
    'Commission Base Amount': commissionBase,
    'Paid From': asText_(data.paidFrom || data.paidTo),
    'Payment Mode': asText_(data.paymentMode),
    'Reference': asText_(data.reference),
    'Status': asText_(data.status) || 'Paid',
    'Notes': asText_(data.notes || data.description)
  };
  appendRow_(APP.SHEETS.EMPLOYEE_PAYMENTS, EMPLOYEE_HEADERS, item);
  recordActivity_('Employee Payment', item['Employee Payment ID'], 'Recorded ' + paymentType + ' payment of ₹' + amount + ' to ' + employee);
  return serializePayment_(item);
}

function createBusinessExpense_(data) {
  var clientId = asText_(data.clientId);
  if (clientId && !getClientService_(clientId)) throw new Error('Client not found.');
  var category = asText_(data.expenseCategory) || 'Other';
  var relatedTo = asText_(data.relatedTo) || 'Other';
  if (CONFIG.EXPENSE_CATEGORIES.indexOf(category) === -1) throw new Error('Invalid expense category.');
  if (CONFIG.EXPENSE_CONTEXTS.indexOf(relatedTo) === -1) throw new Error('Invalid expense context.');
  var amount = positiveAmount_(data.amount);
  var item = {
    'Expense ID': generateId_(CONFIG.ID_PREFIX.BUSINESS_EXPENSE),
    'Client ID': clientId,
    'Expense Category': category,
    'Related To': relatedTo,
    'Amount': amount,
    'Expense Date': data.paymentDate ? new Date(data.paymentDate) : now_(),
    'Paid To': asText_(data.paidTo),
    'Payment Mode': asText_(data.paymentMode),
    'Reference': asText_(data.reference),
    'Description': asText_(data.description),
    'Status': asText_(data.status) || 'Paid'
  };
  appendRow_(APP.SHEETS.BUSINESS_EXPENSES, EXPENSE_HEADERS, item);
  recordActivity_('Business Expense', item['Expense ID'], 'Recorded ' + category + ' expense of ₹' + amount);
  return serializePayment_(item);
}

function positiveAmount_(value) {
  var amount = Number(value);
  if (!(amount > 0)) throw new Error('Amount must be greater than zero.');
  return amount;
}

function getPaymentsService_(clientId, type) {
  ensurePaymentSheets_();
  var all = [];
  if (!type || type === PAYMENT_TYPES.CLIENT_RECEIPT) all = all.concat(readPaymentSheet_(APP.SHEETS.CLIENT_RECEIPTS, RECEIPT_HEADERS, PAYMENT_TYPES.CLIENT_RECEIPT, clientId));
  if (!type || type === PAYMENT_TYPES.EMPLOYEE_PAYMENT) all = all.concat(readPaymentSheet_(APP.SHEETS.EMPLOYEE_PAYMENTS, EMPLOYEE_HEADERS, PAYMENT_TYPES.EMPLOYEE_PAYMENT, clientId));
  if (!type || type === PAYMENT_TYPES.BUSINESS_EXPENSE || type === 'Miscellaneous Expense') all = all.concat(readPaymentSheet_(APP.SHEETS.BUSINESS_EXPENSES, EXPENSE_HEADERS, PAYMENT_TYPES.BUSINESS_EXPENSE, clientId));
  return all;
}

function readPaymentSheet_(sheetName, headers, type, clientId) {
  var s = getCRMSpreadsheet_().getSheetByName(sheetName);
  if (!s || s.getLastRow() < 2) return [];
  var values = s.getDataRange().getValues();
  var actualHeaders = values.shift();
  return values.filter(function(r){return r.some(function(x){return x!=='';});}).filter(function(r){
    var idx = actualHeaders.indexOf('Client ID');
    return !clientId || (idx >= 0 && String(r[idx]) === String(clientId));
  }).map(function(r){
    var o = {'Type': type};
    actualHeaders.forEach(function(h,i){o[h]=r[i] instanceof Date?Utilities.formatDate(r[i],CONFIG.TIMEZONE,'dd-MMM-yyyy HH:mm'):r[i];});
    return o;
  });
}

function getPaymentOptionsService_() {
  var users = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  var employees = [];
  if (users && users.getLastRow() >= 2) {
    employees = users.getRange(2,2,users.getLastRow()-1,1).getValues().map(function(r){return String(r[0]||'').trim();}).filter(Boolean);
  }
  return {
    types: [PAYMENT_TYPES.CLIENT_RECEIPT, PAYMENT_TYPES.EMPLOYEE_PAYMENT, PAYMENT_TYPES.BUSINESS_EXPENSE],
    employees: employees,
    paymentAccounts: CONFIG.PAYMENT_ACCOUNTS,
    employeePaymentTypes: CONFIG.EMPLOYEE_PAYMENT_TYPES,
    expenseCategories: CONFIG.EXPENSE_CATEGORIES,
    expenseContexts: CONFIG.EXPENSE_CONTEXTS,
    paymentModes: CONFIG.PAYMENT_MODES
  };
}

function getClientProfitabilityService_(clientId) {
  var payments = getPaymentsService_(clientId), receipts = 0, employee = 0, misc = 0;
  payments.forEach(function(p){
    var a = Number(p.Amount || 0);
    if (p.Type === PAYMENT_TYPES.CLIENT_RECEIPT) receipts += a;
    if (p.Type === PAYMENT_TYPES.EMPLOYEE_PAYMENT) employee += a;
    if (p.Type === PAYMENT_TYPES.BUSINESS_EXPENSE) misc += a;
  });
  return {clientId: clientId, receipts: receipts, employeePayments: employee, miscExpenses: misc, totalCosts: employee + misc, profit: receipts - employee - misc};
}

function getProfitabilitySummaryService_() {
  return getClientsService_().map(function(c){
    return {clientId:c['Client ID'], clientName:c['Client Name'], businessName:c['Business Name'], profitability:getClientProfitabilityService_(c['Client ID'])};
  });
}

function serializePayment_(item) {
  var copy = {};
  Object.keys(item).forEach(function(k){copy[k] = item[k] instanceof Date ? Utilities.formatDate(item[k], CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm') : item[k];});
  return copy;
}
