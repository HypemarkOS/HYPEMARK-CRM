var PAYMENT_TYPES = {
  CLIENT_RECEIPT: 'Client Receipt',
  EMPLOYEE_PAYMENT: 'Employee Payment',
  BUSINESS_EXPENSE: 'Business Expense'
};

var RECEIPT_HEADERS = ['Receipt ID','Client ID','Amount','Payment Date','Received By','Payment Mode','Reference','Status','Notes','Allocated Amount','Remaining Balance'];
var EMPLOYEE_HEADERS = ['Employee Payment ID','Employee','Client ID','Receipt ID','Payment Type','Amount','Payment Date','Commission %','Commission Base Amount','Commission Amount','Paid From','Payment Mode','Reference','Status','Notes'];
var EXPENSE_HEADERS = ['Expense ID','Client ID','Receipt ID','Expense Category','Related To','Amount','Expense Date','Paid To','Payment Mode','Reference','Description','Status'];

function ensurePaymentSheets_() {
  var ss = getCRMSpreadsheet_();
  ensurePaymentSheet_(ss, APP.SHEETS.CLIENT_RECEIPTS, RECEIPT_HEADERS);
  ensurePaymentSheet_(ss, APP.SHEETS.EMPLOYEE_PAYMENTS, EMPLOYEE_HEADERS);
  ensurePaymentSheet_(ss, APP.SHEETS.BUSINESS_EXPENSES, EXPENSE_HEADERS);
  applyFinanceValidations_(ss);
  applyReceiptBalanceFormulas_(ss);
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
    'Notes': asText_(data.notes || data.description),
    'Allocated Amount': 0,
    'Remaining Balance': amount
  };
  appendRow_(APP.SHEETS.CLIENT_RECEIPTS, RECEIPT_HEADERS, item);
  applyReceiptBalanceFormulas_(getCRMSpreadsheet_());
  recordActivity_('Client Receipt', item['Receipt ID'], 'Recorded client receipt of ₹' + amount);
  return serializePayment_(item);
}

function createEmployeePayment_(data) {
  var employee = asText_(data.employee);
  if (!employee) throw new Error('Please select an employee or partner.');
  var allowed = (CONFIG.PAYMENT_RECIPIENTS || []).concat(getPaymentOptionsService_().employees || []);
  if (allowed.length && allowed.indexOf(employee) === -1) throw new Error('Invalid employee or partner.');
  var clientId = asText_(data.clientId);
  if (clientId && !getClientService_(clientId)) throw new Error('Client not found.');
  var receiptId = asText_(data.receiptId);
  if (receiptId && !receiptBelongsToClient_(receiptId, clientId)) throw new Error('Selected receipt does not belong to the selected client.');
  var paymentType = asText_(data.paymentType) || 'Other';
  if (CONFIG.EMPLOYEE_PAYMENT_TYPES.indexOf(paymentType) === -1) throw new Error('Invalid employee payment type.');

  var commissionPercent = data.commissionPercent === '' || data.commissionPercent == null ? '' : Number(data.commissionPercent);
  var commissionBase = data.commissionBaseAmount === '' || data.commissionBaseAmount == null ? '' : Number(data.commissionBaseAmount);
  if (commissionPercent !== '' && (!(commissionPercent >= 0) || commissionPercent > 100)) throw new Error('Commission % must be between 0 and 100.');
  if (commissionBase !== '' && commissionBase < 0) throw new Error('Commission base amount cannot be negative.');

  var amount;
  if (paymentType === 'Commission' && commissionPercent !== '' && commissionBase !== '') {
    amount = Math.round((commissionBase * commissionPercent / 100) * 100) / 100;
  } else {
    amount = positiveAmount_(data.amount);
  }
  if (receiptId) validateReceiptBalance_(receiptId, amount);

  var item = {
    'Employee Payment ID': generateId_(CONFIG.ID_PREFIX.EMPLOYEE_PAYMENT),
    'Employee': employee,
    'Client ID': clientId,
    'Receipt ID': receiptId,
    'Payment Type': paymentType,
    'Amount': amount,
    'Payment Date': data.paymentDate ? new Date(data.paymentDate) : now_(),
    'Commission %': commissionPercent,
    'Commission Base Amount': commissionBase,
    'Commission Amount': paymentType === 'Commission' ? amount : '',
    'Paid From': asText_(data.paidFrom || data.paidTo),
    'Payment Mode': asText_(data.paymentMode),
    'Reference': asText_(data.reference),
    'Status': asText_(data.status) || 'Paid',
    'Notes': asText_(data.notes || data.description)
  };
  appendRow_(APP.SHEETS.EMPLOYEE_PAYMENTS, EMPLOYEE_HEADERS, item);
  applyReceiptBalanceFormulas_(getCRMSpreadsheet_());
  recordActivity_('Employee Payment', item['Employee Payment ID'], 'Recorded ' + paymentType + ' payment of ₹' + amount + ' to ' + employee);
  return serializePayment_(item);
}

function createBusinessExpense_(data) {
  var clientId = asText_(data.clientId);
  if (clientId && !getClientService_(clientId)) throw new Error('Client not found.');
  var receiptId = asText_(data.receiptId);
  if (receiptId && !receiptBelongsToClient_(receiptId, clientId)) throw new Error('Selected receipt does not belong to the selected client.');
  var category = asText_(data.expenseCategory) || 'Other';
  var relatedTo = asText_(data.relatedTo) || 'Other';
  if (CONFIG.EXPENSE_CATEGORIES.indexOf(category) === -1) throw new Error('Invalid expense category.');
  if (CONFIG.EXPENSE_CONTEXTS.indexOf(relatedTo) === -1) throw new Error('Invalid expense context.');
  var amount = positiveAmount_(data.amount);
  if (receiptId) validateReceiptBalance_(receiptId, amount);
  var item = {
    'Expense ID': generateId_(CONFIG.ID_PREFIX.BUSINESS_EXPENSE),
    'Client ID': clientId,
    'Receipt ID': receiptId,
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
  applyReceiptBalanceFormulas_(getCRMSpreadsheet_());
  recordActivity_('Business Expense', item['Expense ID'], 'Recorded ' + category + ' expense of ₹' + amount);
  return serializePayment_(item);
}

function receiptBelongsToClient_(receiptId, clientId) {
  var s = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  if (!s || s.getLastRow() < 2) return false;
  var v = s.getDataRange().getValues(), h = v.shift(), idIdx=h.indexOf('Receipt ID'), clientIdx=h.indexOf('Client ID');
  for (var i=0;i<v.length;i++) if (String(v[i][idIdx])===String(receiptId)) return !clientId || String(v[i][clientIdx])===String(clientId);
  return false;
}

function validateReceiptBalance_(receiptId, amount) {
  var s = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  if (!s || s.getLastRow() < 2) throw new Error('Receipt not found.');
  var v=s.getDataRange().getValues(),h=v.shift(),idIdx=h.indexOf('Receipt ID'),amountIdx=h.indexOf('Amount'),allocIdx=h.indexOf('Allocated Amount');
  for(var i=0;i<v.length;i++) if(String(v[i][idIdx])===String(receiptId)) {
    var allocated=allocIdx>=0?Number(v[i][allocIdx]||0):0, total=Number(v[i][amountIdx]||0), remaining=total-allocated;
    if(amount>remaining+0.001) throw new Error('This payment exceeds the remaining balance of receipt ' + receiptId + ' (₹' + remaining.toFixed(2) + ').');
    return true;
  }
  throw new Error('Receipt not found.');
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
  var employees = (CONFIG.PAYMENT_RECIPIENTS || []).slice();
  if (users && users.getLastRow() >= 2) {
    users.getRange(2,2,users.getLastRow()-1,1).getValues().forEach(function(r){var n=String(r[0]||'').trim();if(n&&employees.indexOf(n)<0)employees.push(n);});
  }
  return {
    types: [PAYMENT_TYPES.CLIENT_RECEIPT, PAYMENT_TYPES.EMPLOYEE_PAYMENT, PAYMENT_TYPES.BUSINESS_EXPENSE],
    employees: employees,
    paymentAccounts: CONFIG.PAYMENT_ACCOUNTS,
    employeePaymentTypes: CONFIG.EMPLOYEE_PAYMENT_TYPES,
    expenseCategories: CONFIG.EXPENSE_CATEGORIES,
    expenseContexts: CONFIG.EXPENSE_CONTEXTS,
    paymentModes: CONFIG.PAYMENT_MODES,
    commissionEnabled: true
  };
}

function getReceiptOptionsService_() {
  var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  if(!s||s.getLastRow()<2)return [];
  var v=s.getDataRange().getValues(),h=v.shift(),id=h.indexOf('Receipt ID'),client=h.indexOf('Client ID'),amount=h.indexOf('Amount'),remain=h.indexOf('Remaining Balance');
  return v.filter(function(r){return r[id]&&Number(r[remain]||0)>0;}).map(function(r){return {receiptId:String(r[id]),clientId:String(r[client]),amount:Number(r[amount]||0),remaining:Number(r[remain]||0)};});
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
