var PAYMENT_HEADERS = ['Payment ID','Type','Client ID','Amount','Payment Date','Paid To','Employee','Expense Category','Related To','Description','Status'];
var PAYMENT_TYPES = { CLIENT_RECEIPT:'Client Receipt', EMPLOYEE_PAYMENT:'Employee Payment', MISC_EXPENSE:'Miscellaneous Expense' };
var EXPENSE_CATEGORIES = ['Petrol','Camera Rental','Transport','Food','Accommodation','Equipment','Other'];
var EXPENSE_CONTEXTS = ['Shoot','Client Meeting','Travel','Office','Other'];

function ensurePaymentHeaders_(){
  var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.PAYMENTS); if(!s)return null;
  var current=s.getLastColumn()||0;
  if(s.getLastRow()===0){s.getRange(1,1,1,PAYMENT_HEADERS.length).setValues([PAYMENT_HEADERS]);}
  else {var headers=s.getRange(1,1,1,current).getValues()[0]; PAYMENT_HEADERS.forEach(function(h){if(headers.indexOf(h)===-1)s.getRange(1,s.getLastColumn()+1).setValue(h);});}
  s.getRange(1,1,1,s.getLastColumn()).setFontWeight('bold'); s.setFrozenRows(1); return s;
}

function createPaymentService_(data){
  if(!data)throw new Error('Payment data is required.');
  var type=asText_(data.type)||PAYMENT_TYPES.CLIENT_RECEIPT;
  if([PAYMENT_TYPES.CLIENT_RECEIPT,PAYMENT_TYPES.EMPLOYEE_PAYMENT,PAYMENT_TYPES.MISC_EXPENSE].indexOf(type)===-1)throw new Error('Invalid payment type.');
  var amount=Number(data.amount); if(!(amount>0))throw new Error('Amount must be greater than zero.');
  var clientId=asText_(data.clientId);
  if(type===PAYMENT_TYPES.CLIENT_RECEIPT&&!clientId)throw new Error('Client is required for a client receipt.');
  if(clientId&&!getClientService_(clientId))throw new Error('Client not found.');
  var category=asText_(data.expenseCategory), relatedTo=asText_(data.relatedTo);
  if(type===PAYMENT_TYPES.MISC_EXPENSE&&category&&EXPENSE_CATEGORIES.indexOf(category)===-1)throw new Error('Invalid expense category.');
  if(type===PAYMENT_TYPES.MISC_EXPENSE&&relatedTo&&EXPENSE_CONTEXTS.indexOf(relatedTo)===-1)throw new Error('Invalid expense context.');
  ensurePaymentHeaders_();
  var item={'Payment ID':generateId_(CONFIG.ID_PREFIX.PAYMENT),'Type':type,'Client ID':clientId,'Amount':amount,'Payment Date':data.paymentDate?new Date(data.paymentDate):now_(),'Paid To':asText_(data.paidTo),'Employee':asText_(data.employee),'Expense Category':category,'Related To':relatedTo,'Description':asText_(data.description),'Status':asText_(data.status)||'Completed'};
  appendRow_(APP.SHEETS.PAYMENTS,PAYMENT_HEADERS,item); recordActivity_('Payment',item['Payment ID'],'Recorded '+type); return serializePayment_(item);
}

function getPaymentsService_(clientId,type){
  var s=ensurePaymentHeaders_(); if(!s||s.getLastRow()<2)return [];
  var v=s.getDataRange().getValues(),h=v.shift();
  return v.filter(function(r){if(!r.some(function(x){return x!=='';}))return false;var c=!clientId||String(r[h.indexOf('Client ID')])===String(clientId);var t=!type||String(r[h.indexOf('Type')])===String(type);return c&&t;}).map(function(r){var o={};h.forEach(function(k,i){o[k]=r[i]instanceof Date?Utilities.formatDate(r[i],CONFIG.TIMEZONE,'dd-MMM-yyyy'):r[i];});return o;});
}

function getPaymentOptionsService_(){return {types:[PAYMENT_TYPES.CLIENT_RECEIPT,PAYMENT_TYPES.EMPLOYEE_PAYMENT,PAYMENT_TYPES.MISC_EXPENSE],expenseCategories:EXPENSE_CATEGORIES,expenseContexts:EXPENSE_CONTEXTS};}

function getClientProfitabilityService_(clientId){
  var payments=getPaymentsService_(clientId),receipts=0,employee=0,misc=0;
  payments.forEach(function(p){var a=Number(p['Amount']||0);if(p['Type']===PAYMENT_TYPES.CLIENT_RECEIPT)receipts+=a;if(p['Type']===PAYMENT_TYPES.EMPLOYEE_PAYMENT)employee+=a;if(p['Type']===PAYMENT_TYPES.MISC_EXPENSE)misc+=a;});
  return {clientId:clientId,receipts:receipts,employeePayments:employee,miscExpenses:misc,totalCosts:employee+misc,profit:receipts-employee-misc};
}
function getProfitabilitySummaryService_(){return getClientsService_().map(function(c){return {clientId:c['Client ID'],clientName:c['Client Name'],businessName:c['Business Name'],profitability:getClientProfitabilityService_(c['Client ID'])};});}
function serializePayment_(item){var copy={};Object.keys(item).forEach(function(k){copy[k]=item[k]instanceof Date?Utilities.formatDate(item[k],CONFIG.TIMEZONE,'dd-MMM-yyyy HH:mm'):item[k];});return copy;}
