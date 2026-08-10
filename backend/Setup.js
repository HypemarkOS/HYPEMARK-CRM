function initializeCRMService_() {
  var props = PropertiesService.getScriptProperties();
  var ss, id = props.getProperty('CRM_SPREADSHEET_ID');
  if (id) ss = SpreadsheetApp.openById(id);
  else { ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create(APP.NAME + ' Data'); props.setProperty('CRM_SPREADSHEET_ID', ss.getId()); }

  var definitions = [
    [APP.SHEETS.CLIENTS, ['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On']],
    [APP.SHEETS.ENGAGEMENTS, ['Engagement ID','Client ID','Service','Start Date','End Date','Status']],
    [APP.SHEETS.PROJECTS, ['Project ID','Client ID','Engagement ID','Project Name','Status','Created On']],
    [APP.SHEETS.DELIVERABLES, ['Deliverable ID','Project ID','Type','Quantity','Status']],
    [APP.SHEETS.CONTENT_BANK, ['Content ID','Deliverable ID','Title','Status','Assigned To','Publish Date']],
    [APP.SHEETS.ACTIVITIES, ['Activity ID','Entity','Entity ID','Action','User','Date Time']],
    [APP.SHEETS.CLIENT_RECEIPTS, ['Receipt ID','Client ID','Amount','Payment Date','Received By','Payment Mode','Reference','Status','Notes','Allocated Amount','Remaining Balance']],
    [APP.SHEETS.EMPLOYEE_PAYMENTS, ['Employee Payment ID','Employee','Client ID','Receipt ID','Payment Type','Amount','Payment Date','Commission %','Commission Base Amount','Commission Amount','Paid From','Payment Mode','Reference','Status','Notes']],
    [APP.SHEETS.BUSINESS_EXPENSES, ['Expense ID','Client ID','Receipt ID','Expense Category','Related To','Amount','Expense Date','Paid By','Paid To','Payment Mode','Reference','Description','Status']],
    [APP.SHEETS.USERS, ['User ID','Name','Role','Email','Status']],
    [APP.SHEETS.SETTINGS, ['Key','Value']]
  ];

  definitions.forEach(function(def) { var s=ss.getSheetByName(def[0])||ss.insertSheet(def[0]); ensureHeaders_(s,def[1]); styleSheet_(s); });
  migrateLegacyPayments_(ss);
  applyFinanceValidations_(ss);
  applyReceiptBalanceFormulas_(ss);
  ensureCRMTriggers_(ss);

  var defaultSheet=ss.getSheetByName('Sheet1');
  if(defaultSheet&&ss.getSheets().length>1&&defaultSheet.getLastRow()===0&&defaultSheet.getLastColumn()===1) ss.deleteSheet(defaultSheet);
  console.log('HYPEMARK CRM spreadsheet: '+ss.getUrl());
  return ss.getUrl();
}

function ensureHeaders_(sheet,headers){
  var current=sheet.getLastColumn()||0;
  if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);return;}
  var existing=sheet.getRange(1,1,1,current).getValues()[0];
  headers.forEach(function(h){if(existing.indexOf(h)===-1)sheet.getRange(1,sheet.getLastColumn()+1).setValue(h);});
}
function styleSheet_(sheet){var lastColumn=sheet.getLastColumn();if(!lastColumn)return;sheet.getRange(1,1,1,lastColumn).setFontWeight('bold');sheet.setFrozenRows(1);sheet.autoResizeColumns(1,lastColumn);}

function applyFinanceValidations_(ss){
  var clients=ss.getSheetByName(APP.SHEETS.CLIENTS), users=ss.getSheetByName(APP.SHEETS.USERS), receipts=ss.getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);
  var clientRule=clients?SpreadsheetApp.newDataValidation().requireValueInRange(clients.getRange('A2:A'),true).setAllowInvalid(false).build():null;
  var employeeValues=(CONFIG.PAYMENT_RECIPIENTS||[]).slice();
  if(users&&users.getLastRow()>=2) users.getRange(2,2,users.getLastRow()-1,1).getValues().forEach(function(r){var n=String(r[0]||'').trim();if(n&&employeeValues.indexOf(n)<0)employeeValues.push(n);});
  var employeeRule=SpreadsheetApp.newDataValidation().requireValueInList(employeeValues,true).setAllowInvalid(false).build();
  var receiptRule=receipts?SpreadsheetApp.newDataValidation().requireValueInRange(receipts.getRange('A2:A'),true).setAllowInvalid(false).build():null;
  var accountRule=SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PAYMENT_ACCOUNTS,true).setAllowInvalid(false).build();
  var modeRule=SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PAYMENT_MODES,true).setAllowInvalid(false).build();
  var employeeTypeRule=SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EMPLOYEE_PAYMENT_TYPES,true).setAllowInvalid(false).build();
  var categoryRule=SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EXPENSE_CATEGORIES,true).setAllowInvalid(false).build();
  var contextRule=SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.EXPENSE_CONTEXTS,true).setAllowInvalid(false).build();
  if(receipts){if(clientRule)receipts.getRange('B2:B').setDataValidation(clientRule);receipts.getRange('E2:E').setDataValidation(accountRule);receipts.getRange('F2:F').setDataValidation(modeRule);}
  var employees=ss.getSheetByName(APP.SHEETS.EMPLOYEE_PAYMENTS);
  if(employees){employees.getRange('B2:B').setDataValidation(employeeRule);if(clientRule)employees.getRange('C2:C').setDataValidation(clientRule);if(receiptRule)employees.getRange('D2:D').setDataValidation(receiptRule);employees.getRange('E2:E').setDataValidation(employeeTypeRule);employees.getRange('K2:K').setDataValidation(accountRule);employees.getRange('L2:L').setDataValidation(modeRule);}
  var expenses=ss.getSheetByName(APP.SHEETS.BUSINESS_EXPENSES);
  if(expenses){if(clientRule)expenses.getRange('B2:B').setDataValidation(clientRule);if(receiptRule)expenses.getRange('C2:C').setDataValidation(receiptRule);expenses.getRange('D2:D').setDataValidation(categoryRule);expenses.getRange('E2:E').setDataValidation(contextRule);expenses.getRange('H2:H').setDataValidation(employeeRule);expenses.getRange('J2:J').setDataValidation(modeRule);}
}

function applyReceiptBalanceFormulas_(ss){
  var receipts=ss.getSheetByName(APP.SHEETS.CLIENT_RECEIPTS);if(!receipts)return;var lastRow=Math.max(receipts.getLastRow(),2);
  receipts.getRange('J1:K1').setValues([['Allocated Amount','Remaining Balance']]);
  receipts.getRange('J2').setFormula('=IF(A2="","",SUMIF(EmployeePayments!D:D,A2,EmployeePayments!F:F)+SUMIF(BusinessExpenses!C:C,A2,BusinessExpenses!F:F))');
  receipts.getRange('K2').setFormula('=IF(A2="","",C2-J2)');
  if(lastRow>2)receipts.getRange('J2:K2').copyTo(receipts.getRange(2,10,lastRow-1,2));
  receipts.getRange('C2:C').setNumberFormat('₹#,##0.00');receipts.getRange('J2:K').setNumberFormat('₹#,##0.00');
}

function migrateLegacyPayments_(ss){
  var legacy=ss.getSheetByName(APP.SHEETS.PAYMENTS);if(!legacy||legacy.getLastRow()<2){if(legacy&&ss.getSheets().length>1)ss.deleteSheet(legacy);return;}
  var values=legacy.getDataRange().getValues(),headers=values.shift(),index=function(name){return headers.indexOf(name);};
  var receiptSheet=ss.getSheetByName(APP.SHEETS.CLIENT_RECEIPTS),employeeSheet=ss.getSheetByName(APP.SHEETS.EMPLOYEE_PAYMENTS),expenseSheet=ss.getSheetByName(APP.SHEETS.BUSINESS_EXPENSES);
  values.forEach(function(r){if(!r.some(function(x){return x!=='';}))return;var type=index('Type')>=0?String(r[index('Type')]||''):'Client Receipt',id=index('Payment ID')>=0?r[index('Payment ID')]:'',client=index('Client ID')>=0?r[index('Client ID')]:'',amount=index('Amount')>=0?r[index('Amount'):''],date=index('Payment Date')>=0?r[index('Payment Date')]:'',paidTo=index('Paid To')>=0?r[index('Paid To')]:'',employee=index('Employee')>=0?r[index('Employee')]:'',category=index('Expense Category')>=0?r[index('Expense Category')]:'',related=index('Related To')>=0?r[index('Related To')]:'',description=index('Description')>=0?r[index('Description')]:'',status=index('Status')>=0?r[index('Status')]:'Completed';
    if(type==='Employee Payment')employeeSheet.appendRow([id||generateId_(CONFIG.ID_PREFIX.EMPLOYEE_PAYMENT),employee||'Ajay',client,'','Other',amount,date,'','','',paidTo,'','','',status,description]);
    else if(type==='Miscellaneous Expense'||type==='Business Expense')expenseSheet.appendRow([id||generateId_(CONFIG.ID_PREFIX.BUSINESS_EXPENSE),client,'',category||'Other',related||'Other',amount,date,'',paidTo,'','','',status]);
    else receiptSheet.appendRow([id||generateId_(CONFIG.ID_PREFIX.CLIENT_RECEIPT),client,amount,date,paidTo,'','',status||'Received',description,'','']);
  });
  ss.deleteSheet(legacy);
}
function ensureCRMTriggers_(spreadsheet){var triggers=ScriptApp.getProjectTriggers(),hasOpen=false,hasEdit=false;triggers.forEach(function(t){if(t.getHandlerFunction()==='onOpen')hasOpen=true;if(t.getHandlerFunction()==='handleCRMEdit_')hasEdit=true;});if(!hasOpen)ScriptApp.newTrigger('onOpen').forSpreadsheet(spreadsheet).onOpen().create();if(!hasEdit)ScriptApp.newTrigger('handleCRMEdit_').forSpreadsheet(spreadsheet).onEdit().create();}
function ensureCRMOpenTrigger_(spreadsheet){ensureCRMTriggers_(spreadsheet);}
