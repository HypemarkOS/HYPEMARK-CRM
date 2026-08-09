function initializeCRMService_() {
  var props=PropertiesService.getScriptProperties(), ss, id=props.getProperty('CRM_SPREADSHEET_ID');
  if(id){ss=SpreadsheetApp.openById(id);}else{ss=SpreadsheetApp.getActiveSpreadsheet()||SpreadsheetApp.create(APP.NAME+' Data');props.setProperty('CRM_SPREADSHEET_ID',ss.getId());}
  var definitions=[
    [APP.SHEETS.CLIENTS,['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On']],
    [APP.SHEETS.ENGAGEMENTS,['Engagement ID','Client ID','Service','Start Date','End Date','Status']],
    [APP.SHEETS.PROJECTS,['Project ID','Client ID','Engagement ID','Project Name','Status','Created On']],
    [APP.SHEETS.DELIVERABLES,['Deliverable ID','Project ID','Type','Quantity','Status']],
    [APP.SHEETS.CONTENT_BANK,['Content ID','Deliverable ID','Title','Status','Assigned To','Publish Date']],
    [APP.SHEETS.ACTIVITIES,['Activity ID','Entity','Entity ID','Action','User','Date Time']],
    [APP.SHEETS.PAYMENTS,['Payment ID','Type','Client ID','Amount','Payment Date','Paid To','Employee','Expense Category','Related To','Description','Status']],
    [APP.SHEETS.USERS,['User ID','Name','Role','Email','Status']],
    [APP.SHEETS.SETTINGS,['Key','Value']]
  ];
  definitions.forEach(function(def){
    var s=ss.getSheetByName(def[0])||ss.insertSheet(def[0]);
    var current=s.getLastColumn()||0;
    if(s.getLastRow()===0){s.getRange(1,1,1,def[1].length).setValues([def[1]]);}else{
      var headers=s.getRange(1,1,1,current).getValues()[0];
      def[1].forEach(function(h){if(headers.indexOf(h)===-1){s.getRange(1,s.getLastColumn()+1).setValue(h);}});
    }
    s.getRange(1,1,1,s.getLastColumn()).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  ensureCRMTriggers_(ss);
  console.log('HYPEMARK CRM spreadsheet: '+ss.getUrl());
  return ss.getUrl();
}

function ensureCRMTriggers_(spreadsheet){
  var triggers=ScriptApp.getProjectTriggers();
  var hasOpen=false, hasEdit=false;
  triggers.forEach(function(t){
    if(t.getHandlerFunction()==='onOpen') hasOpen=true;
    if(t.getHandlerFunction()==='handleCRMEdit_') hasEdit=true;
  });
  if(!hasOpen) ScriptApp.newTrigger('onOpen').forSpreadsheet(spreadsheet).onOpen().create();
  if(!hasEdit) ScriptApp.newTrigger('handleCRMEdit_').forSpreadsheet(spreadsheet).onEdit().create();
}

function ensureCRMOpenTrigger_(spreadsheet){ ensureCRMTriggers_(spreadsheet); }
