function initializeCRM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var definitions = [
    [APP.SHEETS.CLIENTS, ['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On']],
    [APP.SHEETS.ENGAGEMENTS, ['Engagement ID','Client ID','Service','Start Date','End Date','Status']],
    [APP.SHEETS.PROJECTS, ['Project ID','Client ID','Engagement ID','Project Name','Status','Created On']],
    [APP.SHEETS.DELIVERABLES, ['Deliverable ID','Project ID','Type','Quantity','Status']],
    [APP.SHEETS.CONTENT_BANK, ['Content ID','Deliverable ID','Title','Status','Assigned To','Publish Date']],
    [APP.SHEETS.ACTIVITIES, ['Activity ID','Entity','Entity ID','Action','User','Date Time']],
    [APP.SHEETS.PAYMENTS, ['Payment ID','Client ID','Amount','Payment Date','Status']],
    [APP.SHEETS.USERS, ['User ID','Name','Role','Email','Status']],
    [APP.SHEETS.SETTINGS, ['Key','Value']]
  ];
  definitions.forEach(function(def) {
    var sheet = ss.getSheetByName(def[0]) || ss.insertSheet(def[0]);
    if (sheet.getLastRow() === 0) sheet.getRange(1,1,1,def[1].length).setValues([def[1]]);
    sheet.getRange(1,1,1,def[1].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert('HYPEMARK CRM initialized successfully.');
}
