/**
 * HYPEMARK CRM v1
 * Public Apps Script entry points.
 * Keep these functions in the root file so the Apps Script editor exposes them.
 */

function doGet() {
  return HtmlService.createHtmlOutput('<h2>HYPEMARK CRM</h2><p>CRM is ready.</p>')
    .setTitle('HYPEMARK CRM');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 HYPEMARK CRM')
    .addItem('➕ Add Client', 'showAddClient')
    .addSeparator()
    .addItem('⚙ Initialize CRM', 'initializeCRM')
    .addToUi();
}

function showAddClient() {
  var html = HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:24px"><h2>Add New Client</h2><p>Client form is ready for the next module.</p></div>')
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Client');
}

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

  definitions.forEach(function(definition) {
    var sheet = ss.getSheetByName(definition[0]) || ss.insertSheet(definition[0]);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, definition[1].length).setValues([definition[1]]);
    }
    sheet.getRange(1, 1, 1, definition[1].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  SpreadsheetApp.getUi().alert('HYPEMARK CRM initialized successfully.');
  return true;
}

function createClient(data) {
  return createClientService_(data);
}

function getClients() {
  return getClientsService_();
}
