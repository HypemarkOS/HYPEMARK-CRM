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
  return initializeCRMService_();
}

function createClient(data) {
  return createClientService_(data);
}

function getClients() {
  return getClientsService_();
}
