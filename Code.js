/**
 * HYPEMARK CRM v1
 * Public Apps Script entry points.
 */
function doGet() {
  return HtmlService.createHtmlOutput('<h2>HYPEMARK CRM</h2><p>CRM is ready.</p>').setTitle('HYPEMARK CRM');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 HYPEMARK CRM')
    .addItem('➕ Add Client', 'showAddClient')
    .addSeparator()
    .addItem('⚙ Initialize CRM', 'initializeCRM')
    .addToUi();
}

function showAddClient() {
  var html = HtmlService.createHtmlOutputFromFile('AddClient')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Client');
}

function initializeCRM() { return initializeCRMService_(); }
function createClient(data) { return createClientService_(data); }
function getClients() { return getClientsService_(); }
