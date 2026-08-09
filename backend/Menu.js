function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 HYPEMARK CRM')
    .addItem('➕ Add Client', 'showAddClient')
    .addSeparator()
    .addItem('⚙ Initialize CRM', 'initializeCRM')
    .addToUi();
}
function showAddClient() {
  var html = HtmlService.createHtmlOutputFromFile('frontend/AddClient').setWidth(1100).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Client');
}
