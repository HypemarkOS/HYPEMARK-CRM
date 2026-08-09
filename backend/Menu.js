function buildCrmMenu_() {
  return SpreadsheetApp.getUi()
    .createMenu('🚀 HYPEMARK CRM')
    .addItem('➕ Add Client', 'showAddClient')
    .addSeparator()
    .addItem('⚙ Initialize CRM', 'initializeCRM');
}
