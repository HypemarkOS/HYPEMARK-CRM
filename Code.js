/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Main Menu
 * ==========================================================
 */

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu("HYPEMARK CRM")

    .addItem("📊 Dashboard", "showDashboard")

    .addSeparator()

    .addItem("➕ Add Client", "showAddClientForm")
    .addItem("👥 View Clients", "viewClients")

    .addSeparator()

    .addItem("🔄 Refresh CRM", "initializeCRM")

    .addToUi();

}