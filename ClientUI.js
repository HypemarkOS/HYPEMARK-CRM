/**
 * ==========================================================
 * HYPEMARK CRM v1.0
 * Client UI
 * ==========================================================
 */

function showAddClientForm() {

  const html = HtmlService
    .createHtmlOutputFromFile("AddClient")
    .setWidth(1200)
    .setHeight(750);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Client Onboarding"
    );

}


/**
 * Temporary function
 * Will be implemented later
 */
function viewClients() {

  SpreadsheetApp.getUi().alert("View Clients - Coming Soon");

}


/**
 * Temporary Dashboard
 */
function showDashboard() {

  SpreadsheetApp.getUi().alert("Dashboard - Coming Soon");

}