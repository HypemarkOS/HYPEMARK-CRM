/**
 * HYPEMARK CRM v1
 * Public Apps Script entry points.
 */
function doGet(e) {
  var page = e && e.parameter ? e.parameter.page : '';
  var file = page === 'add' ? 'AddClient' : 'Index';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(page === 'add' ? 'Add Client | HYPEMARK CRM' : 'HYPEMARK CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 HYPEMARK CRM')
    .addItem('➕ Add Client', 'showAddClient')
    .addSeparator()
    .addItem('⚙ Initialize CRM', 'initializeCRM')
    .addToUi();
}

function showAddClient() {
  var html = HtmlService.createTemplateFromFile('AddClient')
    .evaluate()
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Client');
}

function initializeCRM() { return initializeCRMService_(); }
function createClient(data) { return createClientService_(data); }
function getClients() { return getClientsService_(); }
function getClient(clientId) { return getClientService_(clientId); }
function updateClient(clientId, data) { return updateClientService_(clientId, data); }
function createEngagement(data) { return createEngagementService_(data); }
function getEngagements(clientId) { return getEngagementsService_(clientId); }
function createProject(data) { return createProjectService_(data); }
function getProjects(clientId) { return getProjectsService_(clientId); }
function getProject(projectId) { return getProjectService_(projectId); }
