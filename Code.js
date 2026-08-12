/**
 * HYPEMARK CRM v1
 * Stable single-shell entry point. Critical brand assets are bundled, not loaded at runtime from Drive.
 */
function doGet(e) {
  var auth=getAuthContextService_();
  var file=auth.authenticated?'App':'Login';
  var output=HtmlService.createTemplateFromFile(file).evaluate();
  if(auth.authenticated){
    var theme=HtmlService.createHtmlOutputFromFile('PremiumTheme').getContent();
    var brandFix=HtmlService.createHtmlOutputFromFile('BrandFix').getContent();
    var premiumShell=HtmlService.createHtmlOutputFromFile('PremiumShell').getContent();
    var clientsPro=HtmlService.createHtmlOutputFromFile('ClientsPro').getContent();
    var projectsPro=HtmlService.createHtmlOutputFromFile('ProjectsPro').getContent();
    var contentPro=HtmlService.createHtmlOutputFromFile('ContentPro').getContent();
    var paymentsBalances=HtmlService.createHtmlOutputFromFile('PaymentsBalances').getContent();
    var reports=HtmlService.createHtmlOutputFromFile('Reports').getContent();
    var activities=HtmlService.createHtmlOutputFromFile('Activities').getContent();
    var settings=HtmlService.createHtmlOutputFromFile('Settings').getContent();
    var tasks=HtmlService.createHtmlOutputFromFile('Tasks').getContent();
    var tasksNav=HtmlService.createHtmlOutputFromFile('TasksNav').getContent();
    var roleNav=HtmlService.createHtmlOutputFromFile('RoleNav').getContent();
    var html=output.getContent();
    html=html.replace('</head>',theme+premiumShell+brandFix+'</head>');
    html=html.replace('</body>',clientsPro+contentPro+paymentsBalances+reports+activities+settings+projectsPro+tasks+tasksNav+roleNav+'</body>');
    output=HtmlService.createHtmlOutput(html);
  }
  return output.setTitle(auth.authenticated?'HYPEMARK CRM':'Sign in | HYPEMARK CRM').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(fileName){return HtmlService.createHtmlOutputFromFile(fileName).getContent();}
function onOpen(){SpreadsheetApp.getUi().createMenu('🚀 HYPEMARK CRM').addItem('➕ Add Client','showAddClient').addSeparator().addItem('⚙ Initialize CRM','initializeCRM').addToUi();}
function showAddClient(){requirePermission_(AUTH_PERMISSIONS.CLIENTS);var html=HtmlService.createTemplateFromFile('AddClient').evaluate().setWidth(900).setHeight(700);SpreadsheetApp.getUi().showModalDialog(html,'Add New Client');}
function initializeCRM(){var url=initializeCRMService_();cleanupDefaultSheet_();ensureAuthSetup_();return url;}
function cleanupDefaultSheet_(){var ss=getCRMSpreadsheet_();var sheet=ss.getSheetByName('Sheet1');if(sheet&&ss.getSheets().length>1&&sheet.getLastRow()===0&&sheet.getLastColumn()===0)ss.deleteSheet(sheet);}
function createClient(data){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return createClientService_(data);} function getClients(){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return getClientsService_();} function getClient(id){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return getClientService_(id);} function updateClient(id,data){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return updateClientService_(id,data);}
function createEngagement(data){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return createEngagementService_(data);} function getEngagements(clientId){requirePermission_(AUTH_PERMISSIONS.CLIENTS);return getEngagementsService_(clientId);}
function createProject(data){requirePermission_(AUTH_PERMISSIONS.PROJECTS);return createProjectService_(data);} function getProjects(clientId){requirePermission_(AUTH_PERMISSIONS.PROJECTS);return getProjectsService_(clientId);} function getProject(id){requirePermission_(AUTH_PERMISSIONS.PROJECTS);return getProjectService_(id);}
function createDeliverable(data){requirePermission_(AUTH_PERMISSIONS.DELIVERABLES);return createDeliverableService_(data);} function getDeliverables(projectId){requirePermission_(AUTH_PERMISSIONS.DELIVERABLES);return getDeliverablesService_(projectId);} function updateDeliverableStatus(id,status){requirePermission_(AUTH_PERMISSIONS.DELIVERABLES);return updateDeliverableStatusService_(id,status);}
function createContent(data){requirePermission_(AUTH_PERMISSIONS.CONTENT);return createContentService_(data);} function getContent(deliverableId){requirePermission_(AUTH_PERMISSIONS.CONTENT);return getContentService_(deliverableId);}
function createPayment(data){requirePermission_(AUTH_PERMISSIONS.FINANCE_ALL);return createPaymentService_(data);} function getPayments(clientId,type){var ctx=requireAuth_();if(ctx.permissions.indexOf(AUTH_PERMISSIONS.FINANCE_ALL)<0&&ctx.permissions.indexOf(AUTH_PERMISSIONS.FINANCE_OWN)<0)throw new Error('You do not have permission to access payments.');if(ctx.role===AUTH_ROLES.EMPLOYEE)return getPaymentsForUser_(ctx,clientId,type);return getPaymentsService_(clientId,type);} function getPaymentOptions(){var ctx=requireAuth_();if(ctx.permissions.indexOf(AUTH_PERMISSIONS.FINANCE_ALL)<0&&ctx.permissions.indexOf(AUTH_PERMISSIONS.FINANCE_OWN)<0)throw new Error('You do not have permission to access payment options.');return getPaymentOptionsService_();} function getReceiptOptions(){requirePermission_(AUTH_PERMISSIONS.FINANCE_ALL);return getReceiptOptionsService_();}
function getClientProfitability(clientId){requirePermission_(AUTH_PERMISSIONS.PROFITABILITY);return getClientProfitabilityService_(clientId);} function getProfitabilitySummary(){requirePermission_(AUTH_PERMISSIONS.PROFITABILITY);return getProfitabilitySummaryService_();}
function getActivities(){var ctx=requireAuth_();if(ctx.permissions.indexOf(AUTH_PERMISSIONS.ACTIVITIES_ALL)<0&&ctx.permissions.indexOf(AUTH_PERMISSIONS.ACTIVITIES_OWN)<0)throw new Error('You do not have permission to access activities.');return ctx.permissions.indexOf(AUTH_PERMISSIONS.ACTIVITIES_ALL)>=0?getActivitiesService_():getActivitiesForUser_(ctx);}
function getDashboard(){var ctx=requirePermission_(AUTH_PERMISSIONS.DASHBOARD);var d=getDashboardService_();if(ctx.permissions.indexOf(AUTH_PERMISSIONS.FINANCE_ALL)<0){d.clientReceipts=0;d.employeePayments=0;d.miscExpenses=0;d.totalExpenses=0;d.netContribution=0;d.allocatedAmount=0;d.unallocatedAmount=0;d.pendingReceipts=0;d.clientProfitability=[];}if(ctx.permissions.indexOf(AUTH_PERMISSIONS.ACTIVITIES_ALL)<0)d.recentActivities=getActivitiesForUser_(ctx);return d;}
function getPaymentsForUser_(ctx,clientId,type){return getPaymentsService_(clientId,type).filter(function(p){return p.Type!=='Employee Payment'||String(p.Employee||'').trim().toLowerCase()===String(ctx.name||'').trim().toLowerCase();});}
function getActivitiesForUser_(ctx){return getActivitiesService_().filter(function(a){return String(a.User||'').trim().toLowerCase()===String(ctx.name||'').trim().toLowerCase();});}
function getTaskOptions(){requirePermission_(AUTH_PERMISSIONS.TASKS);return getTaskOptionsService_();}
function getTasks(){requirePermission_(AUTH_PERMISSIONS.TASKS);return getTasksService_();}
function createTask(data){var c=requirePermission_(AUTH_PERMISSIONS.TASKS);if(['Admin','Manager'].indexOf(c.role)<0)throw new Error('Only Admin or Manager can create tasks.');return createTaskService_(data);}
function updateTask(id,data){requirePermission_(AUTH_PERMISSIONS.TASKS);return updateTaskService_(id,data);}
function deleteTask(id){requirePermission_(AUTH_PERMISSIONS.TASKS);return deleteTaskService_(id);}
function getTaskUpdates(id){requirePermission_(AUTH_PERMISSIONS.TASKS);return getTaskUpdatesService_(id);}
function getTaskDashboardStats(){requirePermission_(AUTH_PERMISSIONS.TASKS);return getTaskDashboardStatsService_();}
function approveTask(id,approved,notes){requirePermission_(AUTH_PERMISSIONS.TASKS);return approveTaskService_(id,approved,notes);}
