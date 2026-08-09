function getDashboardService_(){
  var clients=getClientsService_(),projects=getProjectsService_(),payments=getPaymentsService_(),activities=getActivitiesService_();
  var received=0,employee=0,misc=0,pending=0;
  payments.forEach(function(p){
    var amount=Number(p['Amount']||0),type=p['Type'];
    if(type==='Client Receipt'){received+=amount;if(String(p['Status']).toLowerCase()!=='received'&&String(p['Status']).toLowerCase()!=='completed')pending+=amount;}
    if(type==='Employee Payment')employee+=amount;
    if(type==='Miscellaneous Expense')misc+=amount;
  });
  return {totalClients:clients.length,activeClients:clients.filter(function(c){return c['Status']==='Active';}).length,totalProjects:projects.length,clientReceipts:received,employeePayments:employee,miscExpenses:misc,totalExpenses:employee+misc,pendingPayments:pending,clientProfitability:getProfitabilitySummaryService_(),recentActivities:activities.slice(-10).reverse()};
}

function getActivitiesService_(){var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.ACTIVITIES);if(!s||s.getLastRow()<2)return [];var v=s.getDataRange().getValues(),h=v.shift();return v.filter(function(r){return r.some(function(x){return x!=='';});}).map(function(r){var o={};h.forEach(function(k,i){o[k]=r[i] instanceof Date?Utilities.formatDate(r[i],CONFIG.TIMEZONE,'dd-MMM-yyyy HH:mm'):r[i];});return o;});}
