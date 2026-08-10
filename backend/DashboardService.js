function getDashboardService_() {
  var ss = getCRMSpreadsheet_();
  var clients = getClientsService_();
  var projects = getProjectsService_();
  var activities = getActivitiesService_();
  var receipts = readDashboardSheet_(ss, APP.SHEETS.CLIENT_RECEIPTS);
  var employee = readDashboardSheet_(ss, APP.SHEETS.EMPLOYEE_PAYMENTS);
  var expenses = readDashboardSheet_(ss, APP.SHEETS.BUSINESS_EXPENSES);

  var received = sumField_(receipts, 'Amount');
  var employeePaid = sumField_(employee, 'Amount');
  var businessExpenses = sumField_(expenses, 'Amount');
  var allocated = sumField_(receipts, 'Allocated Amount');
  var unallocated = Math.max(0, received - allocated);

  var activeProjects = projects.filter(function(p){ return String(p.Status || '') === 'Active'; }).length;
  var overdue = countOverdue_(projects);
  var pendingReceipts = receipts.filter(function(r){
    var status = String(r.Status || '').toLowerCase();
    return status !== 'received' && status !== 'completed';
  }).length;

  var profitability = getProfitabilitySummaryService_().map(function(x){
    var p = x.profitability || {};
    return {
      clientId: x.clientId,
      clientName: x.clientName,
      businessName: x.businessName,
      revenue: Number(p.receipts || 0),
      costs: Number(p.totalCosts || 0),
      profit: Number(p.profit || 0)
    };
  }).sort(function(a,b){ return b.profit - a.profit; });

  return {
    totalClients: clients.length,
    activeClients: clients.filter(function(c){return c.Status === 'Active';}).length,
    totalProjects: projects.length,
    activeProjects: activeProjects,
    clientReceipts: received,
    employeePayments: employeePaid,
    miscExpenses: businessExpenses,
    totalExpenses: employeePaid + businessExpenses,
    netContribution: received - employeePaid - businessExpenses,
    allocatedAmount: allocated,
    unallocatedAmount: unallocated,
    pendingReceipts: pendingReceipts,
    overdueProjects: overdue,
    clientProfitability: profitability.slice(0, 10),
    recentActivities: activities.slice(-8).reverse()
  };
}

function readDashboardSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values.filter(function(row){ return row.some(function(v){ return v !== ''; }); }).map(function(row){
    var item = {};
    headers.forEach(function(h,i){ item[h] = row[i]; });
    return item;
  });
}

function sumField_(rows, field) {
  return rows.reduce(function(total, row){ return total + Number(row[field] || 0); }, 0);
}

function countOverdue_(rows) {
  var today = new Date();
  today.setHours(0,0,0,0);
  return rows.filter(function(row){
    if (!row['Due Date']) return false;
    var status = String(row.Status || '').toLowerCase();
    if (status === 'completed' || status === 'closed' || status === 'cancelled') return false;
    var date = row['Due Date'] instanceof Date ? new Date(row['Due Date']) : new Date(row['Due Date']);
    if (isNaN(date.getTime())) return false;
    date.setHours(0,0,0,0);
    return date < today;
  }).length;
}
