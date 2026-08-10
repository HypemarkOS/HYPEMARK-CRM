function getDashboardService_() {
  var clients = getClientsService_(), projects = getProjectsService_(), activities = getActivitiesService_();
  var payments = getPaymentsService_();
  var received = 0, employee = 0, misc = 0, pending = 0;
  payments.forEach(function(p) {
    var amount = Number(p.Amount || 0), type = p.Type;
    if (type === 'Client Receipt') {
      received += amount;
      if (String(p.Status).toLowerCase() !== 'received' && String(p.Status).toLowerCase() !== 'completed') pending += amount;
    }
    if (type === 'Employee Payment') employee += amount;
    if (type === 'Business Expense') misc += amount;
  });
  return {
    totalClients: clients.length,
    activeClients: clients.filter(function(c){return c.Status === 'Active';}).length,
    totalProjects: projects.length,
    clientReceipts: received,
    employeePayments: employee,
    miscExpenses: misc,
    totalExpenses: employee + misc,
    pendingPayments: pending,
    clientProfitability: getProfitabilitySummaryService_(),
    recentActivities: activities.slice(-10).reverse()
  };
}
