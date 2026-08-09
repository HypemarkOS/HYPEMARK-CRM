var PROJECT_HEADERS = ['Project ID','Client ID','Engagement ID','Project Name','Status','Created On'];

function createProjectService_(data) {
  if (!data || !data.clientId) throw new Error('Client is required.');
  if (!asText_(data.projectName)) throw new Error('Project Name is required.');
  var client = getClientService_(data.clientId);
  if (!client) throw new Error('Client not found.');
  if (data.engagementId) {
    var engagements = getEngagementsService_(data.clientId);
    var engagement = engagements.find(function(item) { return item['Engagement ID'] === data.engagementId; });
    if (!engagement) throw new Error('Engagement does not belong to this client.');
  }
  var project = {
    'Project ID': generateId_(CONFIG.ID_PREFIX.PROJECT),
    'Client ID': data.clientId,
    'Engagement ID': asText_(data.engagementId),
    'Project Name': asText_(data.projectName),
    'Status': asText_(data.status) || CONFIG.DEFAULTS.PROJECT_STATUS,
    'Created On': now_()
  };
  appendRow_(APP.SHEETS.PROJECTS, PROJECT_HEADERS, project);
  recordActivity_('Project', project['Project ID'], 'Created project');
  return serializeProject_(project);
}

function getProjectsService_(clientId) {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.PROJECTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values.filter(function(row) {
    if (!row.some(function(value) { return value !== ''; })) return false;
    return !clientId || String(row[headers.indexOf('Client ID')]) === String(clientId);
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return serializeProject_(object);
  });
}

function getProjectService_(projectId) {
  var projects = getProjectsService_();
  return projects.find(function(project) { return project['Project ID'] === projectId; }) || null;
}

function serializeProject_(item) {
  var copy = {};
  Object.keys(item || {}).forEach(function(key) {
    var value = item[key];
    copy[key] = value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm') : value;
  });
  return copy;
}
