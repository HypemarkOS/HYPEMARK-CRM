var PROJECT_HEADERS = ['Project ID','Client ID','Engagement ID','Project Name','Status','Priority','Project Value','Start Date','Due Date','Owner / Assignee','Notes','Created On'];

function ensureProjectSheetHeaders_() {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.PROJECTS);
  if (!sheet) throw new Error('Projects sheet not found. Run CRM initialization first.');

  var lastColumn = sheet.getLastColumn();
  var headers = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) { return String(v || '').trim(); }) : [];

  if (!headers.some(function(h) { return h !== ''; })) {
    sheet.getRange(1, 1, 1, PROJECT_HEADERS.length).setValues([PROJECT_HEADERS]);
    return PROJECT_HEADERS.slice();
  }

  PROJECT_HEADERS.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      var nextColumn = headers.length + 1;
      sheet.getRange(1, nextColumn).setValue(header);
      headers.push(header);
    }
  });

  return headers;
}

function createProjectService_(data) {
  data = data || {};
  var clientId = asText_(data.clientId || data.ClientID || data['Client ID']);
  var projectName = asText_(data.projectName || data.Name || data['Project Name']);
  var engagementId = asText_(data.engagementId || data.EngagementID || data['Engagement ID']);
  var status = asText_(data.status || data.Status);
  var priority = asText_(data.priority || data.Priority);
  var projectValue = data.projectValue !== undefined ? data.projectValue : (data.Value !== undefined ? data.Value : data['Project Value']);
  var startDate = asText_(data.startDate || data.StartDate || data['Start Date']);
  var dueDate = asText_(data.dueDate || data.Deadline || data['Due Date']);
  var owner = asText_(data.owner || data.Owner || data['Owner / Assignee']);
  var notes = asText_(data.notes || data.Notes);

  if (!clientId) throw new Error('Client is required.');
  if (!projectName) throw new Error('Project Name is required.');
  var client = getClientService_(clientId);
  if (!client) throw new Error('Client not found.');

  if (engagementId) {
    var engagements = getEngagementsService_(clientId);
    var engagement = engagements.find(function(item) { return item['Engagement ID'] === engagementId; });
    if (!engagement) throw new Error('Engagement does not belong to this client.');
  }

  var value = projectValue === '' || projectValue === null || projectValue === undefined ? 0 : Number(String(projectValue).replace(/,/g, ''));
  if (!isFinite(value) || value < 0) throw new Error('Project Value must be a valid non-negative number.');

  if (owner) {
    var ownerUser = getActiveProjectOwnerById_(owner);
    if (!ownerUser) {
      var legacyOwner = getActiveProjectOwnerByName_(owner);
      if (legacyOwner) owner = legacyOwner.userId;
      else throw new Error('Selected Project Owner is no longer an active CRM user. Please choose another person.');
    }
  }

  var project = {
    'Project ID': generateId_(CONFIG.ID_PREFIX.PROJECT),
    'Client ID': clientId,
    'Engagement ID': engagementId,
    'Project Name': projectName,
    'Status': status || CONFIG.DEFAULTS.PROJECT_STATUS,
    'Priority': priority || 'Medium',
    'Project Value': value,
    'Start Date': startDate,
    'Due Date': dueDate,
    'Owner / Assignee': owner,
    'Notes': notes,
    'Created On': now_()
  };

  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.PROJECTS);
  var headers = ensureProjectSheetHeaders_();
  var row = headers.map(function(header) { return project[header] !== undefined ? project[header] : ''; });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([row]);
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

function getProjectOwners() {
  requirePermission_(AUTH_PERMISSIONS.PROJECTS);
  return getActiveProjectOwnersService_();
}

function getActiveProjectOwnersService_() {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  var idIdx = headers.indexOf('User ID');
  var nameIdx = headers.indexOf('Name');
  var roleIdx = headers.indexOf('Role');
  var jobIdx = headers.indexOf('Job Function');
  var emailIdx = headers.indexOf('Email');
  var statusIdx = headers.indexOf('Status');
  return values.filter(function(row) {
    return String(row[statusIdx] || 'Active') === 'Active' && String(row[idIdx] || '').trim() !== '';
  }).map(function(row) {
    return {userId:String(row[idIdx] || '').trim(),name:String(row[nameIdx] || '').trim(),role:String(row[roleIdx] || '').trim(),jobFunction:jobIdx>=0?String(row[jobIdx] || '').trim():'',email:emailIdx>=0?String(row[emailIdx] || '').trim():''};
  });
}

function getActiveProjectOwnerById_(userId) {
  userId = String(userId || '').trim();
  if (!userId) return null;
  return getActiveProjectOwnersService_().find(function(user) { return user.userId === userId; }) || null;
}

function getActiveProjectOwnerByName_(name) {
  name = String(name || '').trim().toLowerCase();
  if (!name) return null;
  return getActiveProjectOwnersService_().find(function(user) { return user.name.toLowerCase() === name; }) || null;
}

function serializeProject_(item) {
  var copy = {};
  Object.keys(item || {}).forEach(function(key) {
    var value = item[key];
    copy[key] = value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm') : value;
  });
  var ownerId = String(copy['Owner / Assignee'] || '').trim();
  if (ownerId) {
    var ownerUser = getActiveProjectOwnerById_(ownerId);
    if (ownerUser) {
      copy['Owner User ID'] = ownerUser.userId;
      copy['Owner / Assignee'] = ownerUser.name;
      copy.Owner = ownerUser.name;
      copy.OwnerUserID = ownerUser.userId;
    }
  }
  return copy;
}
