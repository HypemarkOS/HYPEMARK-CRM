var DELIVERABLE_HEADERS = ['Deliverable ID','Project ID','Type','Quantity','Status','Priority','Assigned To','Start Date','Due Date','Notes'];

function ensureDeliverableSchema_() {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES);
  if (!sheet) return;
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = false;
  DELIVERABLE_HEADERS.forEach(function(header) {
    if (current.indexOf(header) < 0) { current.push(header); changed = true; }
  });
  if (changed) sheet.getRange(1, 1, 1, current.length).setValues([current]);
}

function createDeliverableService_(data) {
  data = data || {};
  ensureDeliverableSchema_();
  if (!data.projectId) throw new Error('Project is required.');
  if (!asText_(data.type)) throw new Error('Deliverable type is required.');
  if (!getProjectService_(data.projectId)) throw new Error('Project not found.');
  var item = {
    'Deliverable ID': generateId_(CONFIG.ID_PREFIX.DELIVERABLE),
    'Project ID': data.projectId,
    'Type': asText_(data.type),
    'Quantity': Math.max(1, Number(data.quantity) || 1),
    'Status': asText_(data.status) || CONFIG.DEFAULTS.DELIVERABLE_STATUS,
    'Priority': asText_(data.priority) || 'Medium',
    'Assigned To': asText_(data.assignedTo),
    'Start Date': data.startDate ? new Date(data.startDate) : '',
    'Due Date': data.dueDate ? new Date(data.dueDate) : '',
    'Notes': asText_(data.notes)
  };
  appendRow_(APP.SHEETS.DELIVERABLES, DELIVERABLE_HEADERS, item);
  recordActivity_('Deliverable', item['Deliverable ID'], 'Created deliverable');
  return serializeDeliverable_(item);
}

function getDeliverablesService_(projectId) {
  ensureDeliverableSchema_();
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues(), headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(value) { return value !== ''; }) && (!projectId || String(row[headers.indexOf('Project ID')]) === String(projectId));
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return serializeDeliverable_(addDeliverableTaskSummary_(object));
  });
}

function getDeliverableService_(deliverableId) {
  var items = getDeliverablesService_();
  return items.find(function(item) { return String(item['Deliverable ID']) === String(deliverableId); }) || null;
}

function getDeliverableTaskSummaryService_(deliverableId) {
  var item = getDeliverableService_(deliverableId);
  if (!item) throw new Error('Deliverable not found.');
  return {
    'Deliverable ID': item['Deliverable ID'],
    'Task Count': Number(item['Task Count'] || 0),
    'Completed Tasks': Number(item['Completed Tasks'] || 0),
    'In Progress Tasks': Number(item['In Progress Tasks'] || 0),
    'Pending Review Tasks': Number(item['Pending Review Tasks'] || 0),
    'On Hold Tasks': Number(item['On Hold Tasks'] || 0),
    'To Do Tasks': Number(item['To Do Tasks'] || 0),
    'Overdue Tasks': Number(item['Overdue Tasks'] || 0),
    'Task Progress': Number(item['Task Progress'] || 0)
  };
}

function addDeliverableTaskSummary_(item) {
  var result = item || {};
  var summary = {total:0,completed:0,inProgress:0,pendingReview:0,onHold:0,toDo:0,overdue:0};
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.TASKS);
  if (sheet && sheet.getLastRow() >= 2) {
    var rows = sheet.getDataRange().getValues(), headers = rows.shift();
    var didx = headers.indexOf('Deliverable ID'), sidx = headers.indexOf('Status'), dueidx = headers.indexOf('Due Date');
    if (didx >= 0) rows.forEach(function(row) {
      if (String(row[didx] == null ? '' : row[didx]).trim() !== String(result['Deliverable ID'] == null ? '' : result['Deliverable ID']).trim()) return;
      summary.total++;
      var status = String(sidx >= 0 ? row[sidx] : '');
      if (status === 'Completed') summary.completed++;
      else if (status === 'In Progress') summary.inProgress++;
      else if (status === 'Pending Review') summary.pendingReview++;
      else if (status === 'On Hold') summary.onHold++;
      else if (status === 'To Do') summary.toDo++;
      if (dueidx >= 0 && status !== 'Completed' && row[dueidx]) {
        var due = row[dueidx] instanceof Date ? new Date(row[dueidx]) : new Date(row[dueidx]);
        if (!isNaN(due.getTime())) { due.setHours(0,0,0,0); var today = new Date(); today.setHours(0,0,0,0); if (due < today) summary.overdue++; }
      }
    });
  }
  result['Task Count'] = summary.total;
  result['Completed Tasks'] = summary.completed;
  result['In Progress Tasks'] = summary.inProgress;
  result['Pending Review Tasks'] = summary.pendingReview;
  result['On Hold Tasks'] = summary.onHold;
  result['To Do Tasks'] = summary.toDo;
  result['Overdue Tasks'] = summary.overdue;
  result['Task Progress'] = summary.total ? Math.round((summary.completed / summary.total) * 100) : 0;
  return result;
}

function serializeDeliverable_(item) {
  var copy = {};
  Object.keys(item || {}).forEach(function(key) {
    var value = item[key];
    copy[key] = value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd-MMM-yyyy') : value;
  });
  return copy;
}

function updateDeliverableStatusService_(id, status) {
  var s = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES);
  if (!s) throw new Error('Deliverable not found.');
  var v = s.getDataRange().getValues(), h = v[0], ii = h.indexOf('Deliverable ID'), si = h.indexOf('Status');
  for (var i = 1; i < v.length; i++) if (String(v[i][ii]) === String(id)) {
    s.getRange(i + 1, si + 1).setValue(asText_(status));
    recordActivity_('Deliverable', id, 'Updated deliverable status');
    return true;
  }
  throw new Error('Deliverable not found.');
}

function getDeliverableAssigneesService_() {
  var users = [];
  if (typeof getUsersService_ === 'function') users = getUsersService_() || [];
  else if (typeof getUsers === 'function') users = getUsers() || [];
  return users.filter(function(u) {
    var status = String(u.Status || u.status || 'Active').toLowerCase();
    return status === 'active' || status === '';
  }).map(function(u) {
    return {
      id: u['User ID'] || u.userId || u.ID || u.Email || u.email || '',
      name: u.Name || u.name || u.Email || u.email || '',
      role: u.Role || u.role || '',
      jobFunction: u['Job Function'] || u.jobFunction || ''
    };
  }).filter(function(u) { return u.id && u.name; });
}
