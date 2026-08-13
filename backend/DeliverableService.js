var DELIVERABLE_HEADERS = ['Deliverable ID','Project ID','Type','Quantity','Status','Priority','Due Date'];

function createDeliverableService_(data) {
  data = data || {};
  if (!data.projectId) throw new Error('Project is required.');
  if (!asText_(data.type)) throw new Error('Deliverable type is required.');
  if (!getProjectService_(data.projectId)) throw new Error('Project not found.');
  var item = {
    'Deliverable ID': generateId_(CONFIG.ID_PREFIX.DELIVERABLE),
    'Project ID': data.projectId,
    'Type': asText_(data.type),
    'Quantity': Number(data.quantity) || 1,
    'Status': asText_(data.status) || CONFIG.DEFAULTS.DELIVERABLE_STATUS,
    'Priority': asText_(data.priority) || 'Medium',
    'Due Date': data.dueDate ? new Date(data.dueDate) : ''
  };
  appendRow_(APP.SHEETS.DELIVERABLES, DELIVERABLE_HEADERS, item);
  recordActivity_('Deliverable', item['Deliverable ID'], 'Created deliverable');
  return serializeDeliverable_(item);
}

function getDeliverablesService_(projectId) {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues(), headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(value) { return value !== ''; }) && (!projectId || String(row[headers.indexOf('Project ID')]) === String(projectId));
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return serializeDeliverable_(object);
  });
}

function getDeliverableService_(deliverableId) {
  var items = getDeliverablesService_();
  return items.find(function(item) { return String(item['Deliverable ID']) === String(deliverableId); }) || null;
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
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][ii]) === String(id)) {
      s.getRange(i + 1, si + 1).setValue(asText_(status));
      recordActivity_('Deliverable', id, 'Updated deliverable status');
      return true;
    }
  }
  throw new Error('Deliverable not found.');
}
