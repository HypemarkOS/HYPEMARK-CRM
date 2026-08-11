var CONTENT_HEADERS = ['Content ID','Deliverable ID','Title','Status','Assigned To','Publish Date'];

function contentNormalizeHeader_(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function contentFindColumn_(headers, name) {
  var wanted = contentNormalizeHeader_(name);
  for (var i = 0; i < headers.length; i++) {
    if (contentNormalizeHeader_(headers[i]) === wanted) return i;
  }
  return -1;
}

function contentSerializeValue_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (value == null) return '';
  return value;
}

function createContentService_(data) {
  if (!data || !data.deliverableId) throw new Error('Deliverable is required.');
  if (!asText_(data.title)) throw new Error('Content title is required.');

  var item = {
    'Content ID': generateId_(CONFIG.ID_PREFIX.CONTENT),
    'Deliverable ID': asText_(data.deliverableId),
    'Title': asText_(data.title),
    'Status': asText_(data.status) || CONFIG.DEFAULTS.DELIVERABLE_STATUS,
    'Assigned To': asText_(data.assignedTo),
    'Publish Date': data.publishDate ? new Date(data.publishDate) : ''
  };

  appendRow_(APP.SHEETS.CONTENT_BANK, CONTENT_HEADERS, item);
  recordActivity_('Content', item['Content ID'], 'Created content');

  // Never return a native Date to google.script.run.
  return {
    'Content ID': item['Content ID'],
    'Deliverable ID': item['Deliverable ID'],
    'Title': item['Title'],
    'Status': item['Status'],
    'Assigned To': item['Assigned To'],
    'Publish Date': contentSerializeValue_(item['Publish Date'])
  };
}

function getContentService_(deliverableId) {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CONTENT_BANK);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  var headers = values.shift().map(function (h) { return String(h == null ? '' : h).trim(); });
  var idCol = contentFindColumn_(headers, 'Content ID');
  var deliverableCol = contentFindColumn_(headers, 'Deliverable ID');
  var titleCol = contentFindColumn_(headers, 'Title');
  var statusCol = contentFindColumn_(headers, 'Status');
  var assignedCol = contentFindColumn_(headers, 'Assigned To');
  var publishCol = contentFindColumn_(headers, 'Publish Date');

  return values.filter(function (row) {
    var hasData = row.some(function (cell) { return cell !== '' && cell != null; });
    if (!hasData) return false;
    if (!deliverableId || deliverableCol < 0) return true;
    return String(row[deliverableCol] == null ? '' : row[deliverableCol]).trim() === String(deliverableId).trim();
  }).map(function (row) {
    return {
      'Content ID': idCol >= 0 ? contentSerializeValue_(row[idCol]) : '',
      'Deliverable ID': deliverableCol >= 0 ? contentSerializeValue_(row[deliverableCol]) : '',
      'Title': titleCol >= 0 ? contentSerializeValue_(row[titleCol]) : '',
      'Status': statusCol >= 0 ? contentSerializeValue_(row[statusCol]) : '',
      'Assigned To': assignedCol >= 0 ? contentSerializeValue_(row[assignedCol]) : '',
      'Publish Date': publishCol >= 0 ? contentSerializeValue_(row[publishCol]) : ''
    };
  });
}
