var ENGAGEMENT_HEADERS = ['Engagement ID','Client ID','Service','Start Date','End Date','Status','Priority'];

function createEngagementService_(data) {
  data = data || {};
  if (!data.clientId) throw new Error('Client is required.');
  if (!asText_(data.service)) throw new Error('Service is required.');
  var client = getClientService_(data.clientId);
  if (!client) throw new Error('Client not found.');
  var engagement = {
    'Engagement ID': generateId_(CONFIG.ID_PREFIX.ENGAGEMENT),
    'Client ID': data.clientId,
    'Service': asText_(data.service),
    'Start Date': data.startDate ? new Date(data.startDate) : '',
    'End Date': data.endDate ? new Date(data.endDate) : '',
    'Status': asText_(data.status) || CONFIG.DEFAULTS.CLIENT_STATUS,
    'Priority': asText_(data.priority) || 'Medium'
  };
  appendRow_(APP.SHEETS.ENGAGEMENTS, ENGAGEMENT_HEADERS, engagement);
  recordActivity_('Engagement', engagement['Engagement ID'], 'Created engagement');
  return serializeEngagement_(engagement);
}

function getEngagementsService_(clientId) {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.ENGAGEMENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values.filter(function(row) {
    if (!row.some(function(value) { return value !== ''; })) return false;
    return !clientId || String(row[headers.indexOf('Client ID')]) === String(clientId);
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return serializeEngagement_(object);
  });
}

function serializeEngagement_(item) {
  var copy = {};
  Object.keys(item || {}).forEach(function(key) {
    var value = item[key];
    copy[key] = value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd-MMM-yyyy') : value;
  });
  return copy;
}
