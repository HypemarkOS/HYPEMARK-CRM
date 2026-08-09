var CLIENT_HEADERS = ['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On'];

function createClientService_(data) {
  var validation = validateClient_(data);
  if (!validation.valid) throw new Error(validation.errors.join(' '));

  var client = {
    'Client ID': generateId_(CONFIG.ID_PREFIX.CLIENT),
    'Client Name': asText_(data.clientName),
    'Business Name': asText_(data.businessName),
    'Contact Person': asText_(data.contactPerson),
    'Mobile': normalizeMobile_(data.mobile),
    'Email': normalizeEmail_(data.email),
    'Address': asText_(data.address),
    'GST': asText_(data.gst),
    'Status': CONFIG.DEFAULTS.CLIENT_STATUS,
    'Created On': now_()
  };

  appendRow_(APP.SHEETS.CLIENTS, CLIENT_HEADERS, client);
  recordActivity_('Client', client['Client ID'], 'Created client');
  logInfo_('Client created', { id: client['Client ID'] });
  return serializeClient_(client);
}

function getClientsService_() {
  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CLIENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(value) { return value !== ''; });
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return serializeClient_(object);
  });
}

function getClientService_(clientId) {
  var clients = getClientsService_();
  return clients.find(function(client) { return client['Client ID'] === clientId; }) || null;
}

function updateClientService_(clientId, data) {
  var validation = validateClient_(data);
  if (!validation.valid) throw new Error(validation.errors.join(' '));

  var sheet = getCRMSpreadsheet_().getSheetByName(APP.SHEETS.CLIENTS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Client not found.');

  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idIndex = headers.indexOf('Client ID');
  var rowIndex = -1;

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(clientId)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Client not found.');

  var existing = getClientService_(clientId);
  var updated = {
    'Client ID': clientId,
    'Client Name': asText_(data.clientName),
    'Business Name': asText_(data.businessName),
    'Contact Person': asText_(data.contactPerson),
    'Mobile': normalizeMobile_(data.mobile),
    'Email': normalizeEmail_(data.email),
    'Address': asText_(data.address),
    'GST': asText_(data.gst),
    'Status': existing['Status'] || CONFIG.DEFAULTS.CLIENT_STATUS,
    'Created On': existing['Created On'] || now_()
  };

  sheet.getRange(rowIndex, 1, 1, CLIENT_HEADERS.length)
    .setValues([CLIENT_HEADERS.map(function(header) { return updated[header]; })]);

  recordActivity_('Client', clientId, 'Updated client');
  return serializeClient_(updated);
}

function serializeClient_(client) {
  if (!client) return null;
  var copy = {};
  Object.keys(client).forEach(function(key) {
    var value = client[key];
    copy[key] = value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm') : value;
  });
  return copy;
}
