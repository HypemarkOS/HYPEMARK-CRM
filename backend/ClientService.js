function createClient(data) {
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
  appendRow_(APP.SHEETS.CLIENTS, ['Client ID','Client Name','Business Name','Contact Person','Mobile','Email','Address','GST','Status','Created On'], client);
  logInfo_('Client created', { id: client['Client ID'] });
  return client;
}
function getClients() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP.SHEETS.CLIENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var headers = values.shift();
  return values.map(function(row){ var o={}; headers.forEach(function(h,i){ o[h]=row[i]; }); return o; });
}
