var DELIVERABLE_HEADERS = ['Deliverable ID','Project ID','Type','Quantity','Status'];

function createDeliverableService_(data) {
  if (!data || !data.projectId) throw new Error('Project is required.');
  if (!asText_(data.type)) throw new Error('Deliverable type is required.');
  if (!getProjectService_(data.projectId)) throw new Error('Project not found.');
  var item = {
    'Deliverable ID': generateId_(CONFIG.ID_PREFIX.DELIVERABLE),
    'Project ID': data.projectId,
    'Type': asText_(data.type),
    'Quantity': Number(data.quantity) || 1,
    'Status': asText_(data.status) || CONFIG.DEFAULTS.DELIVERABLE_STATUS
  };
  appendRow_(APP.SHEETS.DELIVERABLES, DELIVERABLE_HEADERS, item);
  recordActivity_('Deliverable', item['Deliverable ID'], 'Created deliverable');
  return item;
}
function getDeliverablesService_(projectId) {
  var sheet=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES); if(!sheet||sheet.getLastRow()<2)return [];
  var v=sheet.getDataRange().getValues(), h=v.shift();
  return v.filter(function(r){return r.some(function(x){return x!=='';})&&(!projectId||String(r[h.indexOf('Project ID')])===String(projectId));}).map(function(r){var o={};h.forEach(function(k,i){o[k]=r[i]});return o;});
}
function updateDeliverableStatusService_(id,status){var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.DELIVERABLES);if(!s)throw new Error('Deliverable not found.');var v=s.getDataRange().getValues(),h=v[0],ii=h.indexOf('Deliverable ID'),si=h.indexOf('Status');for(var i=1;i<v.length;i++){if(String(v[i][ii])===String(id)){s.getRange(i+1,si+1).setValue(asText_(status));recordActivity_('Deliverable',id,'Updated deliverable status');return true;}}throw new Error('Deliverable not found.');}
