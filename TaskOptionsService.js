function getTaskOptionsFast(){
  var c=requirePermission_(AUTH_PERMISSIONS.TASKS);
  var clients=getClientsService_();
  var projects=getProjectsService_();
  var deliverables=getDeliverablesService_();
  var assignees=[];
  var users=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  if(users&&users.getLastRow()>1){
    var v=users.getDataRange().getValues();
    v.shift();
    v.forEach(function(r){
      var name=asText_(r[1]),role=asText_(r[2]),email=asText_(r[3]),status=asText_(r[4]);
      if(name&&status==='Active'&&(c.role!=='Employee'||name.toLowerCase()===c.name.toLowerCase())) assignees.push({name:name,role:role,email:email});
    });
  }
  return {clients:clients,projects:projects,deliverables:deliverables,assignees:assignees,currentUser:{name:c.name,role:c.role},statuses:['To Do','In Progress','Pending Review','Completed','On Hold'],priorities:['Low','Medium','High','Critical']};
}
