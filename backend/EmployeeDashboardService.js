function getEmployeeDashboardService_(ctx){
  var name=String(ctx&&ctx.name||'').trim().toLowerCase();
  var tasks=getTasksService_().filter(function(t){return String(t['Assigned To']||'').trim().toLowerCase()===name;});
  var deliverables=getDeliverablesService_().filter(function(d){return String(d['Assigned To']||'').trim().toLowerCase()===name;});
  var content=[];try{content=getContentService_();}catch(e){content=[];}
  content=content.filter(function(x){return String(x['Assigned To']||'').trim().toLowerCase()===name;});
  var deadlines=[];
  tasks.forEach(function(t){if(t['Due Date']&&t.Status!=='Completed')deadlines.push({type:'Task',title:t['Task Name'],due:t['Due Date'],status:t.Status});});
  deliverables.forEach(function(d){if(d['Due Date']&&d.Status!=='Completed')deadlines.push({type:'Deliverable',title:d.Type||d['Deliverable ID'],due:d['Due Date'],status:d.Status});});
  deadlines.sort(function(a,b){return new Date(a.due)-new Date(b.due);});
  return {mode:'employee',name:ctx.name,jobFunction:ctx.jobFunction||'',tasks:tasks.slice(0,12),deliverables:deliverables.slice(0,12),content:content.slice(0,12),upcomingDeadlines:deadlines.slice(0,10),pendingReviews:tasks.filter(function(t){return t.Status==='Pending Review';}).slice(0,10),completedWork:tasks.filter(function(t){return t.Status==='Completed';}).slice(-10).reverse()};
}
