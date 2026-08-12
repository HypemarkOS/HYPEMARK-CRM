/**
 * HYPEMARK CRM authentication and role-based access control.
 * Authentication uses the Google account identity available to Apps Script.
 * No passwords are stored in the CRM.
 */
var PRIMARY_ADMIN_EMAIL = 'namani.ajay@gmail.com';
var AUTH_ROLES = { ADMIN:'Admin', PARTNER:'Partner', MANAGER:'Manager', EMPLOYEE:'Employee' };
var AUTH_PERMISSIONS = { DASHBOARD:'dashboard', CLIENTS:'clients', PROJECTS:'projects', TASKS:'tasks', DELIVERABLES:'deliverables', CONTENT:'content', FINANCE_ALL:'finance_all', FINANCE_OWN:'finance_own', PROFITABILITY:'profitability', ACTIVITIES_ALL:'activities_all', ACTIVITIES_OWN:'activities_own', USERS:'users', SETTINGS:'settings' };
var ROLE_PERMISSIONS = {};
ROLE_PERMISSIONS[AUTH_ROLES.ADMIN] = Object.keys(AUTH_PERMISSIONS).map(function(k){return AUTH_PERMISSIONS[k];});
ROLE_PERMISSIONS[AUTH_ROLES.PARTNER] = [AUTH_PERMISSIONS.DASHBOARD,AUTH_PERMISSIONS.CLIENTS,AUTH_PERMISSIONS.PROJECTS,AUTH_PERMISSIONS.TASKS,AUTH_PERMISSIONS.DELIVERABLES,AUTH_PERMISSIONS.CONTENT,AUTH_PERMISSIONS.FINANCE_ALL,AUTH_PERMISSIONS.PROFITABILITY,AUTH_PERMISSIONS.ACTIVITIES_ALL];
ROLE_PERMISSIONS[AUTH_ROLES.MANAGER] = [AUTH_PERMISSIONS.DASHBOARD,AUTH_PERMISSIONS.CLIENTS,AUTH_PERMISSIONS.PROJECTS,AUTH_PERMISSIONS.TASKS,AUTH_PERMISSIONS.DELIVERABLES,AUTH_PERMISSIONS.CONTENT,AUTH_PERMISSIONS.ACTIVITIES_ALL];
ROLE_PERMISSIONS[AUTH_ROLES.EMPLOYEE] = [AUTH_PERMISSIONS.DASHBOARD,AUTH_PERMISSIONS.PROJECTS,AUTH_PERMISSIONS.TASKS,AUTH_PERMISSIONS.DELIVERABLES,AUTH_PERMISSIONS.CONTENT,AUTH_PERMISSIONS.FINANCE_OWN,AUTH_PERMISSIONS.ACTIVITIES_OWN];

function getCurrentUserEmail_(){
  var email='';
  try{ email=String(Session.getActiveUser().getEmail()||'').trim().toLowerCase(); }catch(e){}
  if(!email){
    try{ email=String(Session.getEffectiveUser().getEmail()||'').trim().toLowerCase(); }catch(e){}
  }
  return email;
}

function getAuthContextService_(){
  var email=getCurrentUserEmail_();
  var users=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);
  var rows=users&&users.getLastRow()>=2?users.getDataRange().getValues():[];
  if(rows.length)rows.shift();
  if(email===PRIMARY_ADMIN_EMAIL&&rows.length===0){
    var id=generateId_('USR');
    users.appendRow([id,'Ajay',AUTH_ROLES.ADMIN,email,'Active']);
    rows=[[id,'Ajay',AUTH_ROLES.ADMIN,email,'Active']];
    ensureAuthSetup_();
  }
  var found=null;
  rows.forEach(function(r){
    if(String(r[3]||'').trim().toLowerCase()===email){
      found={userId:String(r[0]||''),name:String(r[1]||''),role:String(r[2]||''),email:email,status:String(r[4]||'')};
    }
  });
  if(!email)return{authenticated:false,reason:'identity_unavailable',email:'',name:'',role:'',permissions:[]};
  if(!found)return{authenticated:false,reason:'not_registered',email:email,name:'',role:'',permissions:[]};
  if(found.role===AUTH_ROLES.ADMIN&&email!==PRIMARY_ADMIN_EMAIL)return{authenticated:false,reason:'admin_restricted',email:email,name:found.name,role:found.role,permissions:[]};
  if(found.status!=='Active')return{authenticated:false,reason:'inactive',email:email,name:found.name,role:found.role,permissions:[]};
  return{authenticated:true,reason:'ok',userId:found.userId,name:found.name,role:found.role,email:found.email,status:found.status,permissions:ROLE_PERMISSIONS[found.role]||[]};
}

function requireAuth_(){var ctx=getAuthContextService_();if(!ctx.authenticated)throw new Error('CRM access denied. Please sign in with an authorised Google account.');return ctx;}
function requirePermission_(permission){var ctx=requireAuth_();if(ctx.permissions.indexOf(permission)===-1)throw new Error('You do not have permission to access this section.');return ctx;}
function getAuthContext(){return getAuthContextService_();}
function getUsers(){requirePermission_(AUTH_PERMISSIONS.USERS);var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);if(!s||s.getLastRow()<2)return[];var v=s.getDataRange().getValues(),h=v.shift();return v.filter(function(r){return r.some(function(x){return x!=='';});}).map(function(r){var o={};h.forEach(function(k,i){o[k]=r[i];});return o;});}
function saveUser(data){
  var ctx=requirePermission_(AUTH_PERMISSIONS.USERS);
  var email=String(data.email||'').trim().toLowerCase(),name=String(data.name||'').trim(),role=String(data.role||AUTH_ROLES.EMPLOYEE).trim(),status=String(data.status||'Active').trim();
  if(!email||email.indexOf('@')<1)throw new Error('A valid Google account email is required.');
  if(!ROLE_PERMISSIONS[role])throw new Error('Invalid role.');
  if(['Active','Inactive'].indexOf(status)<0)throw new Error('Invalid user status.');
  if(role===AUTH_ROLES.ADMIN&&email!==PRIMARY_ADMIN_EMAIL)throw new Error('The Admin role is reserved exclusively for the primary Admin account.');
  if(email===PRIMARY_ADMIN_EMAIL&&(role!==AUTH_ROLES.ADMIN||status!=='Active'))throw new Error('The primary Admin account cannot be downgraded or deactivated.');
  var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS),values=s.getDataRange().getValues(),h=values.shift(),emailIdx=h.indexOf('Email');
  for(var i=0;i<values.length;i++)if(String(values[i][emailIdx]||'').trim().toLowerCase()===email){
    var userId=String(values[i][0]||''),currentRole=String(values[i][2]||''),currentStatus=String(values[i][4]||'');
    if(userId===ctx.userId&&status==='Inactive')throw new Error('You cannot deactivate your own account.');
    if(currentRole===AUTH_ROLES.ADMIN&&email!==PRIMARY_ADMIN_EMAIL)throw new Error('Only the primary Admin account may hold the Admin role.');
    if(currentRole===AUTH_ROLES.ADMIN&&role!==AUTH_ROLES.ADMIN&&currentStatus==='Active'&&countActiveAdmins_(s)===1)throw new Error('The last active Admin cannot be downgraded.');
    if(currentRole===AUTH_ROLES.ADMIN&&role!==AUTH_ROLES.ADMIN&&status==='Inactive'&&countActiveAdmins_(s)===1)throw new Error('The last active Admin cannot be deactivated.');
    s.getRange(i+2,2,1,4).setValues([[name||email.split('@')[0],role,email,status]]);
    recordActivity_('User',userId,'Updated user: '+email+' · Role: '+role+' · Status: '+status);
    return{updated:true,email:email};
  }
  if(role===AUTH_ROLES.ADMIN&&status==='Inactive')throw new Error('The primary Admin account must be Active.');
  var newId=generateId_('USR');
  s.appendRow([newId,name||email.split('@')[0],role,email,status]);
  recordActivity_('User',newId,'Created user: '+email+' · Role: '+role+' · Status: '+status);
  return{created:true,email:email,userId:newId};
}
function updateUserStatus(userId,status){
  var ctx=requirePermission_(AUTH_PERMISSIONS.USERS);
  status=String(status);
  if(['Active','Inactive'].indexOf(status)<0)throw new Error('Invalid user status.');
  var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS),values=s.getDataRange().getValues(),h=values.shift(),idIdx=h.indexOf('User ID'),statusIdx=h.indexOf('Status'),emailIdx=h.indexOf('Email'),roleIdx=h.indexOf('Role');
  for(var i=0;i<values.length;i++)if(String(values[i][idIdx])===String(userId)){
    var targetRole=String(values[i][roleIdx]||''),targetEmail=String(values[i][emailIdx]||'').trim().toLowerCase();
    if(String(userId)===String(ctx.userId)&&status==='Inactive')throw new Error('You cannot deactivate your own account.');
    if(targetRole===AUTH_ROLES.ADMIN&&targetEmail!==PRIMARY_ADMIN_EMAIL)throw new Error('Only the primary Admin account may hold the Admin role.');
    if(targetRole===AUTH_ROLES.ADMIN&&targetEmail===PRIMARY_ADMIN_EMAIL&&status==='Inactive')throw new Error('The primary Admin account cannot be deactivated.');
    if(targetRole===AUTH_ROLES.ADMIN&&status==='Inactive'&&countActiveAdmins_(s)===1)throw new Error('The last active Admin cannot be deactivated.');
    s.getRange(i+2,statusIdx+1).setValue(status);
    recordActivity_('User',String(userId),'Changed user status: '+targetEmail+' → '+status);
    return true;
  }
  throw new Error('User not found.');
}
function countActiveAdmins_(sheet){
  if(!sheet||sheet.getLastRow()<2)return 0;
  var values=sheet.getDataRange().getValues(),headers=values.shift(),roleIdx=headers.indexOf('Role'),statusIdx=headers.indexOf('Status');
  var count=0;
  values.forEach(function(r){if(String(r[roleIdx]||'')===AUTH_ROLES.ADMIN&&String(r[statusIdx]||'')==='Active')count++;});
  return count;
}
function ensureAuthSetup_(){var s=getCRMSpreadsheet_().getSheetByName(APP.SHEETS.USERS);if(!s)return;var roleCol=s.getRange('C2:C');roleCol.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([AUTH_ROLES.ADMIN,AUTH_ROLES.PARTNER,AUTH_ROLES.MANAGER,AUTH_ROLES.EMPLOYEE],true).setAllowInvalid(false).build());var statusCol=s.getRange('E2:E');statusCol.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive'],true).setAllowInvalid(false).build());}
