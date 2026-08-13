/**
 * Central authorization policy for HYPEMARK CRM.
 * System roles are only: Admin, Partner, Employee.
 * Employee Job Function narrows operational access.
 */
var JOB_FUNCTIONS = ['SMM','Editor','DOP','Graphic Designer','DOP + Editor'];
var AUTH_ACTIONS = { VIEW:'view', CREATE:'create', EDIT:'edit', ASSIGN:'assign', APPROVE:'approve', DELETE:'delete' };

var ROLE_PERMISSION_MAP = {};
ROLE_PERMISSION_MAP['Admin'] = ['dashboard','clients','projects','tasks','deliverables','content','finance_all','profitability','reports','activities_all','users','permissions','settings'];
ROLE_PERMISSION_MAP['Partner'] = ['dashboard','clients','projects','tasks','deliverables','content','finance_all','profitability','reports'];
ROLE_PERMISSION_MAP['Employee'] = ['dashboard','projects','tasks','deliverables','content'];

var JOB_PERMISSION_MAP = {};
JOB_PERMISSION_MAP['SMM'] = ['content','deliverables','tasks','projects'];
JOB_PERMISSION_MAP['Editor'] = ['content','deliverables','tasks','projects'];
JOB_PERMISSION_MAP['DOP'] = ['content','deliverables','tasks','projects'];
JOB_PERMISSION_MAP['Graphic Designer'] = ['content','deliverables','tasks','projects'];
JOB_PERMISSION_MAP['DOP + Editor'] = ['content','deliverables','tasks','projects'];

function rolePermissions_(role){ return (ROLE_PERMISSION_MAP[role] || []).slice(); }
function jobFunctionPermissions_(jobFunction){ return (JOB_PERMISSION_MAP[jobFunction] || []).slice(); }
function hasPermission_(ctx, permission){ return !!ctx && (ctx.permissions || []).indexOf(permission) >= 0; }
function requirePermission_(permission){
  var ctx = requireAuth_();
  if (!hasPermission_(ctx, permission)) throw new Error('You do not have permission to access this section.');
  return ctx;
}
function canManageTasks_(ctx){ return !!ctx && (ctx.role === 'Admin' || ctx.role === 'Partner' || (ctx.role === 'Employee' && ctx.jobFunction === 'SMM')); }
function canApproveTasks_(ctx, task){
  if (!ctx) return false;
  if (ctx.role === 'Admin' || ctx.role === 'Partner') return true;
  if (ctx.role === 'Employee' && ctx.jobFunction === 'SMM') return isContentRelatedTask_(task);
  return false;
}
function isContentRelatedTask_(task){ return !!task && !!String(task['Content ID'] || task['Deliverable ID'] || task['Project ID'] || '').trim(); }
function canSeeAssignedWork_(ctx, ownerName){
  if (!ctx) return false;
  if (ctx.role !== 'Employee') return true;
  return String(ownerName || '').trim().toLowerCase() === String(ctx.name || '').trim().toLowerCase();
}
function getJobFunctions(){ requirePermission_('users'); return JOB_FUNCTIONS.slice(); }
