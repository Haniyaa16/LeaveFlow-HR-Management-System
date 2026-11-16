// Code.gs - entry point, helpers, authentication, dashboard data
// NOTE: functions ending in "_" are private (not callable from the browser).
var TZ_ = Session.getScriptTimeZone();
var RX_ = /^\d{4}-\d{2}-\d{2}$/;
var ACTIVE_ = ['Pending First Approval', 'Pending Second Approval', 'Approved'];
var H_ = {
  Employees: ['EmployeeID','FullName','Email','Salt','PasswordHash','Department','ManagerID','Role','JoiningDate','AccountStatus'],
  LeaveRequests: ['RequestID','EmployeeID','EmployeeName','LeaveType','StartDate','EndDate','NumberOfDays','Reason','AIClassification','Status','FirstApproverID','FirstApprovalStatus','FirstApprovalDate','SecondApproverID','SecondApprovalStatus','SecondApprovalDate','RejectionReason','CalendarEventID','CreatedAt','UpdatedAt','CurrentApproverID','DaysReserved'],
  LeaveBalance: ['BalanceID','EmployeeID','LeaveType','TotalLeaves','UsedLeaves','PendingLeaves','RemainingLeaves','LastUpdated'],
  LeaveTypes: ['LeaveTypeID','LeaveTypeName','Description','DefaultAnnualAllowance','IsActive'],
  ApprovalHistory: ['ApprovalID','RequestID','ApproverID','ApprovalLevel','Action','Comments','ActionDate'],
  Notifications: ['NotificationID','EmployeeID','Message','NotificationType','IsRead','CreatedAt'],
  AuditLogs: ['LogID','UserID','Action','Description','Timestamp']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LeaveFlow HR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- helpers ----------
function db_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) {
  var v = db_(n).getDataRange().getValues(), h = v.shift();
  return v.map(function (r, i) {
    var o = { _row: i + 2 };
    h.forEach(function (k, j) { var x = r[j]; o[k] = x instanceof Date ? Utilities.formatDate(x, TZ_, 'yyyy-MM-dd') : x; });
    return o;
  });
}
function add_(n, o) {
  var s = db_(n), h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  s.appendRow(h.map(function (k) { return o[k] === undefined ? '' : o[k]; }));
}
function upd_(n, row, o) {
  var s = db_(n), h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  Object.keys(o).forEach(function (k) { var c = h.indexOf(k); if (c >= 0) s.getRange(row, c + 1).setValue(o[k]); });
}
function uid_(p) { return p + '-' + Utilities.getUuid().slice(0, 8).toUpperCase(); }
function now_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd'); }
function days_(s, e) { return Math.round((new Date(e + 'T00:00:00') - new Date(s + 'T00:00:00')) / 864e5) + 1; }
function wrap_(fn) { try { return fn(); } catch (e) { return { ok: false, error: e.message }; } }
function locked_(fn) {
  var l = LockService.getScriptLock(); l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
function audit_(user, action, desc) {
  add_('AuditLogs', { LogID: uid_('LOG'), UserID: user, Action: action, Description: desc || '', Timestamp: now_() });
}
function pub_(u) { return { EmployeeID: u.EmployeeID, FullName: u.FullName, Email: u.Email, Department: u.Department, ManagerID: u.ManagerID, Role: u.Role, JoiningDate: u.JoiningDate, AccountStatus: u.AccountStatus }; }
function hash_(p, salt) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + p, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function me_(t) {
  var id = t && CacheService.getScriptCache().get('t:' + t);
  if (!id) throw new Error('Session expired. Please sign in again.');
  var u = rows_('Employees').filter(function (e) { return e.EmployeeID === id; })[0];
  if (!u || u.AccountStatus !== 'Active') throw new Error('Session expired. Please sign in again.');
  return u;
}

// ---------- authentication ----------
function login(email, pw) {
  return wrap_(function () {
    email = String(email || '').trim().toLowerCase(); pw = String(pw || '');
    var u = rows_('Employees').filter(function (e) { return String(e.Email).toLowerCase() === email; })[0];
    if (!u || u.AccountStatus !== 'Active' || hash_(pw, u.Salt) !== u.PasswordHash) {
      audit_('-', 'LOGIN_FAIL', email);
      return { ok: false, error: 'Invalid email or password' };
    }
    var t = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('t:' + t, u.EmployeeID, 21600);
    audit_(u.EmployeeID, 'LOGIN', '');
    return { ok: true, token: t };
  });
}

// ---------- dashboard data (scoped by role) ----------
function getData(t) {
  return wrap_(function () {
    var u = me_(t), emps = rows_('Employees'), all = rows_('LeaveRequests');
    var team = emps.filter(function (e) { return e.ManagerID === u.EmployeeID; }).map(function (e) { return e.EmployeeID; });
    var mine = all.filter(function (r) { return r.EmployeeID === u.EmployeeID; });
    var scope = u.Role === 'Admin' ? all : u.Role === 'Manager' ? all.filter(function (r) { return team.indexOf(r.EmployeeID) >= 0 || r.EmployeeID === u.EmployeeID; }) : mine;
    var out = {
      ok: true, user: pub_(u), mine: mine, scope: scope,
      pending: all.filter(function (r) { return r.CurrentApproverID === u.EmployeeID && r.Status.indexOf('Pending') === 0; }),
      balances: rows_('LeaveBalance').filter(function (b) { return b.EmployeeID === u.EmployeeID; }),
      notes: rows_('Notifications').filter(function (n) { return n.EmployeeID === u.EmployeeID; }).slice(-8).reverse(),
      types: rows_('LeaveTypes').filter(function (x) { return x.IsActive === 'Yes'; }).map(function (x) { return x.LeaveTypeName; }),
      totalEmployees: u.Role === 'Admin' ? emps.length : u.Role === 'Manager' ? team.length : null
    };
    if (u.Role === 'Admin') {
      out.employees = emps.map(pub_);
      out.audit = rows_('AuditLogs').slice(-25).reverse();
    }
    return out;
  });
}

// ---------- admin ----------
function addEmployee(t, d) {
  return wrap_(function () { return locked_(function () {
    var u = me_(t); if (u.Role !== 'Admin') throw new Error('Admin only');
    d = d || {};
    var name = String(d.name || '').trim(), email = String(d.email || '').trim().toLowerCase();
    if (name.length < 2) throw new Error('Enter a valid name');
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid email');
    if (['Employee', 'Manager', 'Admin'].indexOf(d.role) < 0) throw new Error('Invalid role');
    var emps = rows_('Employees');
    if (emps.some(function (e) { return String(e.Email).toLowerCase() === email; })) throw new Error('Email already exists');
    if (d.managerId && !emps.some(function (e) { return e.EmployeeID === d.managerId; })) throw new Error('Manager not found');
    var id = uid_('EMP'), salt = Utilities.getUuid();
    add_('Employees', { EmployeeID: id, FullName: name, Email: email, Salt: salt, PasswordHash: hash_('Welcome@123', salt), Department: String(d.dept || '').trim(), ManagerID: d.managerId || '', Role: d.role, JoiningDate: today_(), AccountStatus: 'Active' });
    if (d.role !== 'Admin') makeBalances_(id);
    audit_(u.EmployeeID, 'ADD_EMPLOYEE', id + ' ' + email);
    return { ok: true, id: id };
  }); });
}
function makeBalances_(empId) {
  rows_('LeaveTypes').filter(function (x) { return x.IsActive === 'Yes'; }).forEach(function (x) {
    add_('LeaveBalance', { BalanceID: uid_('BAL'), EmployeeID: empId, LeaveType: x.LeaveTypeName, TotalLeaves: Number(x.DefaultAnnualAllowance), UsedLeaves: 0, PendingLeaves: 0, RemainingLeaves: Number(x.DefaultAnnualAllowance), LastUpdated: now_() });
  });
}
function setBalance(t, empId, type, total) {
  return wrap_(function () { return locked_(function () {
    var u = me_(t); if (u.Role !== 'Admin') throw new Error('Admin only');
    total = Number(total);
    var b = bal_(empId, type); if (!b) throw new Error('Balance not found');
    if (!(total >= Number(b.UsedLeaves) + Number(b.PendingLeaves))) throw new Error('Total cannot be less than used + pending');
    upd_('LeaveBalance', b._row, { TotalLeaves: total, RemainingLeaves: total - b.UsedLeaves - b.PendingLeaves, LastUpdated: now_() });
    audit_(u.EmployeeID, 'SET_BALANCE', empId + ' ' + type + '=' + total);
    return { ok: true };
  }); });
}
function bal_(emp, type) {
  return rows_('LeaveBalance').filter(function (b) { return b.EmployeeID === emp && b.LeaveType === type; })[0];
}
