// LeaveWorkflow.gs - submit, cancel, two-step approval, balance, calendar, notifications
function classify_(text) {
  var t = String(text).toLowerCase();
  var m = [['Emergency Leave', /emergency|accident|urgent|death|funeral/], ['Sick Leave', /sick|fever|ill|doctor|hospital|medical|flu|injur/],
           ['Annual Leave', /vacation|holiday|trip|travel|festival|annual/], ['Personal Leave', /personal|family|wedding|exam|home/]];
  for (var i = 0; i < m.length; i++) if (m[i][1].test(t)) return m[i][0];
  return 'Other';
}
function adj_(emp, type, dUsed, dPend) {
  var b = bal_(emp, type); if (!b) return;
  var used = Number(b.UsedLeaves) + dUsed, pen = Number(b.PendingLeaves) + dPend;
  upd_('LeaveBalance', b._row, { UsedLeaves: used, PendingLeaves: pen, RemainingLeaves: Number(b.TotalLeaves) - used - pen, LastUpdated: now_() });
}
function notify_(empId, msg, type) {
  add_('Notifications', { NotificationID: uid_('NTF'), EmployeeID: empId, Message: msg, NotificationType: type, IsRead: 'No', CreatedAt: now_() });
  try {
    var e = rows_('Employees').filter(function (x) { return x.EmployeeID === empId; })[0];
    if (e && e.Email) MailApp.sendEmail(e.Email, 'LeaveFlow: ' + type, msg);
  } catch (x) { /* quota/permission problems: in-app notification is the fallback */ }
}
function event_(r) {
  if (r.CalendarEventID) return r.CalendarEventID;
  try {
    var id = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
    var cal = (id && CalendarApp.getCalendarById(id)) || CalendarApp.getDefaultCalendar();
    var end = new Date(r.EndDate + 'T00:00:00'); end.setDate(end.getDate() + 1);
    var ev = cal.createAllDayEvent(r.EmployeeName, new Date(r.StartDate + 'T00:00:00'), end,
      { description: r.LeaveType + ': ' + r.StartDate + ' to ' + r.EndDate + ' (' + r.NumberOfDays + ' day(s))\nRequest ' + r.RequestID });
    return ev.getId();
  } catch (x) { audit_('SYSTEM', 'CALENDAR_FAIL', x.message); return ''; }
}

function submitLeave(t, d) {
  return wrap_(function () { return locked_(function () {
    var u = me_(t); d = d || {};
    if (u.Role === 'Admin') throw new Error('Admins cannot apply for leave');
    var type = String(d.type || ''), s = String(d.start || ''), e = String(d.end || ''), reason = String(d.reason || '').trim();
    if (!rows_('LeaveTypes').some(function (x) { return x.LeaveTypeName === type && x.IsActive === 'Yes'; })) throw new Error('Invalid leave type');
    if (!RX_.test(s) || !RX_.test(e)) throw new Error('Enter valid dates');
    if (s < today_()) throw new Error('Start date cannot be in the past');
    if (e < s) throw new Error('End date cannot be before start date');
    if (reason.length < 3 || reason.length > 500) throw new Error('Reason must be 3-500 characters');
    var n = days_(s, e), b = bal_(u.EmployeeID, type);
    if (!b) throw new Error('No leave balance found for this type');
    var left = Number(b.TotalLeaves) - Number(b.UsedLeaves) - Number(b.PendingLeaves);
    if (n > left) throw new Error('Insufficient balance: you requested ' + n + ' day(s) but only ' + left + ' remain');
    var live = rows_('LeaveRequests').filter(function (r) { return ACTIVE_.indexOf(r.Status) >= 0; });
    var clash = live.filter(function (r) { return r.EmployeeID === u.EmployeeID && r.StartDate <= e && s <= r.EndDate; })[0];
    if (clash) throw new Error('These dates overlap your ' + clash.Status + ' request ' + clash.RequestID);
    var emps = rows_('Employees');
    var admin = emps.filter(function (x) { return x.Role === 'Admin' && x.AccountStatus === 'Active'; })[0];
    var first = u.ManagerID || (admin && admin.EmployeeID);
    if (!first || first === u.EmployeeID) throw new Error('No approver is assigned to you. Contact the admin.');
    var fm = emps.filter(function (x) { return x.EmployeeID === first; })[0];
    var second = (fm && fm.ManagerID) || (admin && admin.EmployeeID) || '';
    if (second === first || second === u.EmployeeID) second = '';
    var mates = emps.filter(function (x) { return x.ManagerID && x.ManagerID === u.ManagerID && x.EmployeeID !== u.EmployeeID; }).map(function (x) { return x.EmployeeID; });
    var tc = live.filter(function (r) { return mates.indexOf(r.EmployeeID) >= 0 && r.StartDate <= e && s <= r.EndDate; }).length;
    var id = uid_('REQ'), ts = now_();
    add_('LeaveRequests', { RequestID: id, EmployeeID: u.EmployeeID, EmployeeName: u.FullName, LeaveType: type, StartDate: s, EndDate: e, NumberOfDays: n, Reason: reason, AIClassification: classify_(reason), Status: 'Pending First Approval', FirstApproverID: first, FirstApprovalStatus: 'Pending', SecondApproverID: second, SecondApprovalStatus: second ? 'Pending' : 'N/A', CreatedAt: ts, UpdatedAt: ts, CurrentApproverID: first, DaysReserved: 'Yes' });
    adj_(u.EmployeeID, type, 0, n);
    add_('ApprovalHistory', { ApprovalID: uid_('APR'), RequestID: id, ApproverID: u.EmployeeID, ApprovalLevel: 0, Action: 'Submitted', Comments: '', ActionDate: ts });
    notify_(first, u.FullName + ' submitted a ' + type + ' request (' + s + ' to ' + e + ') awaiting your approval.', 'Leave request submitted');
    notify_(u.EmployeeID, 'Your ' + type + ' request ' + id + ' was submitted.', 'Leave request submitted');
    audit_(u.EmployeeID, 'SUBMIT_LEAVE', id);
    return { ok: true, id: id, warning: tc ? tc + ' teammate(s) also have leave on these dates (request not blocked).' : '' };
  }); });
}

function cancelLeave(t, id) {
  return wrap_(function () { return locked_(function () {
    var u = me_(t);
    var r = rows_('LeaveRequests').filter(function (x) { return x.RequestID === id; })[0];
    if (!r || r.EmployeeID !== u.EmployeeID) throw new Error('Request not found');
    if (r.Status.indexOf('Pending') !== 0) throw new Error('Only pending requests can be cancelled');
    if (r.DaysReserved === 'Yes') adj_(r.EmployeeID, r.LeaveType, 0, -Number(r.NumberOfDays));
    upd_('LeaveRequests', r._row, { Status: 'Cancelled', CurrentApproverID: '', DaysReserved: 'No', UpdatedAt: now_() });
    add_('ApprovalHistory', { ApprovalID: uid_('APR'), RequestID: id, ApproverID: u.EmployeeID, ApprovalLevel: 0, Action: 'Cancelled', Comments: '', ActionDate: now_() });
    notify_(r.CurrentApproverID, u.FullName + ' cancelled request ' + id + '.', 'Leave request cancelled');
    notify_(u.EmployeeID, 'You cancelled request ' + id + '.', 'Leave request cancelled');
    audit_(u.EmployeeID, 'CANCEL_LEAVE', id);
    return { ok: true };
  }); });
}

function decide(t, id, action, comment) {
  return wrap_(function () { return locked_(function () {
    var u = me_(t);
    if (u.Role === 'Employee') throw new Error('Not allowed');
    if (action !== 'Approve' && action !== 'Reject') throw new Error('Invalid action');
    var r = rows_('LeaveRequests').filter(function (x) { return x.RequestID === id; })[0];
    if (!r) throw new Error('Request not found');
    if (r.Status.indexOf('Pending') !== 0) throw new Error('Request is already ' + r.Status);
    if (r.CurrentApproverID !== u.EmployeeID) throw new Error('This request is not assigned to you');
    if (rows_('ApprovalHistory').some(function (h) { return h.RequestID === id && h.ApproverID === u.EmployeeID && (h.Action === 'Approved' || h.Action === 'Rejected'); }))
      throw new Error('You have already acted on this request');
    comment = String(comment || '').trim().slice(0, 300);
    if (action === 'Reject' && !comment) throw new Error('A comment is required when rejecting');
    var lvl = r.Status === 'Pending First Approval' ? 1 : 2, p = lvl === 1 ? 'First' : 'Second', ts = now_(), ch = {};
    var done = action === 'Approve' ? 'Approved' : 'Rejected';
    ch[p + 'ApprovalStatus'] = done; ch[p + 'ApprovalDate'] = ts; ch.UpdatedAt = ts;
    add_('ApprovalHistory', { ApprovalID: uid_('APR'), RequestID: id, ApproverID: u.EmployeeID, ApprovalLevel: lvl, Action: done, Comments: comment, ActionDate: ts });
    var n = Number(r.NumberOfDays);
    if (action === 'Reject') {
      if (r.DaysReserved === 'Yes') adj_(r.EmployeeID, r.LeaveType, 0, -n);
      ch.Status = 'Rejected'; ch.RejectionReason = comment; ch.CurrentApproverID = ''; ch.DaysReserved = 'No';
      notify_(r.EmployeeID, 'Your ' + r.LeaveType + ' request ' + id + ' was rejected: ' + comment, 'Leave request rejected');
    } else if (lvl === 1 && r.SecondApproverID) {
      ch.Status = 'Pending Second Approval'; ch.CurrentApproverID = r.SecondApproverID;
      notify_(r.SecondApproverID, r.EmployeeName + '\'s ' + r.LeaveType + ' request (' + r.StartDate + ' to ' + r.EndDate + ') awaits your final approval.', 'First-level approval completed');
      notify_(r.EmployeeID, 'Request ' + id + ' passed first approval and is with the second approver.', 'First-level approval completed');
    } else {
      if (r.DaysReserved === 'Yes') adj_(r.EmployeeID, r.LeaveType, n, -n);
      ch.Status = 'Approved'; ch.CurrentApproverID = ''; ch.DaysReserved = 'No';
      ch.CalendarEventID = event_(r);
      notify_(r.EmployeeID, 'Your ' + r.LeaveType + ' request ' + id + ' (' + r.StartDate + ' to ' + r.EndDate + ') is approved.', 'Leave request approved');
    }
    upd_('LeaveRequests', r._row, ch);
    audit_(u.EmployeeID, 'DECISION_' + done.toUpperCase(), id + ' L' + lvl);
    return { ok: true };
  }); });
}
