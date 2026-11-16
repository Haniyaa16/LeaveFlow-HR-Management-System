// Setup.gs - run setupDatabase() ONCE from the editor. Edit SEED_ emails first.
// Demo password for all seeded users: Welcome@123
var SEED_ = [
  ['EMP-0001', 'Admin User',    'admin@example.com',     'HR',          '',         'Admin'],
  ['EMP-0002', 'Meera Manager', 'manager@example.com',   'Engineering', 'EMP-0001', 'Manager'],
  ['EMP-0003', 'Ravi Employee', 'employee1@example.com', 'Engineering', 'EMP-0002', 'Employee'],
  ['EMP-0004', 'Sana Employee', 'employee2@example.com', 'Engineering', 'EMP-0002', 'Employee']
];
function ownerOnly_() {
  if (Session.getActiveUser().getEmail() !== Session.getEffectiveUser().getEmail()) throw new Error('Owner only');
}
function setupDatabase() {
  ownerOnly_();
  var ss = SpreadsheetApp.getActive();
  Object.keys(H_).forEach(function (n) {
    if (ss.getSheetByName(n)) return;
    var s = ss.insertSheet(n);
    s.getRange(1, 1, 1, H_[n].length).setValues([H_[n]]).setFontWeight('bold').setBackground('#6d28d9').setFontColor('#ffffff');
    s.setFrozenRows(1);
  });
  var d = ss.getSheetByName('Sheet1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  if (rows_('LeaveTypes').length === 0) {
    [['Annual Leave', 'Yearly vacation', 20], ['Sick Leave', 'Illness or medical', 12], ['Personal Leave', 'Personal matters', 5], ['Emergency Leave', 'Urgent situations', 3]]
      .forEach(function (x, i) { add_('LeaveTypes', { LeaveTypeID: 'LT-' + (i + 1), LeaveTypeName: x[0], Description: x[1], DefaultAnnualAllowance: x[2], IsActive: 'Yes' }); });
  }
  if (rows_('Employees').length === 0) {
    SEED_.forEach(function (x) {
      var salt = Utilities.getUuid();
      add_('Employees', { EmployeeID: x[0], FullName: x[1], Email: x[2], Salt: salt, PasswordHash: hash_('Welcome@123', salt), Department: x[3], ManagerID: x[4], Role: x[5], JoiningDate: today_(), AccountStatus: 'Active' });
      if (x[5] !== 'Admin') makeBalances_(x[0]);
    });
  }
  Logger.log('Setup complete. Sign in with your SEED_ emails and password Welcome@123');
}
function setCalendarId(id) { // run from editor: setCalendarId('xxxx@group.calendar.google.com')
  ownerOnly_();
  PropertiesService.getScriptProperties().setProperty('CALENDAR_ID', id);
}
