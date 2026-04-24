function UserReader_getDeptMaster() {
  var sheet = getSharedSheet(DEPT_MASTER_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(function(row) {
    return {
      email: String(row[0] || '').trim(),
      division: String(row[1] || '').trim(),
      department: String(row[2] || '').trim(),
      dept_key: String(row[3] || '').trim(),
      deptKey: String(row[3] || '').trim(),
      userName: String(row[4] || row[2] || '').trim(),
      displayName: String(row[4] || row[2] || '').trim()
    };
  }).filter(function(row) {
    return row.email;
  });
}

function UserReader_getDeptUsers(deptKey) {
  return UserReader_getDeptMaster().filter(function(row) {
    return row.deptKey === deptKey;
  });
}

function UserReader_getUserDefaultDept(email) {
  var found = UserReader_getDeptMaster().filter(function(row) {
    return row.email === email;
  })[0];
  if (!found) return null;
  if (found.deptKey && isValidDeptKey_(found.deptKey)) return found.deptKey;
  return null;
}
