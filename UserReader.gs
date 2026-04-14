function UserReader_getUsers(deptKey) {
  var sheet = UserReader_getSheet_(deptKey);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return {};

  var values = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  return values.reduce(function(map, row) {
    var name = String(row[3] || '').trim();
    if (!name) return map;

    map[name] = {
      group: String(row[2] || '').trim(),
      dept: String(row[1] || '').trim(),
      sortOrder: UserReader_toNumber_(row[5])
    };
    return map;
  }, {});
}

function UserReader_getSheet_(deptKey) {
  return getSharedSheet(SF_USER_SHEET_NAME);
}

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
  }).filter(function(r) {
    return r.email;
  });
}

function UserReader_getDeptUsers(deptKey) {
  var master = UserReader_getDeptMaster();
  return master.filter(function(r) { return r.deptKey === deptKey; });
}

function UserReader_getUserDefaultDept(email) {
  var rows = UserReader_getDeptMaster();
  var found = rows.filter(function(r) { return r.email === email; })[0];
  if (!found) return null;
  if (found.deptKey && DEPT_CONFIG[found.deptKey]) return found.deptKey;
  return null;
}

function UserReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}
