function MonthlyMasterReader_getContext(deptKey) {
  var rows = MonthlyMasterReader_getRows_(deptKey);
  var users = {};
  var monthlyUsers = {};
  var targets = {};
  var latestMonthByUser = {};

  rows.forEach(function(row) {
    var sourceUserName = MonthlyMasterReader_getSourceUserName_(row.headerMap, row.values);
    if (!sourceUserName) return;

    var targetMonth = AssignmentMaster_normalizeMonth_(
      AssignmentMaster_valueByKeys_(row.values, row.headerMap, ['target_month', '対象月'])
    );
    var user = {
      sourceUserName: sourceUserName,
      displayName: MonthlyMasterReader_getDisplayName_(row.headerMap, row.values, sourceUserName),
      dept: MonthlyMasterReader_getDept_(row.headerMap, row.values, deptKey),
      group: MonthlyMasterReader_getGroupLabel_(row.headerMap, row.values),
      groupCode: MonthlyMasterReader_getGroupCode_(row.headerMap, row.values),
      sortOrder: MonthlyMasterReader_getSortOrder_(row.headerMap, row.values)
    };

    if (!users[sourceUserName] || (targetMonth && targetMonth >= (latestMonthByUser[sourceUserName] || ''))) {
      users[sourceUserName] = user;
      latestMonthByUser[sourceUserName] = targetMonth || latestMonthByUser[sourceUserName] || '';
    }

    if (targetMonth) {
      monthlyUsers[sourceUserName + '|' + targetMonth] = user;
      targets[sourceUserName + '|' + targetMonth.replace('-', '')] =
        MonthlyMasterReader_getTarget_(row.headerMap, row.values);
    }
  });

  return {
    users: users,
    monthlyUsers: monthlyUsers,
    targets: targets
  };
}

function MonthlyMasterReader_getRows_(deptKey) {
  var sheet = getSharedSheet(ASSIGNMENT_MASTER_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = AssignmentMaster_buildHeaderMap_(headers);
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return rows.reduce(function(list, values) {
    var rowDept = MonthlyMasterReader_getDept_(headerMap, values, '');
    var displayFlag = MonthlyMasterReader_getDisplayFlag_(headerMap, values);
    if (rowDept !== String(deptKey || '').trim()) return list;
    if (!displayFlag) return list;
    list.push({ headerMap: headerMap, values: values });
    return list;
  }, []);
}

function MonthlyMasterReader_getDisplayFlag_(headerMap, row) {
  if (!MonthlyMasterReader_hasAnyKey_(headerMap, ['display_flag', '表示フラグ'])) return true;
  return AssignmentMaster_toBoolean_(
    AssignmentMaster_valueByKeys_(row, headerMap, ['display_flag', '表示フラグ'])
  );
}

function MonthlyMasterReader_getSourceUserName_(headerMap, row) {
  var sourceUserName = AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, [
      'source_user_name',
      'sourceUserName',
      '集計用ユーザー名',
      '担当者',
      'ユーザー名',
      'user_name'
    ])
  ).trim();
  if (sourceUserName) return sourceUserName;
  return AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, ['display_name', 'displayName', '表示用ユーザー名', '表示名'])
  ).trim();
}

function MonthlyMasterReader_getDisplayName_(headerMap, row, fallbackName) {
  return AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, [
      'display_name',
      'displayName',
      '表示用ユーザー名',
      '表示名',
      '担当者名'
    ])
  ).trim() || String(fallbackName || '').trim();
}

function MonthlyMasterReader_getDept_(headerMap, row, fallbackDeptKey) {
  return AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, [
      'department_code',
      'departmentCode',
      'dept',
      '担当部署'
    ])
  ).trim() || String(fallbackDeptKey || '').trim();
}

function MonthlyMasterReader_getGroupCode_(headerMap, row) {
  return AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, [
      'group_code',
      'groupCode',
      'グループコード'
    ])
  ).trim();
}

function MonthlyMasterReader_getGroupLabel_(headerMap, row) {
  return AssignmentMaster_formatCell_(
    AssignmentMaster_valueByKeys_(row, headerMap, [
      'group_label',
      'groupLabel',
      'group',
      'グループ',
      'グループ名'
    ])
  ).trim() || MonthlyMasterReader_getGroupCode_(headerMap, row);
}

function MonthlyMasterReader_getSortOrder_(headerMap, row) {
  return Number(AssignmentMaster_valueByKeys_(row, headerMap, [
    'sort_order',
    'sortOrder',
    'ソート順'
  ]) || 0) || 0;
}

function MonthlyMasterReader_getTarget_(headerMap, row) {
  return {
    net: MonthlyMasterReader_toNumber_(AssignmentMaster_valueByKeys_(row, headerMap, ['target_net', 'Net目標'])),
    newExp: MonthlyMasterReader_toNumber_(AssignmentMaster_valueByKeys_(row, headerMap, ['target_new_exp', 'New+Exp目標'])),
    churn: MonthlyMasterReader_toNumber_(AssignmentMaster_valueByKeys_(row, headerMap, ['target_churn', 'Churn目標']))
  };
}

function MonthlyMasterReader_hasAnyKey_(headerMap, keys) {
  return (keys || []).some(function(key) {
    return Object.prototype.hasOwnProperty.call(headerMap || {}, key);
  });
}

function MonthlyMasterReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}
