function MonthlyMasterReader_getContext(deptKey) {
  var sheet = getSharedSheet(MONTHLY_TARGET_MASTER_SHEET_NAME);
  if (!sheet) return { users: {}, monthlyUsers: {}, targets: {} };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { users: {}, monthlyUsers: {}, targets: {} };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = MonthlyMasterReader_buildHeaderMap_(headers);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().filter(function(row) {
    var dept = MonthlyMasterReader_formatCell_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['dept', '担当部署'])).trim();
    return dept === deptKey;
  });
  if (!values.length) return { users: {}, monthlyUsers: {}, targets: {} };

  var users = {};
  var monthlyUsers = {};
  var targets = {};
  var latestMonthByUser = {};

  values.forEach(function(row) {
    var ym = MonthlyMasterReader_normalizeMonth_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['対象月']));
    if (!ym) return;

    var sourceUserName = MonthlyMasterReader_formatCell_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['集計用ユーザー名'])).trim();
    var displayUserName = MonthlyMasterReader_formatCell_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['表示用ユーザー名'])).trim();
    if (!sourceUserName) sourceUserName = displayUserName;
    if (!displayUserName) displayUserName = sourceUserName;
    if (!sourceUserName) return;

    var dept = MonthlyMasterReader_formatCell_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['担当部署', 'dept'])).trim();
    var group = MonthlyMasterReader_formatCell_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['グループ'])).trim();
    var sortOrder = MonthlyMasterReader_toNumber_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['ソート順', 'sortOrder']));
    var target = {
      net: MonthlyMasterReader_toNumberOrBlank_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['Net目標'])),
      newExp: MonthlyMasterReader_toNumberOrBlank_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['New+Exp目標'])),
      churn: MonthlyMasterReader_toNumberOrBlank_(MonthlyMasterReader_valueByKeys_(row, headerMap, ['Churn目標']))
    };

    monthlyUsers[sourceUserName + '|' + ym] = {
      sourceUserName: sourceUserName,
      displayName: displayUserName,
      dept: dept,
      group: group,
      sortOrder: sortOrder
    };
    targets[sourceUserName + '|' + ym.replace('-', '')] = target;

    if (!users[sourceUserName] || ym >= latestMonthByUser[sourceUserName]) {
      users[sourceUserName] = {
        sourceUserName: sourceUserName,
        displayName: displayUserName,
        dept: dept,
        group: group,
        sortOrder: sortOrder
      };
      latestMonthByUser[sourceUserName] = ym;
    }
  });

  return { users: users, monthlyUsers: monthlyUsers, targets: targets };
}

function MonthlyMasterReader_buildHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function(header, idx) {
    var raw = String(header || '').trim();
    if (!raw) return;
    map[raw] = idx;
    map[MonthlyMasterReader_normalize_(raw)] = idx;
  });
  return map;
}

function MonthlyMasterReader_valueByKeys_(row, headerMap, keys) {
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var candidates = [key, MonthlyMasterReader_normalize_(key)];
    for (var j = 0; j < candidates.length; j++) {
      var idx = headerMap[candidates[j]];
      if (idx !== undefined) return row[idx];
    }
  }
  return '';
}

function MonthlyMasterReader_normalizeMonth_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var text = String(value || '').trim();
  var match = text.match(/(\d{4})[-\/年](\d{1,2})/);
  if (!match) return '';
  return match[1] + '-' + String(Number(match[2])).padStart(2, '0');
}

function MonthlyMasterReader_formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value === null || value === undefined ? '' : String(value);
}

function MonthlyMasterReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function MonthlyMasterReader_toNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  return Number(value) || 0;
}

function MonthlyMasterReader_normalize_(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[()（）]/g, '')
    .toLowerCase();
}
