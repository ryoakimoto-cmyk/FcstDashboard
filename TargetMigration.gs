function TargetMigration_buildMonthlyMaster() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  var userSheet = ss.getSheetByName(SF_USER_SHEET_NAME);
  if (!targetSheet) throw new Error('目標シートが見つかりません');
  if (!userSheet) throw new Error('SFユーザーシートが見つかりません');

  var outName = MONTHLY_TARGET_MASTER_SHEET_NAME;
  var outSheet = ss.getSheetByName(outName);
  if (!outSheet) outSheet = ss.insertSheet(outName);
  outSheet.clearContents();

  var userValues = userSheet.getLastRow() > 2
    ? userSheet.getRange(3, 1, userSheet.getLastRow() - 2, 6).getValues()
    : [];
  var userMap = {};
  userValues.forEach(function(row) {
    var displayName = String(row[3] || '').trim();
    if (!displayName) return;
    userMap[displayName] = {
      displayUserName: displayName,
      sourceUserName: displayName,
      dept: String(row[1] || '').trim(),
      group: String(row[2] || '').trim(),
      sortPriority: typeof row[5] === 'number' ? row[5] : Number(row[5]) || 0
    };
  });

  var headers = [
    '対象月',
    '集計元ユーザー名',
    '表示用ユーザー名',
    'dept',
    'グループ',
    'sortOrder',
    'Net目標',
    'New+Exp目標',
    'Churn目標',
    '目標ソース',
    '備考'
  ];

  var targetValues = targetSheet.getLastRow() > 0
    ? targetSheet.getRange(1, 1, targetSheet.getLastRow(), Math.max(targetSheet.getLastColumn(), 11)).getValues()
    : [];
  var rowsByKey = {};

  targetValues.forEach(function(row) {
    var date = TargetMigration_parseDate_(row[0]);
    if (!date) return;
    var ym = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
    var orgName = String(row[6] || '').trim();
    var metricType = String(row[8] || '').trim();
    var targetValue = typeof row[9] === 'number' ? row[9] : Number(row[9]) || 0;
    if (!orgName) return;

    var userInfo = userMap[orgName] || null;
    var sourceUserName = userInfo ? userInfo.sourceUserName : '';
    var displayUserName = userInfo ? userInfo.displayUserName : orgName;
    var dept = userInfo ? userInfo.dept : '';
    var group = userInfo ? userInfo.group : orgName;
    var sortPriority = userInfo ? userInfo.sortPriority : 0;
    var note = userInfo ? '' : 'SFユーザーに該当なし';
    var key = [ym, sourceUserName || displayUserName].join('|');

    if (!rowsByKey[key]) {
      rowsByKey[key] = {
        ym: ym,
        sourceUserName: sourceUserName || displayUserName,
        displayUserName: displayUserName,
        dept: dept,
        group: group,
        sortPriority: sortPriority,
        net: '',
        newExp: '',
        churn: '',
        source: '',
        note: note
      };
    }

    if (metricType === 'Net') rowsByKey[key].net = targetValue;
    else if (metricType === 'New+Exp') rowsByKey[key].newExp = targetValue;
    else if (metricType === 'Churn') rowsByKey[key].churn = targetValue;

    rowsByKey[key].source = rowsByKey[key].source ? rowsByKey[key].source + ',目標' : '目標';
  });

  Object.keys(userMap).forEach(function(name) {
    var userInfo = userMap[name];
    Object.keys(rowsByKey).forEach(function(key) {
      var row = rowsByKey[key];
      if (row.displayUserName !== name) return;
      row.dept = row.dept || userInfo.dept;
      row.group = row.group || userInfo.group;
      row.sortPriority = row.sortPriority || userInfo.sortPriority;
    });
  });

  var output = Object.keys(rowsByKey).sort().map(function(key) {
    var row = rowsByKey[key];
    return [
      row.ym,
      row.sourceUserName,
      row.displayUserName,
      row.dept,
      row.group,
      row.sortPriority,
      row.net,
      row.newExp,
      row.churn,
      row.source,
      row.note
    ];
  });

  outSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (output.length) {
    outSheet.getRange(2, 1, output.length, headers.length).setValues(output);
  }
  outSheet.setFrozenRows(1);
  return { ok: true, sheetName: outName, count: output.length };
}

function TargetMigration_parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  var text = String(value || '').trim();
  if (!text || /Q/.test(text)) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  var date = new Date(text + 'T00:00:00');
  return isNaN(date) ? null : date;
}
