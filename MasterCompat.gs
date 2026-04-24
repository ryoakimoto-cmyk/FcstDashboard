function AssignmentMaster_build() {
  var setupResult = MasterSchema_setupSheets();
  var sheet = getSharedSheet(ASSIGNMENT_MASTER_SHEET_NAME);
  return {
    ok: true,
    stub: true,
    createdSheets: setupResult.createdSheets || [],
    updatedSheets: setupResult.updatedSheets || [],
    rowCount: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0
  };
}

function AssignmentMaster_getDepartmentUsers(deptKey) {
  var context = AssignmentMaster_getContext(deptKey);
  var names = {};

  Object.keys(context.users || {}).forEach(function(key) {
    var user = context.users[key] || {};
    var sourceUserName = String(user.sourceUserName || key || '').trim();
    if (sourceUserName) names[sourceUserName] = true;
  });

  return Object.keys(names).sort();
}

function MasterSchema_setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var specs = [
    {
      name: ORG_MASTER_SHEET_NAME,
      headers: [
        'group_code',
        'group_label',
        'department_code',
        'department_name',
        'division_code',
        'division_name',
        'start_month',
        'end_month',
        'display_flag'
      ]
    },
    {
      name: TARGET_MASTER_SHEET_NAME,
      headers: [
        'target_month',
        'group_code',
        'display_flag'
      ]
    },
    {
      name: ASSIGNMENT_MASTER_SHEET_NAME,
      headers: [
        'display_flag',
        'department_code',
        'department_name',
        'division_code',
        'division_name',
        'target_month',
        'group_code',
        'group_label',
        'source_user_name',
        'display_name',
        'sort_order',
        'target_net',
        'target_new_exp',
        'target_churn'
      ]
    }
  ];
  var createdSheets = [];
  var updatedSheets = [];

  specs.forEach(function(spec) {
    var status = MasterSchema_ensureSheet_(ss, spec.name, spec.headers);
    if (status.created) createdSheets.push(spec.name);
    if (status.updated) updatedSheets.push(spec.name);
  });

  return {
    ok: true,
    createdSheets: createdSheets,
    updatedSheets: updatedSheets
  };
}

function MasterSchema_ensureSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  var created = false;
  var updated = false;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  if (sheet.getLastColumn() < headers.length) {
    sheet.insertColumnsAfter(Math.max(sheet.getLastColumn(), 1), headers.length - sheet.getLastColumn());
  }

  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (MasterSchema_rowIsBlank_(currentHeaders)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    updated = true;
  }

  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }

  return {
    created: created,
    updated: updated
  };
}

function MasterSchema_rowIsBlank_(row) {
  return (row || []).every(function(cell) {
    return String(cell || '').trim() === '';
  });
}

function OrgMasterReader_getRows() {
  return MasterCompat_readSheetRows_(ORG_MASTER_SHEET_NAME, function(row, headerMap) {
    var groupCode = MasterCompat_readText_(row, headerMap, ['group_code', 'groupCode', 'グループコード', 'group']);
    var departmentCode = MasterCompat_readText_(row, headerMap, ['department_code', 'departmentCode', 'dept', '部署コード']);
    var divisionCode = MasterCompat_readText_(row, headerMap, ['division_code', 'divisionCode', 'division', '本部コード']);
    if (!groupCode || !departmentCode || !divisionCode) return null;

    return {
      groupCode: groupCode,
      groupLabel: MasterCompat_readText_(row, headerMap, ['group_label', 'groupLabel', 'グループ', 'グループ名']) || groupCode,
      departmentCode: departmentCode,
      departmentName: MasterCompat_readText_(row, headerMap, ['department_name', 'departmentName', '部署名']) || departmentCode,
      divisionCode: divisionCode,
      divisionName: MasterCompat_readText_(row, headerMap, ['division_name', 'divisionName', '本部名']) || divisionCode,
      startMonth: MasterCompat_readMonth_(row, headerMap, ['start_month', 'startMonth', '開始月', '適用開始月']) || '0000-01',
      endMonth: MasterCompat_readMonth_(row, headerMap, ['end_month', 'endMonth', '終了月', '適用終了月']) || '9999-12',
      displayFlag: MasterCompat_readBoolean_(row, headerMap, ['display_flag', '表示フラグ'], true)
    };
  });
}

function TargetMasterReader_getRows() {
  return MasterCompat_readSheetRows_(TARGET_MASTER_SHEET_NAME, function(row, headerMap) {
    var groupCode = MasterCompat_readText_(row, headerMap, ['group_code', 'groupCode', 'グループコード', 'group']);
    var targetMonth = MasterCompat_readMonth_(row, headerMap, ['target_month', 'targetMonth', '対象月']);
    if (!groupCode || !targetMonth) return null;

    return {
      groupCode: groupCode,
      targetMonth: targetMonth,
      displayFlag: MasterCompat_readBoolean_(row, headerMap, ['display_flag', '表示フラグ'], true)
    };
  });
}

function MasterCompat_readSheetRows_(sheetName, mapper) {
  var sheet = getSharedSheet(sheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = AssignmentMaster_buildHeaderMap_(headers);
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.reduce(function(rows, row) {
    var mapped = mapper(row, headerMap);
    if (mapped) rows.push(mapped);
    return rows;
  }, []);
}

function MasterCompat_readText_(row, headerMap, keys) {
  return AssignmentMaster_formatCell_(AssignmentMaster_valueByKeys_(row, headerMap, keys)).trim();
}

function MasterCompat_readMonth_(row, headerMap, keys) {
  return AssignmentMaster_normalizeMonth_(AssignmentMaster_valueByKeys_(row, headerMap, keys));
}

function MasterCompat_readBoolean_(row, headerMap, keys, defaultValue) {
  if (!MasterCompat_hasAnyKey_(headerMap, keys)) return !!defaultValue;
  return AssignmentMaster_toBoolean_(AssignmentMaster_valueByKeys_(row, headerMap, keys));
}

function MasterCompat_hasAnyKey_(headerMap, keys) {
  return (keys || []).some(function(key) {
    return Object.prototype.hasOwnProperty.call(headerMap || {}, key);
  });
}

function AssignmentMaster_buildHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function(header, index) {
    var key = String(header == null ? '' : header).trim();
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = index;
    }
  });
  return map;
}

function AssignmentMaster_valueByKeys_(row, headerMap, keys) {
  if (!row || !headerMap || !keys) return '';
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      var index = headerMap[key];
      if (index >= 0 && index < row.length) return row[index];
    }
  }
  return '';
}

function AssignmentMaster_formatCell_(value) {
  if (value == null) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value);
}

function AssignmentMaster_toBoolean_(value) {
  if (value === true) return true;
  if (value === false || value == null) return false;
  if (typeof value === 'number') return value !== 0;
  var text = String(value).trim().toLowerCase();
  if (!text) return false;
  return text === 'true' || text === '1' || text === 'yes' || text === 'y' ||
         text === 'on' || text === '有効' || text === '表示' || text === '○' || text === '◯';
}

function AssignmentMaster_normalizeMonth_(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    var year = value.getFullYear();
    var month = value.getMonth() + 1;
    return year + '-' + (month < 10 ? '0' + month : String(month));
  }
  var text = String(value).trim();
  if (!text) return '';
  var m = text.match(/^(\d{4})[-\/\.年](\d{1,2})/);
  if (m) {
    var mm = parseInt(m[2], 10);
    if (mm >= 1 && mm <= 12) {
      return m[1] + '-' + (mm < 10 ? '0' + mm : String(mm));
    }
  }
  var m2 = text.match(/^(\d{4})(\d{2})$/);
  if (m2) {
    var mm2 = parseInt(m2[2], 10);
    if (mm2 >= 1 && mm2 <= 12) return m2[1] + '-' + m2[2];
  }
  return '';
}

function invalidateDeptCaches_(deptKey, options) {
  CacheLayer_invalidate(deptKey);
  return {
    ok: true,
    deptKey: String(deptKey || '').trim(),
    options: options || {}
  };
}

// Opportunity snapshots use group_name as deptKey.
// Org master does not have display_flag.
function OrgMasterReader_getRows() {
  return MasterCompat_readSheetRows_(ORG_MASTER_SHEET_NAME, function(row, headerMap) {
    var groupName = MasterCompat_readText_(row, headerMap, ['group_name', 'groupName']);
    var groupCode = MasterCompat_readText_(row, headerMap, ['group_code', 'groupCode', 'group']);
    var departmentCode = MasterCompat_readText_(row, headerMap, ['department_code', 'departmentCode', 'dept']);
    var divisionCode = MasterCompat_readText_(row, headerMap, ['division_code', 'divisionCode', 'division']);
    if (!groupName || !departmentCode || !divisionCode) return null;

    return {
      groupName: groupName,
      groupCode: groupCode,
      groupLabel: groupName,
      departmentCode: departmentCode,
      departmentName: MasterCompat_readText_(row, headerMap, ['department_name', 'departmentName']) || departmentCode,
      divisionCode: divisionCode,
      divisionName: MasterCompat_readText_(row, headerMap, ['division_name', 'divisionName']) || divisionCode,
      startMonth: MasterCompat_readMonth_(row, headerMap, ['start_month', 'startMonth']) || '0000-01',
      endMonth: MasterCompat_readMonth_(row, headerMap, ['end_month', 'endMonth']) || '9999-12'
    };
  });
}
