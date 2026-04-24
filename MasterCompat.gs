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

  if (!Object.keys(names).length) {
    UserReader_getDeptUsers(deptKey).forEach(function(row) {
      var sourceUserName = String(row && (row.userName || row.displayName) || '').trim();
      if (sourceUserName) names[sourceUserName] = true;
    });
  }

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

function invalidateDeptCaches_(deptKey, options) {
  CacheLayer_invalidate(deptKey);
  return {
    ok: true,
    deptKey: String(deptKey || '').trim(),
    options: options || {}
  };
}
