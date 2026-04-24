function OppListReader_getLastUpdated(deptKey) {
  var sheet = getSfDataSheet_(deptKey);
  if (!sheet) return '';
  return OppListReader_extractLastUpdated_(sheet.getRange(1, 1));
}

function OppListReader_getLiveRows(deptKey) {
  var sheet = getSfDataSheet_(deptKey);
  if (!sheet) return { rows: [], lastUpdated: '' };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return { rows: [], lastUpdated: '' };

  var deptFilter = OppListReader_getDeptFilter_(deptKey);
  if (deptFilter.error) {
    return {
      rows: [],
      lastUpdated: OppListReader_extractLastUpdated_(sheet.getRange(1, 1)),
      error: deptFilter.error
    };
  }

  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  if (isSscsDept_(deptKey)) headers = normalizeSSCSHeaders_(headers);
  var headerMap = OppListReader_buildHeaderMap_(headers);
  var values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  var rows = [];

  values.forEach(function(row) {
    var oppId = OppListReader_valueByKeys_(row, headerMap, ['ID', '案件ID']);
    var dealName = OppListReader_valueByKeys_(row, headerMap, ['案件名']);
    if (!oppId || !dealName) return;

    var rowDept = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['担当部署'])).trim();
    if (!OppListReader_matchesDept_(deptFilter, rowDept)) return;

    var subOwner = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['ユーザー', 'サブオーナー', '担当者'])).trim();

    oppId = String(oppId).trim();
    var proposalIds = {
      received: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_受領'])).trim(),
      debtMgmt: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_債権管理'])).trim(),
      debtMgmtLite: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_債権管理Lite', '提案商品ID_債権管理 Lite'])).trim(),
      expense: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_経費'])).trim()
    };

    rows.push({
      oppId: oppId,
      completedMonth: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['完了予定月'])),
      dept: rowDept,
      type: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['種別'])),
      subOwner: subOwner,
      phase: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['フェーズ'])),
      forecast: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['Forecast'])),
      scheduleOrCloseDate: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['予定日 / 確定日'])),
      dealName: OppListReader_formatCell_(dealName),
      allocationPercent: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['計上割合 (%)'])),
      mrr: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['MRR'])),
      initialCost: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['初期費用'])),
      keyDeal: OppListReader_toBoolean_(OppListReader_valueByKeys_(row, headerMap, ['KeyDeal_最新'])),
      fcstCommit: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)(換算値)'])),
      fcstMin: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['FCST(MIN)(換算値)'])),
      fcstMax: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['FCST(MAX)(換算値)'])),
      received: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['月額_受領'])),
      debtMgmt: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['月額_債権管理'])),
      debtMgmtLite: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['月額_債権管理 Lite'])),
      expense: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['月額_経費'])),
      proposalProductIds: proposalIds,
      fcstComment: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['FCSTコメント_最新'])),
      firstMeetingDate: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['初回営業日 / ｺﾝﾀｸﾄ日'])),
      summary: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['案件概要'])),
      optionFeatures: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['オプション機能'])),
      initiative: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['施策']))
    });
  });

  return {
    lastUpdated: OppListReader_extractLastUpdated_(sheet.getRange(1, 1)),
    rows: rows
  };
}

function OppListReader_getDeptFilter_(deptKey) {
  var row = OppListReader_getOrgDeptRow_(deptKey);
  if (!row) {
    return { error: '組織マスタにgroup_nameがありません dept=' + String(deptKey || '') };
  }

  var normalizedGroupName = OppListReader_normalizeDept_(row.groupName);
  if (!normalizedGroupName) {
    return { error: '組織マスタのgroup_nameが空です dept=' + String(deptKey || '') };
  }

  return {
    row: row,
    value: normalizedGroupName
  };
}

function OppListReader_getOrgDeptRow_(deptKey) {
  var key = OppListReader_normalizeDept_(deptKey);
  if (!key) return null;

  var monthKey = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  var rows = OrgMasterReader_getRows().filter(function(row) {
    if (!row) return false;
    if (OppListReader_normalizeDept_(row.groupName) !== key) return false;
    if (row.startMonth && row.startMonth > monthKey) return false;
    if (row.endMonth && row.endMonth < monthKey) return false;
    return true;
  }).sort(function(left, right) {
    return String(right.startMonth || '').localeCompare(String(left.startMonth || ''));
  });

  return rows.length ? rows[0] : null;
}

function OppListReader_matchesDept_(deptFilter, rowDept) {
  var normalized = OppListReader_normalizeDept_(rowDept);
  return !!(normalized && deptFilter && deptFilter.value === normalized);
}

function OppListReader_buildHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function(header, idx) {
    var raw = String(header || '').trim();
    if (!raw) return;
    map[raw] = idx;
    map[OppListReader_normalize_(raw)] = idx;
  });
  return map;
}

function OppListReader_valueByKeys_(row, headerMap, keys) {
  for (var i = 0; i < (keys || []).length; i++) {
    var key = keys[i];
    var candidates = [key, OppListReader_normalize_(key)];
    for (var j = 0; j < candidates.length; j++) {
      var idx = headerMap[candidates[j]];
      if (idx !== undefined) return row[idx];
    }
  }
  return '';
}

function OppListReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function OppListReader_toBoolean_(value) {
  if (value === true || value === false) return value;
  return /^(true|yes|1|y)$/i.test(String(value || '').trim());
}

function OppListReader_formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value === null || value === undefined ? '' : String(value);
}

function OppListReader_normalize_(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[()（）]/g, '')
    .toLowerCase();
}

function OppListReader_normalizeDept_(text) {
  return OppListReader_normalize_(text);
}

function OppListReader_extractLastUpdated_(cellOrText) {
  var formula = cellOrText && typeof cellOrText.getFormula === 'function' ? String(cellOrText.getFormula() || '') : '';
  var display = cellOrText && typeof cellOrText.getDisplayValue === 'function'
    ? String(cellOrText.getDisplayValue() || '')
    : String(cellOrText || '');
  var text = formula || display;
  var match = text.match(/(\d{4}[\/-]\d{1,2}[\/-]\d{1,2})\s+(\d{1,2}:\d{2})(?::\d{2})?/);
  if (!match) return '';
  var normalizedDate = match[1].replace(/\//g, '-');
  var normalizedTime = match[2];
  var date = new Date(normalizedDate + 'T' + normalizedTime + ':00');
  if (isNaN(date)) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
}
