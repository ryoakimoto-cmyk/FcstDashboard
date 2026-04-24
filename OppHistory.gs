var OPP_HISTORY_V2_SCHEMA = 'v2';
var OPP_HISTORY_TARGET = 'opp';
var OPP_HISTORY_STATUS_REMOVED_FROM_P = 'removed_from_p';
var OPP_HISTORY_TIMEZONE = 'Asia/Tokyo';

var OPP_WEBAPP_INPUT_HEADERS = ['opp_id', 'key_deal', 'fcst_comment', 'updated_at', 'updated_by'];
var OPP_HISTORY_V2_HEADERS = ['snapshot_at', 'snapshot_date', 'snapshot_type', 'dept', 'opp_id', '_status', 'payload_json'];
var SNAPSHOT_META_HEADERS = ['snapshot_date', 'snapshot_type', 'dept', 'target', '実行開始', '実行終了', 'ステータス', '件数_継続', '件数_新規', '件数_脱落', '件数_合計', '前週比_合計', 'エラー内容'];
var FCST_DECLARATION_CURRENT_HEADERS = ['dept', '階層', 'entity_id', '期間', 'mode', 'manual_value', 'note', 'updated_at', 'updated_by'];

function OppHistory_createWeekly_(deptKey) {
  var infra = OppHistory_ensureInfrastructure_();
  var now = new Date();
  var snapshotDate = OppHistory_formatDate_(now);
  var snapshotType = OppHistory_getSnapshotType_(now);
  var timestamp = OppHistory_formatTimestamp_(now);
  var deptKeys = deptKey ? [deptKey] : Object.keys(DEPT_CONFIG);
  var results = [];
  var errors = [];

  deptKeys.forEach(function(key) {
    if (!DEPT_CONFIG[key]) {
      errors.push({ dept: key, error: 'Unknown deptKey: ' + key });
      return;
    }

    var metaRow = OppHistory_appendMetaStart_(infra.metaSheet, snapshotDate, snapshotType, key, timestamp);
    try {
      var result = OppHistory_processDept_(key, snapshotDate, snapshotType, timestamp, infra);
      OppHistory_updateMetaSuccess_(infra.metaSheet, metaRow, result, timestamp);
      results.push(result);
    } catch (e) {
      var message = String(e && e.stack ? e.stack : (e && e.message ? e.message : e));
      OppHistory_updateMetaFailure_(infra.metaSheet, metaRow, message, timestamp);
      errors.push({ dept: key, error: message });
    }
  });

  var trimmedCount = OppHistory_trimOld_(infra.historySheet, snapshotDate);
  return {
    ok: errors.length === 0,
    snapshotDate: snapshotDate,
    snapshotType: snapshotType,
    results: results,
    errors: errors,
    trimmedCount: trimmedCount
  };
}

function OppHistory_getSnapshotDates_(deptKey) {
  var sheet = getSharedSheet(OPP_HISTORY_V2_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var seen = {};
  var dates = [];
  values.forEach(function(row) {
    var dateStr = String(row[1] || '').trim();
    var dept = String(row[3] || '').trim();
    if (!dateStr || dept !== deptKey || seen[dateStr]) return;
    seen[dateStr] = true;
    dates.push(dateStr);
  });
  dates.sort(function(a, b) { return a > b ? -1 : a < b ? 1 : 0; });
  return dates;
}

function OppHistory_getByDate_(deptKey, dateStr) {
  var sheet = getSharedSheet(OPP_HISTORY_V2_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { date: dateStr, rows: [] };

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, OPP_HISTORY_V2_HEADERS.length).getValues();
  var rows = [];
  values.forEach(function(row) {
    var rowDate = String(row[1] || '').trim();
    var rowDept = String(row[3] || '').trim();
    var status = String(row[5] || '').trim();
    if (rowDate !== dateStr || rowDept !== deptKey || status === OPP_HISTORY_STATUS_REMOVED_FROM_P) return;
    try {
      rows.push(OppHistory_payloadToLegacyRow_(JSON.parse(String(row[6] || '{}')), dateStr));
    } catch (e) {}
  });
  return { date: dateStr, rows: rows };
}

function OppHistory_ensureInfrastructure_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    spreadsheet: ss,
    webAppInputSheet: OppHistory_ensureSheet_(ss, OPP_WEBAPP_INPUT_SHEET_NAME, OPP_WEBAPP_INPUT_HEADERS),
    historySheet: OppHistory_ensureSheet_(ss, OPP_HISTORY_V2_SHEET_NAME, OPP_HISTORY_V2_HEADERS),
    metaSheet: OppHistory_ensureSheet_(ss, SNAPSHOT_META_SHEET_NAME, SNAPSHOT_META_HEADERS),
    fcstDeclarationSheet: OppHistory_ensureSheet_(ss, FCST_DECLARATION_CURRENT_SHEET_NAME, FCST_DECLARATION_CURRENT_HEADERS)
  };
}

function OppHistory_ensureSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var isBlank = current.every(function(value) { return value === '' || value === null; });
  if (isBlank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function OppHistory_processDept_(deptKey, snapshotDate, snapshotType, timestamp, infra) {
  var liveInfo = OppListReader_getLiveRows(deptKey) || {};
  var currentContext = OppHistory_buildCurrentContext_(liveInfo.rows || []);
  var webAppInputMap = OppHistory_getWebAppInputMap_(infra.webAppInputSheet);
  var previousSnapshotDate = OppHistory_shiftDateString_(snapshotDate, -7);
  var previousContext = OppHistory_getPreviousSnapshotContext_(infra.historySheet, deptKey, previousSnapshotDate);
  var tracking = OppHistory_determineTracking_(snapshotType, currentContext.map, currentContext.order, previousContext.activePayloadById);
  var rows = OppHistory_buildHistoryRows_(
    deptKey,
    snapshotDate,
    snapshotType,
    timestamp,
    currentContext,
    webAppInputMap,
    previousContext,
    tracking
  );

  OppHistory_deleteByDeptAndDate_(infra.historySheet, deptKey, snapshotDate);
  if (rows.length) {
    infra.historySheet
      .getRange(infra.historySheet.getLastRow() + 1, 1, rows.length, OPP_HISTORY_V2_HEADERS.length)
      .setValues(rows);
  }

  return {
    dept: deptKey,
    snapshotDate: snapshotDate,
    snapshotType: snapshotType,
    counts: {
      continued: tracking.continuedIds.length,
      newCount: tracking.newIds.length,
      dropped: tracking.dropoutIds.length,
      total: rows.length
    },
    previousTotalCount: previousContext.totalCount,
    sourceLastUpdated: liveInfo.lastUpdated || ''
  };
}

function OppHistory_buildCurrentContext_(rows) {
  var map = {};
  var order = [];
  (rows || []).forEach(function(row) {
    var oppId = String(row && row.oppId || '').trim();
    if (!oppId) return;
    if (!map[oppId]) order.push(oppId);
    map[oppId] = row;
  });
  return { map: map, order: order };
}

function OppHistory_determineTracking_(snapshotType, currentMap, currentOrder, previousActivePayloadById) {
  var previousTrackedIds = Object.keys(previousActivePayloadById || {});
  var continuedCandidates = previousTrackedIds.filter(function(oppId) {
    return !!currentMap[oppId];
  });
  var dropoutIds = snapshotType === 'monthly_dropout'
    ? continuedCandidates.filter(function(oppId) {
        return !OppHistory_phaseIncludesP_(currentMap[oppId] && currentMap[oppId].phase);
      })
    : [];
  var dropoutSet = OppHistory_toSet_(dropoutIds);
  var previousTrackedSet = OppHistory_toSet_(previousTrackedIds);
  var continuedIds = continuedCandidates.filter(function(oppId) {
    return !dropoutSet[oppId];
  });
  var continuedSet = OppHistory_toSet_(continuedIds);
  var trackedIds = [];
  var newIds = [];

  currentOrder.forEach(function(oppId) {
    if (continuedSet[oppId]) {
      trackedIds.push(oppId);
      return;
    }
    if (previousTrackedSet[oppId]) return;
    if (!OppHistory_phaseIncludesP_(currentMap[oppId] && currentMap[oppId].phase)) return;
    trackedIds.push(oppId);
    newIds.push(oppId);
  });

  return {
    continuedIds: continuedIds,
    newIds: newIds,
    trackedIds: trackedIds,
    dropoutIds: dropoutIds
  };
}

function OppHistory_buildHistoryRows_(deptKey, snapshotDate, snapshotType, timestamp, currentContext, webAppInputMap, previousContext, tracking) {
  var rows = [];
  var newIdSet = OppHistory_toSet_(tracking.newIds);

  tracking.trackedIds.forEach(function(oppId) {
    var row = currentContext.map[oppId];
    var previousPayload = previousContext.activePayloadById[oppId] || null;
    var payload = OppHistory_buildPayload_(
      deptKey,
      snapshotType,
      row,
      webAppInputMap[oppId] || null,
      previousPayload,
      !!newIdSet[oppId]
    );
    rows.push([
      timestamp,
      snapshotDate,
      snapshotType,
      deptKey,
      oppId,
      '',
      JSON.stringify(payload)
    ]);
  });

  tracking.dropoutIds.forEach(function(oppId) {
    var payload = OppHistory_buildDroppedPayload_(deptKey, snapshotType, previousContext.activePayloadById[oppId] || null);
    if (!payload) return;
    rows.push([
      timestamp,
      snapshotDate,
      snapshotType,
      deptKey,
      oppId,
      OPP_HISTORY_STATUS_REMOVED_FROM_P,
      JSON.stringify(payload)
    ]);
  });

  return rows;
}

function OppHistory_getWebAppInputMap_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, OPP_WEBAPP_INPUT_HEADERS.length).getValues();
  var map = {};
  values.forEach(function(row) {
    var oppId = String(row[0] || '').trim();
    if (!oppId) return;
    map[oppId] = {
      keyDeal: OppHistory_toBoolean_(row[1]),
      fcstComment: String(row[2] || ''),
      updatedAt: row[3] || '',
      updatedBy: String(row[4] || '')
    };
  });
  return map;
}

function OppHistory_getPreviousSnapshotContext_(sheet, deptKey, snapshotDate) {
  var context = {
    totalCount: 0,
    activePayloadById: {}
  };
  if (!sheet || sheet.getLastRow() < 2) return context;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, OPP_HISTORY_V2_HEADERS.length).getValues();
  values.forEach(function(row) {
    var rowDate = String(row[1] || '').trim();
    var rowDept = String(row[3] || '').trim();
    if (rowDate !== snapshotDate || rowDept !== deptKey) return;

    context.totalCount += 1;
    var oppId = String(row[4] || '').trim();
    var status = String(row[5] || '').trim();
    if (!oppId || status === OPP_HISTORY_STATUS_REMOVED_FROM_P) return;
    try {
      context.activePayloadById[oppId] = JSON.parse(String(row[6] || '{}'));
    } catch (e) {
      context.activePayloadById[oppId] = {};
    }
  });
  return context;
}

function OppHistory_buildPayload_(deptKey, snapshotType, row, webInput, previousPayload, isNew) {
  var proposalEnabled = isProposalProductsEnabled_(deptKey);
  var scheduleDate = OppHistory_nullIfEmpty_(row.scheduleDate || row.scheduleOrCloseDate);
  var closeDate = OppHistory_nullIfEmpty_(row.closeDate);
  return {
    _schema: OPP_HISTORY_V2_SCHEMA,
    _meta: {
      status: null,
      dept: deptKey,
      snapshot_type: snapshotType
    },
    opp_id: String(row.oppId || ''),
    '案件名': String(row.dealName || ''),
    '担当部署': String(row.dept || deptKey),
    '種別': String(row.type || ''),
    '案件副担当': String(row.subOwner || ''),
    '確度': OppHistory_numberOrNull_(row.confidence),
    'フェーズ': String(row.phase || ''),
    'フォーキャスト': String(row.forecast || ''),
    '完了予定月': OppHistory_nullIfEmpty_(row.completedMonth),
    '予定日': scheduleDate,
    '確定日': closeDate,
    '計上割合': OppHistory_numberOrNull_(row.allocationPercent),
    'MRR': OppHistory_numberOrNull_(row.mrr),
    '初期費用': OppHistory_numberOrNull_(row.initialCost),
    'FCSTコミット': OppHistory_numberOrNull_(row.fcstCommit),
    'FCSTMIN': OppHistory_numberOrNull_(row.fcstMin),
    'FCSTMAX': OppHistory_numberOrNull_(row.fcstMax),
    '受領': proposalEnabled ? OppHistory_numberOrNull_(row.received) : null,
    '債権管理': proposalEnabled ? OppHistory_numberOrNull_(row.debtMgmt) : null,
    '債権管理Lite': proposalEnabled ? OppHistory_numberOrNull_(row.debtMgmtLite) : null,
    '経費': proposalEnabled ? OppHistory_numberOrNull_(row.expense) : null,
    '初回商談日': OppHistory_nullIfEmpty_(row.firstMeetingDate),
    '案件概要': String(row.summary || ''),
    'オプション機能': String(row.optionFeatures || ''),
    '施策': String(row.initiative || ''),
    'KeyDeal': webInput ? !!webInput.keyDeal : false,
    'FCSTコメント': webInput ? String(webInput.fcstComment || '') : '',
    'weekOverWeek': OppHistory_buildWeekOverWeek_(row, previousPayload, isNew)
  };
}

function OppHistory_buildDroppedPayload_(deptKey, snapshotType, previousPayload) {
  if (!previousPayload) return null;
  var payload = OppHistory_clone_(previousPayload);
  payload._schema = payload._schema || OPP_HISTORY_V2_SCHEMA;
  payload._meta = payload._meta || {};
  payload._meta.status = OPP_HISTORY_STATUS_REMOVED_FROM_P;
  payload._meta.dept = deptKey;
  payload._meta.snapshot_type = snapshotType;
  return payload;
}

function OppHistory_buildWeekOverWeek_(row, previousPayload, isNew) {
  if (isNew || !previousPayload) {
    return {
      isNew: true,
      amountDelta: null,
      phaseChanged: null,
      closeMonthSlipped: null,
      confidenceDelta: null,
      ownerChanged: null
    };
  }

  var previousAmount = OppHistory_numberOrZero_(previousPayload['MRR']);
  var currentAmount = OppHistory_numberOrZero_(row.mrr);
  var previousPhase = String(previousPayload['フェーズ'] || '');
  var currentPhase = String(row.phase || '');
  var previousCloseMonth = String(previousPayload['完了予定月'] || '');
  var currentCloseMonth = String(row.completedMonth || '');
  var previousOwner = String(previousPayload['案件副担当'] || '');
  var currentOwner = String(row.subOwner || '');
  var previousConfidence = OppHistory_numberOrNull_(previousPayload['確度']);
  var currentConfidence = OppHistory_numberOrNull_(row.confidence);

  return {
    isNew: false,
    amountDelta: currentAmount - previousAmount,
    phaseChanged: previousPhase === currentPhase ? null : { from: previousPhase || null, to: currentPhase || null },
    closeMonthSlipped: previousCloseMonth === currentCloseMonth ? null : { from: previousCloseMonth || null, to: currentCloseMonth || null },
    confidenceDelta: previousConfidence === null || currentConfidence === null ? null : currentConfidence - previousConfidence,
    ownerChanged: previousOwner === currentOwner ? null : { from: previousOwner || null, to: currentOwner || null }
  };
}

function OppHistory_payloadToLegacyRow_(payload, snapshotDate) {
  var scheduleDate = OppHistory_nullIfEmpty_(payload['予定日']);
  var closeDate = OppHistory_nullIfEmpty_(payload['確定日']);
  return {
    oppId: String(payload.opp_id || ''),
    completedMonth: OppHistory_stringOrEmpty_(payload['完了予定月']),
    dept: OppHistory_stringOrEmpty_(payload['担当部署']),
    type: OppHistory_stringOrEmpty_(payload['種別']),
    subOwner: OppHistory_stringOrEmpty_(payload['案件副担当']),
    phase: OppHistory_stringOrEmpty_(payload['フェーズ']),
    forecast: OppHistory_stringOrEmpty_(payload['フォーキャスト']),
    scheduleDate: scheduleDate || '',
    closeDate: closeDate || '',
    scheduleOrCloseDate: scheduleDate || closeDate || '',
    confidence: OppHistory_numberOrNull_(payload['確度']),
    dealName: OppHistory_stringOrEmpty_(payload['案件名']),
    allocationPercent: OppHistory_numberOrZero_(payload['計上割合']),
    mrr: OppHistory_numberOrZero_(payload['MRR']),
    initialCost: OppHistory_numberOrZero_(payload['初期費用']),
    keyDeal: !!payload['KeyDeal'],
    fcstCommit: OppHistory_numberOrZero_(payload['FCSTコミット']),
    fcstMin: OppHistory_numberOrZero_(payload['FCSTMIN']),
    fcstMax: OppHistory_numberOrZero_(payload['FCSTMAX']),
    received: OppHistory_numberOrZero_(payload['受領']),
    debtMgmt: OppHistory_numberOrZero_(payload['債権管理']),
    debtMgmtLite: OppHistory_numberOrZero_(payload['債権管理Lite']),
    expense: OppHistory_numberOrZero_(payload['経費']),
    proposalProductIds: { received: '', debtMgmt: '', debtMgmtLite: '', expense: '' },
    fcstComment: OppHistory_stringOrEmpty_(payload['FCSTコメント']),
    firstMeetingDate: OppHistory_stringOrEmpty_(payload['初回商談日']),
    summary: OppHistory_stringOrEmpty_(payload['案件概要']),
    optionFeatures: OppHistory_stringOrEmpty_(payload['オプション機能']),
    initiative: OppHistory_stringOrEmpty_(payload['施策']),
    snapshotDate: snapshotDate || '',
    weekOverWeek: payload.weekOverWeek || null
  };
}

function OppHistory_appendMetaStart_(sheet, snapshotDate, snapshotType, deptKey, timestamp) {
  var rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, SNAPSHOT_META_HEADERS.length).setValues([[
    snapshotDate,
    snapshotType,
    deptKey,
    OPP_HISTORY_TARGET,
    timestamp,
    '',
    '実行中',
    '',
    '',
    '',
    '',
    '',
    ''
  ]]);
  return rowIndex;
}

function OppHistory_updateMetaSuccess_(sheet, rowIndex, result, timestamp) {
  var previousTotal = Number(result.previousTotalCount || 0);
  var currentTotal = Number(result.counts.total || 0);
  var ratio = previousTotal > 0 ? (currentTotal - previousTotal) / previousTotal : '';
  sheet.getRange(rowIndex, 6, 1, 8).setValues([[
    timestamp,
    '成功',
    result.counts.continued,
    result.counts.newCount,
    result.counts.dropped,
    currentTotal,
    ratio,
    ''
  ]]);
}

function OppHistory_updateMetaFailure_(sheet, rowIndex, errorMessage, timestamp) {
  sheet.getRange(rowIndex, 6, 1, 8).setValues([[
    timestamp,
    '失敗',
    '',
    '',
    '',
    '',
    '',
    errorMessage
  ]]);
}

function OppHistory_deleteByDeptAndDate_(sheet, deptKey, snapshotDate) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  for (var idx = values.length - 1; idx >= 0; idx--) {
    var rowDate = String(values[idx][1] || '').trim();
    var rowDept = String(values[idx][3] || '').trim();
    if (rowDate === snapshotDate && rowDept === deptKey) {
      sheet.deleteRow(idx + 2);
    }
  }
}

function OppHistory_trimOld_(sheet, snapshotDate) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var thresholdDate = OppHistory_shiftDateString_(snapshotDate, -91);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var deleted = 0;
  for (var idx = values.length - 1; idx >= 0; idx--) {
    var rowDate = String(values[idx][1] || '').trim();
    if (rowDate && rowDate < thresholdDate) {
      sheet.deleteRow(idx + 2);
      deleted += 1;
    }
  }
  return deleted;
}

function OppHistory_getSnapshotType_(date) {
  var day = Number(Utilities.formatDate(date, OPP_HISTORY_TIMEZONE, 'd'));
  return Math.ceil(day / 7) === 2 ? 'monthly_dropout' : 'weekly';
}

function OppHistory_phaseIncludesP_(phase) {
  return String(phase || '').toLowerCase().indexOf('p') !== -1;
}

function OppHistory_formatDate_(date) {
  return Utilities.formatDate(date, OPP_HISTORY_TIMEZONE, 'yyyy-MM-dd');
}

function OppHistory_formatTimestamp_(date) {
  return Utilities.formatDate(date, OPP_HISTORY_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssZ");
}

function OppHistory_shiftDateString_(dateStr, offsetDays) {
  var match = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return '';
  var date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  date.setDate(date.getDate() + Number(offsetDays || 0));
  return OppHistory_formatDate_(date);
}

function OppHistory_nullIfEmpty_(value) {
  return value === '' || value === null || value === undefined ? null : value;
}

function OppHistory_stringOrEmpty_(value) {
  return value === null || value === undefined ? '' : String(value);
}

function OppHistory_numberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  var num = typeof value === 'number' ? value : Number(value);
  return isNaN(num) ? null : num;
}

function OppHistory_numberOrZero_(value) {
  var num = OppHistory_numberOrNull_(value);
  return num === null ? 0 : num;
}

function OppHistory_toBoolean_(value) {
  if (value === true || value === false) return value;
  return /^(true|yes|1|y)$/i.test(String(value || '').trim());
}

function OppHistory_toSet_(values) {
  var set = {};
  (values || []).forEach(function(value) {
    set[value] = true;
  });
  return set;
}

function OppHistory_clone_(value) {
  return JSON.parse(JSON.stringify(value || {}));
}
