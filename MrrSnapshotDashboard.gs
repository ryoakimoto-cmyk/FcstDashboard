function MrrDashboard_getSnapshotData_(selection, options) {
  var selectedDivision = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  var opts = options || {};
  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var dateBatch = MrrDashboard_getFcstSnapshotDateBatch_(deptKeys, opts.beforeDate || '', opts.limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT);
  var fcst = MrrDashboard_readFcstSnapshots_(deptKeys, dateBatch.dateSet);
  var weeks = dateBatch.dates.slice().sort();
  var weekLabels = {};
  var data = {};
  var totalKey = selectedDivision;
  var allLabel = MrrDashboard_getTotalLabel_(selectedDivision);
  var deptLabels = deptKeys.map(function(deptKey) {
    return MrrDashboard_getDeptLabel_(deptKey);
  });

  weeks.forEach(function(dateStr) {
    var rowMap = {};
    var total = MrrDashboard_emptyMetric_();
    weekLabels[dateStr] = MrrDashboard_formatWeekLabel_(dateStr);

    deptKeys.forEach(function(deptKey) {
      var deptLabel = MrrDashboard_getDeptLabel_(deptKey);
      var metric = MrrDashboard_buildDeptMetricFromContext_(deptKey, dateStr, fcst);
      rowMap[deptLabel] = metric;
      MrrDashboard_accumulateMetric_(total, metric);
    });

    rowMap[totalKey] = total;
    data[dateStr] = rowMap;
  });

  return {
    division: selectedDivision,
    divisionLabel: MrrDashboard_getDivisionLabel_(selectedDivision),
    choices: MrrDashboard_getDivisionChoices_(),
    totalDeptKey: totalKey,
    allLabel: allLabel,
    metricDefinitions: MrrDashboard_getMetricDefinitions_(),
    weeks: weeks,
    weekLabels: weekLabels,
    depts: deptLabels,
    data: data,
    history: {
      beforeDate: opts.beforeDate || '',
      limit: opts.limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT,
      hasMore: dateBatch.hasMore,
      oldestDate: dateBatch.oldestDate || ''
    }
  };
}

function getMrrDashboardDeals(selection, dateStr, deptLabel) {
  var selectedDivision = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  var snapshotDate = MrrDashboard_normalizeSnapshotDate_(dateStr);
  if (!snapshotDate) return { deals: [] };

  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var scopedDeptKeys = MrrDashboard_filterDeptKeysByLabel_(deptKeys, deptLabel);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + 'deals:' + selectedDivision + ':' + snapshotDate + ':' + (deptLabel || 'all');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var periodByDept = MrrDashboard_getSnapshotPeriodByDept_(scopedDeptKeys, snapshotDate);
  var deals = MrrDashboard_readOppSnapshotDeals_(scopedDeptKeys, snapshotDate, periodByDept);
  var result = { deals: deals };
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
}

function MrrDashboard_getDeptKeysForDivisions_(divisionKeys) {
  var divisionSet = MrrDashboard_toSet_(divisionKeys || []);
  return Object.keys(DEPT_CONFIG).filter(function(deptKey) {
    var cfg = DEPT_CONFIG[deptKey];
    return cfg && divisionSet[cfg.division] && cfg.features && cfg.features.snapshot !== false;
  });
}

function MrrDashboard_filterDeptKeysByLabel_(deptKeys, deptLabel) {
  var label = String(deptLabel || '').trim();
  if (!label || label.indexOf('全体') !== -1 || label === MRR_DASHBOARD_ALL_DIVISION) return deptKeys.slice();
  return (deptKeys || []).filter(function(deptKey) {
    return MrrDashboard_getDeptLabel_(deptKey) === label || deptKey === label;
  });
}

function MrrDashboard_getFcstSnapshotDateBatch_(deptKeys, beforeDate, limit) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var result = { dateSet: {}, dates: [], hasMore: false, oldestDate: '' };
  if (!sheet || sheet.getLastRow() < 1) return result;

  var maxPerDivision = Number(limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT) || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT;
  var normalizedBefore = MrrDashboard_normalizeSnapshotDate_(beforeDate);
  var deptSet = MrrDashboard_toSet_(deptKeys);
  var divisionDates = {};
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

  values.forEach(function(row) {
    var date = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(date instanceof Date) || isNaN(date)) return;

    var deptKey = MrrDashboard_parseDeptFromSnapshotName_(nameRaw, deptSet);
    if (!deptKey || !DEPT_CONFIG[deptKey]) return;

    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (normalizedBefore && dateStr >= normalizedBefore) return;

    var division = DEPT_CONFIG[deptKey].division;
    if (!divisionDates[division]) divisionDates[division] = {};
    divisionDates[division][dateStr] = true;
  });

  Object.keys(divisionDates).forEach(function(division) {
    var dates = Object.keys(divisionDates[division]).sort(function(a, b) {
      return a > b ? -1 : a < b ? 1 : 0;
    });
    if (dates.length > maxPerDivision) result.hasMore = true;
    dates.slice(0, maxPerDivision).forEach(function(dateStr) {
      result.dateSet[dateStr] = true;
    });
  });

  result.dates = Object.keys(result.dateSet).sort(function(a, b) {
    return a > b ? -1 : a < b ? 1 : 0;
  });
  result.oldestDate = result.dates.length ? result.dates[result.dates.length - 1] : '';
  return result;
}

function MrrDashboard_readFcstSnapshots_(deptKeys, dateSet) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var context = { dates: {}, metricsByDeptDate: {} };
  if (!sheet || sheet.getLastRow() < 1) return context;

  var deptSet = MrrDashboard_toSet_(deptKeys);
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  var latestByDeptDate = {};
  var buckets = {};

  values.forEach(function(row) {
    var date = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(date instanceof Date) || isNaN(date)) return;

    var deptKey = MrrDashboard_parseDeptFromSnapshotName_(nameRaw, deptSet);
    if (!deptKey) return;

    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (dateSet && !dateSet[dateStr]) return;

    var timestampKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestByDeptDate[deptKey]) latestByDeptDate[deptKey] = {};
    if (!latestByDeptDate[deptKey][dateStr] || timestampKey > latestByDeptDate[deptKey][dateStr]) {
      latestByDeptDate[deptKey][dateStr] = timestampKey;
    }
  });

  values.forEach(function(row) {
    var date = row[0];
    var nameRaw = String(row[1] || '').trim();
    var period = String(row[2] || '').trim();
    if (!(date instanceof Date) || isNaN(date) || !period) return;

    var deptKey = MrrDashboard_parseDeptFromSnapshotName_(nameRaw, deptSet);
    if (!deptKey) return;

    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (dateSet && !dateSet[dateStr]) return;

    var timestampKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestByDeptDate[deptKey] || latestByDeptDate[deptKey][dateStr] !== timestampKey) return;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    MrrDashboard_addFcstSnapshotPayloadToBucket_(buckets, deptKey, dateStr, period, payload, nameRaw);
  });

  Object.keys(buckets).forEach(function(deptKey) {
    Object.keys(buckets[deptKey]).forEach(function(dateStr) {
      Object.keys(buckets[deptKey][dateStr]).forEach(function(period) {
        var payload = MrrDashboard_selectFcstSnapshotPayload_(buckets[deptKey][dateStr][period]);
        if (!payload) return;

        if (!context.metricsByDeptDate[deptKey]) context.metricsByDeptDate[deptKey] = {};
        if (!context.metricsByDeptDate[deptKey][dateStr]) {
          context.metricsByDeptDate[deptKey][dateStr] = { periods: {}, periodKeys: [] };
        }
        context.metricsByDeptDate[deptKey][dateStr].periods[period] = payload;
        if (context.metricsByDeptDate[deptKey][dateStr].periodKeys.indexOf(period) === -1) {
          context.metricsByDeptDate[deptKey][dateStr].periodKeys.push(period);
        }
        context.dates[dateStr] = true;
      });
    });
  });

  return context;
}

function MrrDashboard_getSnapshotPeriodByDept_(deptKeys, snapshotDate) {
  var dateSet = {};
  dateSet[snapshotDate] = true;
  var fcst = MrrDashboard_readFcstSnapshots_(deptKeys, dateSet);
  var result = {};
  (deptKeys || []).forEach(function(deptKey) {
    var snapshotEntry = fcst &&
      fcst.metricsByDeptDate &&
      fcst.metricsByDeptDate[deptKey] &&
      fcst.metricsByDeptDate[deptKey][snapshotDate];
    result[deptKey] = snapshotEntry
      ? MrrDashboard_pickSnapshotPeriodKeyFromKeys_(snapshotEntry.periodKeys || [], snapshotDate)
      : '';
  });
  return result;
}

function MrrDashboard_readOppSnapshotDeals_(deptKeys, snapshotDate, periodByDept) {
  var sheet = getSharedSheet(OPP_HISTORY_V2_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var deptSet = MrrDashboard_toSet_(deptKeys);
  var deals = [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, OPP_HISTORY_V2_HEADERS.length).getValues();
  values.forEach(function(row) {
    var dateStr = MrrDashboard_normalizeSnapshotDate_(row[1]);
    var deptKey = String(row[3] || '').trim();
    var status = String(row[5] || '').trim();
    if (dateStr !== snapshotDate || !deptSet[deptKey] || status === OPP_HISTORY_STATUS_REMOVED_FROM_P) return;

    var payload;
    try {
      payload = JSON.parse(String(row[6] || '{}'));
    } catch (e) {
      payload = {};
    }

    var legacyRow = OppHistory_payloadToLegacyRow_(payload, snapshotDate);
    if (!legacyRow || legacyRow.keyDeal !== true) return;

    var targetPeriod = periodByDept && periodByDept[deptKey];
    if (targetPeriod && String(legacyRow.completedMonth || '') !== targetPeriod) return;

    deals.push({
      dept: MrrDashboard_getDeptLabel_(deptKey),
      company: String(legacyRow.dealName || ''),
      mrr: Number(legacyRow.mrr) || 0,
      phase: String(legacyRow.phase || ''),
      yomi: Number(legacyRow.fcstCommit) || 0,
      oppId: String(legacyRow.oppId || ''),
      completedMonth: String(legacyRow.completedMonth || ''),
      owner: String(legacyRow.subOwner || '')
    });
  });

  deals.sort(function(left, right) {
    return Math.abs(right.mrr) - Math.abs(left.mrr);
  });
  return deals;
}

function MrrDashboard_buildDeptMetricFromContext_(deptKey, snapshotDate, fcstContext) {
  var snapshotEntry = fcstContext &&
    fcstContext.metricsByDeptDate &&
    fcstContext.metricsByDeptDate[deptKey] &&
    fcstContext.metricsByDeptDate[deptKey][snapshotDate];
  var periodKey = snapshotEntry
    ? MrrDashboard_pickSnapshotPeriodKeyFromKeys_(snapshotEntry.periodKeys || [], snapshotDate)
    : '';
  var periodMetric = periodKey && snapshotEntry ? snapshotEntry.periods[periodKey] : null;

  return {
    target: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.target),
    fcstAdjusted: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstAdjusted),
    fcstCommit: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstCommit),
    confirmed: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmed),
    actual: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmed),
    expectedMrr: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.expectedMrr),
    fcst: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstCommit),
    fcstMax: Number(periodMetric && periodMetric.fcstMax) || 0,
    received: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.received),
    debtMgmt: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.debtMgmt),
    debtMgmtLite: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.debtMgmtLite),
    expense: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.expense),
    periodKey: periodKey,
    keyDeal: '',
    keyDealsData: []
  };
}

function MrrDashboard_getMetricDefinitions_() {
  return [
    { key: 'target', label: '目標', type: 'line', color: '#ea4335', dash: true },
    { key: 'fcstAdjusted', label: 'FCST(調整後)', type: 'bar', color: '#1a73e8' },
    { key: 'fcstCommit', label: 'FCST(コミット)', type: 'bar', color: '#7c4dff' },
    { key: 'confirmed', label: '確定', type: 'bar', color: '#0b8043' },
    { key: 'expectedMrr', label: '期待MRR', type: 'bar', color: '#f29900' },
    { key: 'fcstMax', label: 'FCSTMAX', type: 'bar', color: '#5f6368' },
    { key: 'received', label: '受領', type: 'bar', color: '#188038' },
    { key: 'debtMgmt', label: '債権管理', type: 'bar', color: '#0f766e' },
    { key: 'debtMgmtLite', label: '債権管理 Lite', type: 'bar', color: '#0891b2' },
    { key: 'expense', label: '経費', type: 'bar', color: '#b45309' }
  ];
}

function MrrDashboard_getDivisionDeptKeys_(division) {
  return MrrDashboard_getDeptKeysForDivisions_([division]);
}

function MrrDashboard_getDeptLabel_(deptKey) {
  return (DEPT_CONFIG[deptKey] && DEPT_CONFIG[deptKey].label) || String(deptKey || '');
}

function MrrDashboard_parseDeptFromSnapshotName_(nameRaw, deptSet) {
  var text = String(nameRaw || '');
  var idx = text.indexOf(':');
  if (idx <= 0) return '';
  var deptKey = text.slice(0, idx);
  return deptSet[deptKey] ? deptKey : '';
}

function MrrDashboard_getFcstSnapshotBucket_(buckets, deptKey, dateStr, period) {
  if (!buckets[deptKey]) buckets[deptKey] = {};
  if (!buckets[deptKey][dateStr]) buckets[deptKey][dateStr] = {};
  if (!buckets[deptKey][dateStr][period]) {
    buckets[deptKey][dateStr][period] = {
      departmentPayload: null,
      groupPayload: MrrDashboard_emptyFcstPayload_(),
      groupCount: 0,
      individualPayload: MrrDashboard_emptyFcstPayload_(),
      individualCount: 0
    };
  }
  return buckets[deptKey][dateStr][period];
}

function MrrDashboard_addFcstSnapshotPayloadToBucket_(buckets, deptKey, dateStr, period, payload, nameRaw) {
  var bucket = MrrDashboard_getFcstSnapshotBucket_(buckets, deptKey, dateStr, period);
  var rowKind = MrrDashboard_getFcstSnapshotRowKind_(payload, nameRaw, deptKey);
  if (rowKind === 'department') {
    bucket.departmentPayload = payload;
    return;
  }
  if (rowKind === 'group') {
    bucket.groupPayload = MrrDashboard_addFcstPayload_(bucket.groupPayload, payload);
    bucket.groupCount++;
    return;
  }
  if (rowKind === 'individual') {
    bucket.individualPayload = MrrDashboard_addFcstPayload_(bucket.individualPayload, payload);
    bucket.individualCount++;
  }
}

function MrrDashboard_selectFcstSnapshotPayload_(bucket) {
  if (!bucket) return null;
  if (bucket.departmentPayload) return bucket.departmentPayload;
  if (bucket.groupCount) return bucket.groupPayload;
  if (bucket.individualCount) return bucket.individualPayload;
  return null;
}

function MrrDashboard_getFcstSnapshotRowKind_(payload, nameRaw, deptKey) {
  if (MrrDashboard_isDepartmentTotalPayload_(payload, deptKey)) return 'department';

  var meta = (payload && payload.__meta) || {};
  if (meta.isTotal) return 'group';

  var name = MrrDashboard_parseSnapshotMemberName_(nameRaw);
  if (!name || name === String(deptKey || '') || name === SHARED_ALL_GROUP_LABEL) return '';
  if (/全体/.test(name) || /グループ/.test(name)) return 'group';
  return 'individual';
}

function MrrDashboard_parseSnapshotMemberName_(nameRaw) {
  var text = String(nameRaw || '');
  var idx = text.indexOf(':');
  return idx >= 0 ? text.slice(idx + 1) : text;
}

function MrrDashboard_isDepartmentTotalPayload_(payload, deptKey) {
  var meta = (payload && payload.__meta) || {};
  if (!meta.isTotal) return false;
  if (typeof SHARED_TOTAL_KIND !== 'undefined' && meta.totalKind) {
    return meta.totalKind === SHARED_TOTAL_KIND.DEPARTMENT && String(meta.dept || '') === String(deptKey);
  }
  return SharedAppState_isDepartmentTotal_({
    isTotal: !!meta.isTotal,
    totalKind: meta.totalKind || '',
    group: meta.group || '',
    groupCode: meta.groupCode || '',
    dept: meta.dept || deptKey
  }, deptKey);
}

function MrrDashboard_emptyFcstPayload_() {
  return {
    target: MrrDashboard_emptyBreakdown_(),
    fcstAdjusted: MrrDashboard_emptyBreakdown_(),
    fcstCommit: MrrDashboard_emptyBreakdown_(),
    confirmed: MrrDashboard_emptyBreakdown_(),
    expectedMrr: MrrDashboard_emptyBreakdown_(),
    fcstMax: 0,
    received: MrrDashboard_emptyBreakdown_(),
    debtMgmt: MrrDashboard_emptyBreakdown_(),
    debtMgmtLite: MrrDashboard_emptyBreakdown_(),
    expense: MrrDashboard_emptyBreakdown_()
  };
}

function MrrDashboard_emptyBreakdown_() {
  return { net: 0, newExp: 0, churn: 0 };
}

function MrrDashboard_addFcstPayload_(sum, payload) {
  sum = sum || MrrDashboard_emptyFcstPayload_();
  [
    'target',
    'fcstAdjusted',
    'fcstCommit',
    'confirmed',
    'expectedMrr',
    'received',
    'debtMgmt',
    'debtMgmtLite',
    'expense'
  ].forEach(function(key) {
    MrrDashboard_addBreakdownTo_(sum[key], payload && payload[key]);
  });
  sum.fcstMax += Number(payload && payload.fcstMax) || 0;
  return sum;
}

function MrrDashboard_addBreakdownTo_(target, value) {
  if (!target) return;
  if (typeof value === 'number') {
    target.net += Number(value) || 0;
    return;
  }
  target.net += Number(value && value.net) || 0;
  target.newExp += Number(value && value.newExp) || 0;
  target.churn += Number(value && value.churn) || 0;
}

function MrrDashboard_pickSnapshotPeriodKeyFromKeys_(periodKeys, snapshotDate) {
  var monthKeys = (periodKeys || []).filter(function(periodKey) {
    return /^\d{4}-\d{2}$/.test(String(periodKey || ''));
  }).sort();
  if (!monthKeys.length) return '';

  var snapshotMonth = String(snapshotDate || '').slice(0, 7);
  if (monthKeys.indexOf(snapshotMonth) !== -1) return snapshotMonth;

  var futureMonth = monthKeys.find(function(monthKey) {
    return monthKey >= snapshotMonth;
  });
  return futureMonth || monthKeys[monthKeys.length - 1];
}

function MrrDashboard_formatWeekLabel_(snapshotDate) {
  var match = String(snapshotDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(snapshotDate || '');
  return Number(match[2]) + '/' + Number(match[3]);
}

function MrrDashboard_emptyMetric_() {
  var metric = {
    target: 0,
    actual: 0,
    expectedMrr: 0,
    fcst: 0,
    keyDeal: '',
    keyDealsData: []
  };
  MrrDashboard_getMetricDefinitions_().forEach(function(def) {
    if (!metric.hasOwnProperty(def.key)) metric[def.key] = 0;
  });
  return metric;
}

function MrrDashboard_accumulateMetric_(target, metric) {
  if (!target || !metric) return;
  Object.keys(metric).forEach(function(key) {
    if (key === 'keyDeal' || key === 'keyDealsData' || key === 'periodKey') return;
    if (typeof metric[key] !== 'number') return;
    target[key] = (Number(target[key]) || 0) + (Number(metric[key]) || 0);
  });
}

function MrrDashboard_getBreakdownNet_(value) {
  if (typeof value === 'number') return Number(value) || 0;
  return Number(value && value.net) || 0;
}

function MrrDashboard_toSet_(values) {
  var set = {};
  (values || []).forEach(function(value) {
    set[value] = true;
  });
  return set;
}

function MrrDashboard_normalizeSnapshotDate_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : '';
}
