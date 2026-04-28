function MrrDashboard_getSnapshotData_(selection, options) {
  var selectedDivision = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  var opts = options || {};
  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var dateBatch = { dateSet: {}, dates: [], hasMore: false, oldestDate: '' };
  var fcst = { dates: {}, metricsByDeptDate: {} };
  if (!opts.currentOnly) {
    dateBatch = MrrDashboard_getFcstSnapshotDateBatch_(deptKeys, opts.beforeDate || '', opts.limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT);
    fcst = MrrDashboard_readFcstSnapshots_(deptKeys, dateBatch.dateSet);
  }
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

  if (opts.currentOnly || opts.includeCurrent) {
    MrrDashboard_addCurrentData_(deptKeys, weeks, weekLabels, data, totalKey);
  }

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
  var requestedDate = String(dateStr || '').trim();
  var isLive = requestedDate === MRR_DASHBOARD_LIVE_KEY;
  var snapshotDate = isLive ? MRR_DASHBOARD_LIVE_KEY : MrrDashboard_normalizeSnapshotDate_(requestedDate);
  if (!snapshotDate) return { deals: [] };

  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var scopedDeptKeys = MrrDashboard_filterDeptKeysByLabel_(deptKeys, deptLabel);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + 'deals:' + selectedDivision + ':' + snapshotDate + ':' + (deptLabel || 'all');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var periodByDept = isLive
    ? MrrDashboard_getCurrentPeriodByDept_(scopedDeptKeys)
    : MrrDashboard_getSnapshotPeriodByDept_(scopedDeptKeys, snapshotDate);
  var deals = isLive
    ? MrrDashboard_readCurrentOppDeals_(scopedDeptKeys, periodByDept)
    : MrrDashboard_readOppSnapshotDeals_(scopedDeptKeys, snapshotDate, periodByDept);
  var result = { deals: deals };
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
}

function MrrDashboard_addCurrentData_(deptKeys, weeks, weekLabels, data, totalKey) {
  var liveRows = {};
  var total = MrrDashboard_emptyMetric_();
  var hasLive = false;

  (deptKeys || []).forEach(function(deptKey) {
    var metric = MrrDashboard_buildCurrentDeptMetric_(deptKey);
    if (!metric) return;
    liveRows[MrrDashboard_getDeptLabel_(deptKey)] = metric;
    MrrDashboard_accumulateMetric_(total, metric);
    hasLive = true;
  });

  if (!hasLive) return;
  liveRows[totalKey] = total;
  if (weeks.indexOf(MRR_DASHBOARD_LIVE_KEY) === -1) weeks.push(MRR_DASHBOARD_LIVE_KEY);
  weekLabels[MRR_DASHBOARD_LIVE_KEY] = '現在';
  data[MRR_DASHBOARD_LIVE_KEY] = liveRows;
}

function MrrDashboard_buildCurrentDeptMetric_(deptKey) {
  var live = AppDataCache_getInitData(deptKey);
  if (!live || live.error) {
    throw new Error('MRR current FCST source failed: ' + deptKey + ' / ' + (live && live.error ? live.error : 'empty payload'));
  }

  var member = MrrDashboard_findCurrentDepartmentTotalMember_(deptKey, live.members || []);
  if (!member) return null;

  var periodKey = MrrDashboard_pickCurrentPeriodKey_(live);
  if (!periodKey) return null;

  var periodMetric = MrrDashboard_buildCurrentPeriodMetric_(live, member, periodKey);
  return MrrDashboard_buildMetricFromPeriodMetric_(periodMetric, periodKey);
}

function MrrDashboard_findCurrentDepartmentTotalMember_(deptKey, members) {
  for (var i = 0; i < (members || []).length; i++) {
    var member = members[i] || {};
    if (SharedAppState_isDepartmentTotal_(member, deptKey)) return member;
  }
  return null;
}

function MrrDashboard_pickCurrentPeriodKey_(live) {
  var monthKeys = MrrDashboard_getCurrentMonthKeys_(live && live.periodOptions);
  if (!monthKeys.length) return '';

  var currentMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  if (monthKeys.indexOf(currentMonth) !== -1) return currentMonth;

  var futureMonth = monthKeys.find(function(monthKey) {
    return monthKey >= currentMonth;
  });
  return futureMonth || monthKeys[monthKeys.length - 1];
}

function MrrDashboard_getCurrentMonthKeys_(periodOptions) {
  var seen = {};
  var keys = [];
  (periodOptions || []).forEach(function(option) {
    (option && option.months || []).forEach(function(monthKey) {
      var key = String(monthKey || '').trim();
      if (!/^\d{4}-\d{2}$/.test(key) || seen[key]) return;
      seen[key] = true;
      keys.push(key);
    });
  });
  return keys.sort(function(a, b) { return a < b ? -1 : a > b ? 1 : 0; });
}

function MrrDashboard_buildCurrentPeriodMetric_(live, member, periodKey) {
  var metric = {};
  Object.keys((member && member[periodKey]) || {}).forEach(function(key) {
    metric[key] = member[periodKey][key];
  });
  var adjustedKey = String(member && member.name || '') + '|' + String(periodKey || '');
  if (live && live.fcstAdjusted && live.fcstAdjusted[adjustedKey]) {
    metric.fcstAdjusted = live.fcstAdjusted[adjustedKey];
  }
  return metric;
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
  var result = { dateSet: {}, dates: [], hasMore: false, oldestDate: '' };
  var values = FcstSnapshot_getAllValues_(2);
  if (!values.length) return result;

  var maxPerDivision = Number(limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT) || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT;
  var normalizedBefore = MrrDashboard_normalizeSnapshotDate_(beforeDate);
  var deptSet = MrrDashboard_toSet_(deptKeys);
  var divisionDates = {};

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
  var context = { dates: {}, metricsByDeptDate: {} };
  var values = FcstSnapshot_getAllValues_(4);
  if (!values.length) return context;

  var deptSet = MrrDashboard_toSet_(deptKeys);
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

  return MrrDashboard_buildMetricFromPeriodMetric_(periodMetric, periodKey);
}

function MrrDashboard_buildMetricFromPeriodMetric_(periodMetric, periodKey) {
  return {
    target: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.target),
    fcstAdjusted: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstAdjusted),
    fcstCommit: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstCommit),
    confirmed: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmed),
    actual: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmed),
    expectedMrr: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.expectedMrr),
    fcst: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstCommit),
    received: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.received),
    debtMgmt: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.debtMgmt),
    debtMgmtLite: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.debtMgmtLite),
    expense: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.expense),
    confirmedReceived: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmedReceived),
    confirmedDebtMgmt: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmedDebtMgmt),
    confirmedDebtMgmtLite: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmedDebtMgmtLite),
    confirmedExpense: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmedExpense),
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
    { key: 'received', label: 'FCST 受領', type: 'bar', color: '#188038' },
    { key: 'confirmedReceived', label: '確定 受領', type: 'bar', color: '#38bdf8' },
    { key: 'debtMgmt', label: 'FCST 債権管理', type: 'bar', color: '#0f766e' },
    { key: 'confirmedDebtMgmt', label: '確定 債権管理', type: 'bar', color: '#2dd4bf' },
    { key: 'debtMgmtLite', label: 'FCST 債権管理 Lite', type: 'bar', color: '#0891b2' },
    { key: 'confirmedDebtMgmtLite', label: '確定 債権管理 Lite', type: 'bar', color: '#67e8f9' },
    { key: 'expense', label: 'FCST 経費', type: 'bar', color: '#b45309' },
    { key: 'confirmedExpense', label: '確定 経費', type: 'bar', color: '#fb923c' }
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
    received: MrrDashboard_emptyBreakdown_(),
    debtMgmt: MrrDashboard_emptyBreakdown_(),
    debtMgmtLite: MrrDashboard_emptyBreakdown_(),
    expense: MrrDashboard_emptyBreakdown_(),
    confirmedReceived: MrrDashboard_emptyBreakdown_(),
    confirmedDebtMgmt: MrrDashboard_emptyBreakdown_(),
    confirmedDebtMgmtLite: MrrDashboard_emptyBreakdown_(),
    confirmedExpense: MrrDashboard_emptyBreakdown_()
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
    'expense',
    'confirmedReceived',
    'confirmedDebtMgmt',
    'confirmedDebtMgmtLite',
    'confirmedExpense'
  ].forEach(function(key) {
    MrrDashboard_addBreakdownTo_(sum[key], payload && payload[key]);
  });
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
  if (String(snapshotDate || '') === MRR_DASHBOARD_LIVE_KEY) return '現在';
  var match = String(snapshotDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(snapshotDate || '');
  return Number(match[2]) + '/' + Number(match[3]);
}

function MrrDashboard_getCurrentPeriodByDept_(deptKeys) {
  var result = {};
  (deptKeys || []).forEach(function(deptKey) {
    var live = AppDataCache_getInitData(deptKey);
    if (!live || live.error) {
      throw new Error('MRR current FCST source failed: ' + deptKey + ' / ' + (live && live.error ? live.error : 'empty payload'));
    }
    result[deptKey] = MrrDashboard_pickCurrentPeriodKey_(live);
  });
  return result;
}

function MrrDashboard_readCurrentOppDeals_(deptKeys, periodByDept) {
  var deals = [];
  (deptKeys || []).forEach(function(deptKey) {
    var result = AppDataCache_getOpportunities(deptKey);
    if (!result || result.error) {
      throw new Error('MRR current Opp source failed: ' + deptKey + ' / ' + (result && result.error ? result.error : 'empty payload'));
    }

    var targetPeriod = periodByDept && periodByDept[deptKey];
    (result.rows || []).forEach(function(row) {
      if (!row || row.keyDeal !== true) return;
      if (targetPeriod && String(row.completedMonth || '') !== targetPeriod) return;

      deals.push({
        dept: MrrDashboard_getDeptLabel_(deptKey),
        company: String(row.dealName || ''),
        mrr: Number(row.mrr) || 0,
        phase: String(row.phase || ''),
        yomi: Number(row.fcstCommit) || 0,
        oppId: String(row.oppId || ''),
        completedMonth: String(row.completedMonth || ''),
        owner: String(row.subOwner || '')
      });
    });
  });

  deals.sort(function(left, right) {
    return Math.abs(right.mrr) - Math.abs(left.mrr);
  });
  return deals;
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
