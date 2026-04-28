function MrrDashboard_getBoData_() {
  var deptKeys = MrrDashboard_getDivisionDeptKeys_('BO');
  var fcst = MrrDashboard_readBoFcstSnapshots_(deptKeys);
  var opp = MrrDashboard_readBoOppSnapshots_(deptKeys);
  var weekLabels = {};
  var data = {};
  var datesSeen = {};

  Object.keys(fcst.dates || {}).forEach(function(dateStr) { datesSeen[dateStr] = true; });
  Object.keys(opp.dates || {}).forEach(function(dateStr) { datesSeen[dateStr] = true; });

  var weeks = Object.keys(datesSeen).sort();
  var deptLabels = deptKeys.map(function(deptKey) {
    return MrrDashboard_getDeptLabel_(deptKey);
  });

  weeks.forEach(function(dateStr) {
    var rowMap = {};
    var total = MrrDashboard_emptyMetric_();
    weekLabels[dateStr] = MrrDashboard_formatWeekLabel_(dateStr);

    deptKeys.forEach(function(deptKey) {
      var deptLabel = MrrDashboard_getDeptLabel_(deptKey);
      var metric = MrrDashboard_buildBoDeptMetricFromContexts_(deptKey, dateStr, fcst, opp);
      rowMap[deptLabel] = metric;
      MrrDashboard_accumulateMetric_(total, metric);
    });

    total.keyDeal = MrrDashboard_formatLegacyDeals_(total.keyDealsData);
    data[dateStr] = rowMap;
    data[dateStr].BO = total;
  });

  return {
    division: 'BO',
    totalDeptKey: 'BO',
    allLabel: 'BO全体',
    metricDefinitions: MrrDashboard_getBoMetricDefinitions_(),
    weeks: weeks,
    weekLabels: weekLabels,
    depts: deptLabels,
    data: data
  };
}

function MrrDashboard_readBoFcstSnapshots_(deptKeys) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var context = { dates: {}, metricsByDeptDate: {} };
  if (!sheet || sheet.getLastRow() < 1) return context;

  var deptSet = MrrDashboard_toSet_(deptKeys);
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  var latestByDeptDate = {};

  values.forEach(function(row) {
    var date = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(date instanceof Date) || isNaN(date)) return;

    var deptKey = MrrDashboard_parseDeptFromSnapshotName_(nameRaw, deptSet);
    if (!deptKey) return;

    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
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
    var timestampKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestByDeptDate[deptKey] || latestByDeptDate[deptKey][dateStr] !== timestampKey) return;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    if (!MrrDashboard_isDepartmentTotalPayload_(payload, deptKey)) return;

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

  return context;
}

function MrrDashboard_readBoOppSnapshots_(deptKeys) {
  var sheet = getSharedSheet(OPP_HISTORY_V2_SHEET_NAME);
  var context = { dates: {}, keyDealsByDeptDate: {} };
  if (!sheet || sheet.getLastRow() < 2) return context;

  var deptSet = MrrDashboard_toSet_(deptKeys);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, OPP_HISTORY_V2_HEADERS.length).getValues();
  values.forEach(function(row) {
    var dateStr = MrrDashboard_normalizeSnapshotDate_(row[1]);
    var deptKey = String(row[3] || '').trim();
    var status = String(row[5] || '').trim();
    if (!dateStr || !deptSet[deptKey] || status === OPP_HISTORY_STATUS_REMOVED_FROM_P) return;

    var payload;
    try {
      payload = JSON.parse(String(row[6] || '{}'));
    } catch (e) {
      payload = {};
    }

    var legacyRow = OppHistory_payloadToLegacyRow_(payload, dateStr);
    if (!legacyRow || !legacyRow.keyDeal) return;

    if (!context.keyDealsByDeptDate[deptKey]) context.keyDealsByDeptDate[deptKey] = {};
    if (!context.keyDealsByDeptDate[deptKey][dateStr]) context.keyDealsByDeptDate[deptKey][dateStr] = [];
    context.keyDealsByDeptDate[deptKey][dateStr].push({
      dept: MrrDashboard_getDeptLabel_(deptKey),
      company: String(legacyRow.dealName || ''),
      mrr: Number(legacyRow.mrr) || 0,
      phase: String(legacyRow.phase || ''),
      yomi: Number(legacyRow.fcstCommit) || 0,
      oppId: String(legacyRow.oppId || ''),
      completedMonth: String(legacyRow.completedMonth || ''),
      owner: String(legacyRow.subOwner || '')
    });
    context.dates[dateStr] = true;
  });

  Object.keys(context.keyDealsByDeptDate).forEach(function(deptKey) {
    Object.keys(context.keyDealsByDeptDate[deptKey]).forEach(function(dateStr) {
      context.keyDealsByDeptDate[deptKey][dateStr].sort(function(left, right) {
        return Math.abs(right.mrr) - Math.abs(left.mrr);
      });
    });
  });

  return context;
}

function MrrDashboard_buildBoDeptMetricFromContexts_(deptKey, snapshotDate, fcstContext, oppContext) {
  var snapshotEntry = fcstContext &&
    fcstContext.metricsByDeptDate &&
    fcstContext.metricsByDeptDate[deptKey] &&
    fcstContext.metricsByDeptDate[deptKey][snapshotDate];
  var periodKey = snapshotEntry
    ? MrrDashboard_pickSnapshotPeriodKeyFromKeys_(snapshotEntry.periodKeys || [], snapshotDate)
    : '';
  var periodMetric = periodKey && snapshotEntry ? snapshotEntry.periods[periodKey] : null;
  var keyDealsData = (oppContext &&
    oppContext.keyDealsByDeptDate &&
    oppContext.keyDealsByDeptDate[deptKey] &&
    oppContext.keyDealsByDeptDate[deptKey][snapshotDate]) || [];

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
    keyDeal: MrrDashboard_formatLegacyDeals_(keyDealsData),
    keyDealsData: keyDealsData
  };
}

function MrrDashboard_getBoMetricDefinitions_() {
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
  return Object.keys(DEPT_CONFIG).filter(function(deptKey) {
    return DEPT_CONFIG[deptKey] && DEPT_CONFIG[deptKey].division === division;
  });
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

function MrrDashboard_isDepartmentTotalPayload_(payload, deptKey) {
  var meta = (payload && payload.__meta) || {};
  if (!meta.isTotal) return false;
  if (meta.totalKind) return meta.totalKind === SHARED_TOTAL_KIND.DEPARTMENT && String(meta.dept || '') === String(deptKey);
  return SharedAppState_isDepartmentTotal_({
    isTotal: !!meta.isTotal,
    totalKind: meta.totalKind || '',
    group: meta.group || '',
    groupCode: meta.groupCode || '',
    dept: meta.dept || deptKey
  }, deptKey);
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
  MrrDashboard_getBoMetricDefinitions_().forEach(function(def) {
    if (!metric.hasOwnProperty(def.key)) metric[def.key] = 0;
  });
  return metric;
}

function MrrDashboard_accumulateMetric_(target, metric) {
  if (!target || !metric) return;
  Object.keys(metric).forEach(function(key) {
    if (key === 'keyDeal' || key === 'keyDealsData') return;
    if (typeof metric[key] !== 'number') return;
    target[key] = (Number(target[key]) || 0) + (Number(metric[key]) || 0);
  });
  target.keyDealsData = target.keyDealsData.concat(metric.keyDealsData || []);
}

function MrrDashboard_getBreakdownNet_(value) {
  if (typeof value === 'number') return Number(value) || 0;
  return Number(value && value.net) || 0;
}

function MrrDashboard_formatLegacyDeals_(deals) {
  return (deals || []).map(function(deal) {
    var parts = [
      String(deal.company || '').trim() + ' ' + MrrDashboard_toMan_(deal.mrr),
      String(deal.phase || '').trim(),
      MrrDashboard_toMan_(deal.yomi)
    ];
    return parts.join(' / ').trim();
  }).join('\n');
}

function MrrDashboard_toMan_(value) {
  return String(Math.round((Number(value) || 0) / 10000));
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
  return match ? match[1] : text;
}
