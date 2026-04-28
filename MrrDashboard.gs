var MRR_DASHBOARD_CACHE_DEPT = '__mrr_dashboard__';
var MRR_DASHBOARD_CACHE_KEY = 'snapshotLiveData:v1';
var MRR_DASHBOARD_OLD_CACHE_KEYS = ['snapshotData:v1', 'snapshotData:v2', 'snapshotData:v4', 'snapshotData:v5', 'snapshotData:v6'];
var MRR_DASHBOARD_TOTAL_KEY = 'total';
var MRR_DASHBOARD_LIVE_KEY = 'live';
var MRR_DASHBOARD_DIVISION_ORDER = ['SS', 'BO', 'CO'];
var MRR_DASHBOARD_NUMERIC_METRIC_KEYS = {
  fcstMin: true,
  fcstMax: true
};

function mrrDashboard_doGet_() {
  return HtmlService.createHtmlOutputFromFile('mrr-index')
    .setTitle('MRR進捗ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMrrDashboardData() {
  var cached = CacheLayer_read(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, { skipSharedSheet: true });
  if (cached) return cached;

  var data = MrrDashboard_buildFromSnapshots_();
  if (data && data.divisions && data.divisions.length) {
    CacheLayer_write(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, data, { persistToSheet: false });
  }
  return data;
}

function MrrDashboard_invalidateCache_() {
  if (typeof CacheLayer_remove !== 'function') return;
  [MRR_DASHBOARD_CACHE_KEY].concat(MRR_DASHBOARD_OLD_CACHE_KEYS || []).forEach(function(key) {
    CacheLayer_remove(MRR_DASHBOARD_CACHE_DEPT, key);
  });
}

function MrrDashboard_buildFromSnapshots_() {
  var catalog = MrrDashboard_buildDeptCatalog_();
  var divisions = {};
  var fcstRows = FcstSnapshot_getAllValues_(4);
  var oppRows = OppListSnapshot_getAllValues_(5);
  var dealsByDivisionDate = MrrDashboard_buildDealsByDivisionDate_(oppRows, catalog);

  fcstRows.slice().sort(function(a, b) {
    var ad = a && a[0] instanceof Date ? a[0].getTime() : 0;
    var bd = b && b[0] instanceof Date ? b[0].getTime() : 0;
    return ad - bd;
  }).forEach(function(row) {
    var snapshotAt = row[0];
    if (!(snapshotAt instanceof Date) || isNaN(snapshotAt)) return;

    var nameRaw = String(row[1] || '').trim();
    var rowInfo = FcstSnapshot_parseRowName_(nameRaw);
    var rowDeptKey = String(rowInfo.deptKey || '').trim();
    var periodKey = String(row[2] || '').trim();
    if (!FcstPeriods_parseMonthKey_(periodKey)) return;
    var payload = MrrDashboard_parseJson_(row[3]);
    var meta = FcstSnapshot_normalizeMeta_(payload, rowDeptKey, rowInfo.name);
    var deptKey = rowDeptKey;
    var deptMeta = catalog.byDeptKey[deptKey];
    if (!deptMeta || !periodKey) return;
    if (meta.totalKind !== SHARED_TOTAL_KIND.DEPARTMENT) return;

    var divisionKey = deptMeta.divisionKey;
    var division = MrrDashboard_ensureDivision_(divisions, divisionKey, deptMeta.divisionLabel);
    var dateKey = Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'yyyy-MM-dd');

    MrrDashboard_addWeek_(division, {
      key: dateKey,
      date: dateKey,
      label: Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'M/d'),
      timestamp: dateKey
    });
    MrrDashboard_addPeriod_(division, periodKey);
    MrrDashboard_addDept_(division, deptMeta);

    if (!division.data[periodKey]) division.data[periodKey] = {};
    if (!division.data[periodKey][dateKey]) division.data[periodKey][dateKey] = {};
    division.data[periodKey][dateKey][deptKey] = MrrDashboard_extractMetrics_(payload);
  });

  MrrDashboard_addLiveData_(divisions, catalog, dealsByDivisionDate);

  Object.keys(divisions).forEach(function(divisionKey) {
    var division = divisions[divisionKey];
    MrrDashboard_addDerivedQuarterPeriods_(division);
    MrrDashboard_sortDivision_(division);
    MrrDashboard_fillTotalsAndDeals_(division, dealsByDivisionDate[divisionKey] || {});
  });

  var orderedDivisions = MRR_DASHBOARD_DIVISION_ORDER.filter(function(key) {
    return !!divisions[key];
  }).map(function(key) {
    return divisions[key];
  });

  return {
    source: 'snapshot+current',
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    defaultDivision: divisions.BO ? 'BO' : (orderedDivisions[0] ? orderedDivisions[0].key : ''),
    totalKey: MRR_DASHBOARD_TOTAL_KEY,
    metrics: MrrDashboard_metricDefs_(),
    divisions: orderedDivisions
  };
}

function MrrDashboard_addLiveData_(divisions, catalog, dealsByDivisionDate) {
  Object.keys(catalog.byDeptKey || {}).forEach(function(deptKey) {
    var deptMeta = catalog.byDeptKey[deptKey];
    var live = AppDataCache_getInitData(deptKey);
    if (!live || live.error) {
      throw new Error('MRR live FCST source failed: ' + deptKey + ' / ' + (live && live.error ? live.error : 'empty payload'));
    }

    var member = MrrDashboard_findDepartmentTotalMember_(live.members || []);
    if (member) {
      var division = MrrDashboard_ensureDivision_(divisions, deptMeta.divisionKey, deptMeta.divisionLabel);
      MrrDashboard_addWeek_(division, {
        key: MRR_DASHBOARD_LIVE_KEY,
        date: MRR_DASHBOARD_LIVE_KEY,
        label: '現在',
        timestamp: MRR_DASHBOARD_LIVE_KEY,
        isLive: true
      });
      MrrDashboard_addDept_(division, deptMeta);

      MrrDashboard_liveMonthKeys_(live.periodOptions || []).forEach(function(monthKey) {
        MrrDashboard_addPeriod_(division, monthKey);
        if (!division.data[monthKey]) division.data[monthKey] = {};
        if (!division.data[monthKey][MRR_DASHBOARD_LIVE_KEY]) division.data[monthKey][MRR_DASHBOARD_LIVE_KEY] = {};
        division.data[monthKey][MRR_DASHBOARD_LIVE_KEY][deptKey] = MrrDashboard_extractMetrics_(
          MrrDashboard_buildLiveMetric_(live, member, monthKey)
        );
      });
    }

    MrrDashboard_addLiveDealsForDept_(dealsByDivisionDate, deptMeta);
  });
}

function MrrDashboard_findDepartmentTotalMember_(members) {
  for (var i = 0; i < (members || []).length; i++) {
    var member = members[i] || {};
    if (member.totalKind === SHARED_TOTAL_KIND.DEPARTMENT) return member;
  }
  return null;
}

function MrrDashboard_liveMonthKeys_(periodOptions) {
  var seen = {};
  var result = [];
  (periodOptions || []).forEach(function(option) {
    (option && option.months || []).forEach(function(monthKey) {
      var key = String(monthKey || '').trim();
      if (!FcstPeriods_parseMonthKey_(key) || seen[key]) return;
      seen[key] = true;
      result.push(key);
    });
  });
  return result;
}

function MrrDashboard_addLiveDealsForDept_(dealsByDivisionDate, deptMeta) {
  var oppDeptKeys = deptMeta.oppDeptKeys || [];
  if (!oppDeptKeys.length) {
    throw new Error('MRR live Opp source mapping missing: ' + deptMeta.key);
  }

  var divisionKey = deptMeta.divisionKey;
  if (!dealsByDivisionDate[divisionKey]) dealsByDivisionDate[divisionKey] = {};
  if (!dealsByDivisionDate[divisionKey][MRR_DASHBOARD_LIVE_KEY]) dealsByDivisionDate[divisionKey][MRR_DASHBOARD_LIVE_KEY] = {};
  if (!dealsByDivisionDate[divisionKey][MRR_DASHBOARD_LIVE_KEY][deptMeta.key]) {
    dealsByDivisionDate[divisionKey][MRR_DASHBOARD_LIVE_KEY][deptMeta.key] = [];
  }

  var seen = {};
  oppDeptKeys.forEach(function(oppDeptKey) {
    var result = AppDataCache_getOpportunities(oppDeptKey);
    if (!result || result.error) {
      throw new Error('MRR live Opp source failed: ' + oppDeptKey + ' / ' + (result && result.error ? result.error : 'empty payload'));
    }

    (result.rows || []).forEach(function(row) {
      if (!row || row.keyDeal !== true) return;
      var oppId = String(row.oppId || '').trim();
      if (oppId && seen[oppId]) return;
      if (oppId) seen[oppId] = true;
      dealsByDivisionDate[divisionKey][MRR_DASHBOARD_LIVE_KEY][deptMeta.key].push({
        oppId: oppId,
        deptKey: deptMeta.key,
        deptLabel: deptMeta.label,
        company: String(row.dealName || '').trim(),
        mrr: MrrDashboard_num_(row.mrr),
        phase: String(row.phase || '').trim(),
        forecast: String(row.forecast || '').trim(),
        fcstCommit: MrrDashboard_num_(row.fcstCommit),
        fcstMin: MrrDashboard_num_(row.fcstMin),
        fcstMax: MrrDashboard_num_(row.fcstMax),
        completedMonth: String(row.completedMonth || '').trim(),
        comment: String(row.fcstComment || '').trim()
      });
    });
  });
}

function MrrDashboard_buildLiveMetric_(live, member, periodKey) {
  var metric = {};
  var source = member && member[periodKey] ? member[periodKey] : {};
  Object.keys(source).forEach(function(key) {
    metric[key] = source[key];
  });
  metric.fcstAdjusted = MrrDashboard_getLiveAdjustedMetric_(live, member, periodKey);
  return metric;
}

function MrrDashboard_getLiveAdjustedMetric_(live, member, periodKey) {
  var name = String(member && member.name || '').trim();
  var key = name + '|' + String(periodKey || '').trim();
  var adjusted = live && live.fcstAdjusted && live.fcstAdjusted[key];
  if (adjusted && typeof adjusted.net === 'number') return adjusted;
  return { net: 0, newExp: 0, churn: 0 };
}

function MrrDashboard_addDerivedQuarterPeriods_(division) {
  var monthsByQuarter = {};
  Object.keys(division.data || {}).forEach(function(periodKey) {
    if (!FcstPeriods_parseMonthKey_(periodKey)) return;
    var quarterKey = FcstPeriods_getQuarterKeyFromMonthKey_(periodKey);
    if (!quarterKey) return;
    if (!monthsByQuarter[quarterKey]) monthsByQuarter[quarterKey] = {};
    monthsByQuarter[quarterKey][periodKey] = true;
  });

  Object.keys(monthsByQuarter).forEach(function(quarterKey) {
    MrrDashboard_addPeriod_(division, quarterKey);
    if (!division.data[quarterKey]) division.data[quarterKey] = {};

    division.weeks.forEach(function(week) {
      if (!division.data[quarterKey][week.key]) division.data[quarterKey][week.key] = {};
      division.depts.forEach(function(dept) {
        var rows = [];
        Object.keys(monthsByQuarter[quarterKey]).forEach(function(monthKey) {
          var monthRows = division.data[monthKey] && division.data[monthKey][week.key];
          if (monthRows && monthRows[dept.key]) rows.push(monthRows[dept.key]);
        });
        division.data[quarterKey][week.key][dept.key] = MrrDashboard_sumExtractedMetrics_(rows);
      });
    });
  });
}

function MrrDashboard_sumExtractedMetrics_(rows) {
  var metrics = MrrDashboard_emptyMetrics_();
  (rows || []).forEach(function(row) {
    MrrDashboard_metricDefs_().forEach(function(def) {
      metrics[def.key] += MrrDashboard_num_(row && row[def.key]);
    });
  });
  return metrics;
}

function MrrDashboard_buildDeptCatalog_() {
  var byDeptKey = {};
  var byAnyKey = {};
  var configMap = getDeptConfigMap_();

  Object.keys(configMap || {}).forEach(function(deptKey) {
    var cfg = configMap[deptKey] || {};
    var divisionKey = MrrDashboard_divisionKey_(cfg);
    if (MRR_DASHBOARD_DIVISION_ORDER.indexOf(divisionKey) === -1) return;
    var meta = {
      key: deptKey,
      label: String(cfg.label || cfg.departmentName || deptKey).trim(),
      divisionKey: divisionKey,
      divisionLabel: divisionKey,
      oppDeptKeys: []
    };
    byDeptKey[deptKey] = meta;
    byAnyKey[MrrDashboard_normKey_(deptKey)] = meta;
    byAnyKey[MrrDashboard_normKey_(meta.label)] = meta;
  });

  OrgMasterReader_getRows().forEach(function(row) {
    if (!row) return;
    var deptKey = String(row.departmentCode || '').trim();
    var meta = byDeptKey[deptKey];
    if (!meta) return;
    byAnyKey[MrrDashboard_normKey_(row.groupName)] = meta;
    byAnyKey[MrrDashboard_normKey_(row.departmentName)] = meta;
    if (row.groupName && meta.oppDeptKeys.indexOf(row.groupName) === -1) {
      meta.oppDeptKeys.push(row.groupName);
    }
  });

  return { byDeptKey: byDeptKey, byAnyKey: byAnyKey };
}

function MrrDashboard_divisionKey_(cfg) {
  var division = String(cfg && cfg.division || '').trim().toUpperCase();
  if (MRR_DASHBOARD_DIVISION_ORDER.indexOf(division) !== -1) return division;

  var sfSheetKey = String(cfg && cfg.sfSheetKey || '').trim().toUpperCase();
  if (sfSheetKey === 'SS' || sfSheetKey === 'SSCS') return 'SS';
  if (sfSheetKey === 'BO') return 'BO';
  if (sfSheetKey === 'CO') return 'CO';
  return '';
}

function MrrDashboard_buildDealsByDivisionDate_(rows, catalog) {
  var result = {};
  (rows || []).forEach(function(row) {
    var dateKey = String(row[1] || '').trim();
    var oppDept = String(row[3] || '').trim();
    var deptMeta = catalog.byAnyKey[MrrDashboard_normKey_(oppDept)];
    if (!dateKey || !deptMeta) return;

    var payload = MrrDashboard_parseJson_(row[4]);
    if (!payload || payload.keyDeal !== true) return;

    var divisionKey = deptMeta.divisionKey;
    if (!result[divisionKey]) result[divisionKey] = {};
    if (!result[divisionKey][dateKey]) result[divisionKey][dateKey] = {};
    if (!result[divisionKey][dateKey][deptMeta.key]) result[divisionKey][dateKey][deptMeta.key] = [];

    result[divisionKey][dateKey][deptMeta.key].push({
      oppId: String(payload.oppId || row[2] || '').trim(),
      deptKey: deptMeta.key,
      deptLabel: deptMeta.label,
      company: String(payload.dealName || '').trim(),
      mrr: MrrDashboard_num_(payload.mrr),
      phase: String(payload.phase || '').trim(),
      forecast: String(payload.forecast || '').trim(),
      fcstCommit: MrrDashboard_num_(payload.fcstCommit),
      fcstMin: MrrDashboard_num_(payload.fcstMin),
      fcstMax: MrrDashboard_num_(payload.fcstMax),
      completedMonth: String(payload.completedMonth || '').trim(),
      comment: String(payload.fcstComment || '').trim()
    });
  });
  return result;
}

function MrrDashboard_ensureDivision_(divisions, key, label) {
  if (!divisions[key]) {
    divisions[key] = {
      key: key,
      label: label || key,
      totalKey: MRR_DASHBOARD_TOTAL_KEY,
      totalLabel: key + '全体',
      periods: [],
      weeks: [],
      depts: [],
      data: {},
      periodDeals: {},
      deals: {}
    };
  }
  return divisions[key];
}

function MrrDashboard_addWeek_(division, week) {
  if (!division._weekSeen) division._weekSeen = {};
  if (division._weekSeen[week.key]) return;
  division._weekSeen[week.key] = true;
  division.weeks.push(week);
}

function MrrDashboard_addPeriod_(division, periodKey) {
  if (!division._periodSeen) division._periodSeen = {};
  if (division._periodSeen[periodKey]) return;
  division._periodSeen[periodKey] = true;
  division.periods.push({ key: periodKey, label: MrrDashboard_periodLabel_(periodKey) });
}

function MrrDashboard_addDept_(division, deptMeta) {
  if (!division._deptSeen) division._deptSeen = {};
  if (division._deptSeen[deptMeta.key]) return;
  division._deptSeen[deptMeta.key] = true;
  division.depts.push({ key: deptMeta.key, label: deptMeta.label });
}

function MrrDashboard_sortDivision_(division) {
  division.weeks.sort(function(a, b) { return a.key < b.key ? -1 : a.key > b.key ? 1 : 0; });
  division.periods.sort(function(a, b) {
    var av = MrrDashboard_periodSortValue_(a.key);
    var bv = MrrDashboard_periodSortValue_(b.key);
    return av - bv || a.key.localeCompare(b.key);
  });
  division.depts.sort(function(a, b) { return a.label.localeCompare(b.label, 'ja'); });
  division.defaultPeriod = division.periods.length ? division.periods[0].key : '';
  delete division._weekSeen;
  delete division._periodSeen;
  delete division._deptSeen;
}

function MrrDashboard_fillTotalsAndDeals_(division, dealsByDate) {
  Object.keys(division.data).forEach(function(periodKey) {
    division.weeks.forEach(function(week) {
      var weekRows = division.data[periodKey][week.key] || {};
      var total = MrrDashboard_emptyMetrics_();
      division.depts.forEach(function(dept) {
        var metrics = weekRows[dept.key] || MrrDashboard_emptyMetrics_();
        MrrDashboard_metricDefs_().forEach(function(def) {
          total[def.key] += MrrDashboard_num_(metrics[def.key]);
        });
      });
      weekRows[MRR_DASHBOARD_TOTAL_KEY] = total;
      division.data[periodKey][week.key] = weekRows;
    });
  });

  Object.keys(division.data).forEach(function(periodKey) {
    if (!division.periodDeals[periodKey]) division.periodDeals[periodKey] = {};
    division.weeks.forEach(function(week) {
      var dateDeals = dealsByDate[week.date] || {};
      var weekDeals = {};
      var totalDeals = [];
      division.depts.forEach(function(dept) {
        var rows = (dateDeals[dept.key] || []).filter(function(deal) {
          return MrrDashboard_dealMatchesPeriod_(deal, periodKey);
        }).sort(function(a, b) {
          return Math.abs(b.mrr) - Math.abs(a.mrr);
        });
        weekDeals[dept.key] = rows;
        totalDeals = totalDeals.concat(rows);
      });
      weekDeals[MRR_DASHBOARD_TOTAL_KEY] = totalDeals.sort(function(a, b) {
        return Math.abs(b.mrr) - Math.abs(a.mrr);
      });
      division.periodDeals[periodKey][week.key] = weekDeals;
    });
  });

  var defaultPeriod = division.defaultPeriod || (division.periods[0] && division.periods[0].key) || '';
  division.deals = defaultPeriod ? (division.periodDeals[defaultPeriod] || {}) : {};
}

function MrrDashboard_extractMetrics_(payload) {
  var metrics = MrrDashboard_emptyMetrics_();
  MrrDashboard_metricDefs_().forEach(function(def) {
    metrics[def.key] = MrrDashboard_metricValue_(def.key, payload && payload[def.key]);
  });
  return metrics;
}

function MrrDashboard_emptyMetrics_() {
  var metrics = {};
  MrrDashboard_metricDefs_().forEach(function(def) {
    metrics[def.key] = 0;
  });
  return metrics;
}

function MrrDashboard_metricDefs_() {
  return [
    { key: 'target', label: '目標', color: '#ef4444' },
    { key: 'confirmed', label: '確定', color: '#16a34a' },
    { key: 'fcstAdjusted', label: 'FCST(調整後)', color: '#2563eb' },
    { key: 'fcstCommit', label: 'FCST(コミット)', color: '#7c3aed' },
    { key: 'fcstMin', label: 'FCSTMIN', color: '#64748b' },
    { key: 'fcstMax', label: 'FCSTMAX', color: '#334155' },
    { key: 'expectedMrr', label: '期待MRR', color: '#f59e0b' },
    { key: 'received', label: '受領', color: '#0ea5e9' },
    { key: 'debtMgmt', label: '債権管理', color: '#14b8a6' },
    { key: 'debtMgmtLite', label: '債権管理 Lite', color: '#06b6d4' },
    { key: 'expense', label: '経費', color: '#f97316' },
    { key: 'confirmedReceived', label: '確定 受領', color: '#38bdf8' },
    { key: 'confirmedDebtMgmt', label: '確定 債権管理', color: '#2dd4bf' },
    { key: 'confirmedDebtMgmtLite', label: '確定 債権管理 Lite', color: '#67e8f9' },
    { key: 'confirmedExpense', label: '確定 経費', color: '#fb923c' }
  ];
}

function MrrDashboard_metricValue_(metricKey, value) {
  if (MRR_DASHBOARD_NUMERIC_METRIC_KEYS[metricKey]) {
    return typeof value === 'number' ? MrrDashboard_num_(value) : 0;
  }
  if (!value || typeof value !== 'object') return 0;
  return typeof value.net === 'number' ? MrrDashboard_num_(value.net) : 0;
}

function MrrDashboard_periodLabel_(periodKey) {
  var key = String(periodKey || '').trim();
  if (/^\d+Q[1-4]$/.test(key)) {
    var def = FcstPeriods_getQuarterDefinitionByKey_(key);
    return def && def.label ? def.label : key;
  }
  var month = key.match(/^(\d{4})-(\d{2})$/);
  if (month) return Number(month[2]) + '月';
  return key;
}

function MrrDashboard_periodSortValue_(periodKey) {
  var key = String(periodKey || '').trim();
  var quarter = key.match(/^(\d+)Q([1-4])$/);
  if (quarter) return Number(quarter[1]) * 10 + Number(quarter[2]);
  var month = FcstPeriods_parseMonthKey_(key);
  if (month) return month.year * 100 + month.month + 0.5;
  return 999999;
}

function MrrDashboard_parseJson_(value) {
  try {
    return JSON.parse(String(value || '{}')) || {};
  } catch (e) {
    return {};
  }
}

function MrrDashboard_num_(value) {
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

function MrrDashboard_dealMatchesPeriod_(deal, periodKey) {
  var monthKey = MrrDashboard_normalizeMonthKey_(deal && deal.completedMonth);
  if (!monthKey) return false;
  var months = MrrDashboard_periodMonthSet_(periodKey);
  return !!months[monthKey];
}

function MrrDashboard_periodMonthSet_(periodKey) {
  var key = String(periodKey || '').trim();
  var months = {};
  if (FcstPeriods_parseMonthKey_(key)) {
    months[key] = true;
    return months;
  }
  var quarter = FcstPeriods_getQuarterDefinitionByKey_(key);
  (quarter && quarter.months || []).forEach(function(monthKey) {
    months[monthKey] = true;
  });
  return months;
}

function MrrDashboard_normalizeMonthKey_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}$/.test(text)) return text;
  var compact = text.match(/^(\d{4})(\d{2})$/);
  if (compact) return compact[1] + '-' + compact[2];
  var match = text.match(/^(\d{4})[-\/年](\d{1,2})/);
  if (!match) return '';
  return match[1] + '-' + String(Number(match[2])).padStart(2, '0');
}

function MrrDashboard_normKey_(value) {
  return String(value || '').trim().toUpperCase();
}
