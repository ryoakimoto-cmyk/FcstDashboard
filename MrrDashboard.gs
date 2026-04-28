var MRR_DASHBOARD_CACHE_DEPT = '__mrr_dashboard__';
var MRR_DASHBOARD_CACHE_KEY = 'snapshotData:v1';
var MRR_DASHBOARD_TOTAL_KEY = '__total__';

function mrrDashboard_doGet_() {
  return HtmlService.createHtmlOutputFromFile('mrr-index')
    .setTitle('MRR進捗ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMrrDashboardData() {
  var cached = CacheLayer_read(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, { skipSharedSheet: true });
  if (cached) return cached;

  var data = MrrDashboard_buildFromSnapshots_();
  CacheLayer_write(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, data, { persistToSheet: false });
  return data;
}

function MrrDashboard_buildFromSnapshots_() {
  var catalog = MrrDashboard_buildDeptCatalog_();
  var divisions = {};
  var fcstRows = FcstSnapshot_getAllValues_(4);
  var oppRows = OppListSnapshot_getAllValues_(5);
  var dealsByDivisionDate = MrrDashboard_buildDealsByDivisionDate_(oppRows, catalog);

  fcstRows.forEach(function(row) {
    var snapshotAt = row[0];
    if (!(snapshotAt instanceof Date) || isNaN(snapshotAt)) return;

    var nameRaw = String(row[1] || '').trim();
    var periodKey = String(row[2] || '').trim();
    var payload = MrrDashboard_parseJson_(row[3]);
    var meta = payload.__meta || {};
    var deptKey = String(meta.dept || nameRaw.split(':')[0] || '').trim();
    var deptMeta = catalog.byDeptKey[deptKey];
    if (!deptMeta || !periodKey) return;
    if (!meta.isTotal || meta.totalKind !== 'department' || String(meta.dept || '') !== deptKey) return;

    var divisionKey = deptMeta.divisionKey;
    var division = MrrDashboard_ensureDivision_(divisions, divisionKey, deptMeta.divisionLabel);
    var timestampKey = Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    var dateKey = Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'yyyy-MM-dd');

    MrrDashboard_addWeek_(division, {
      key: timestampKey,
      date: dateKey,
      label: Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'M/d'),
      timestamp: timestampKey
    });
    MrrDashboard_addPeriod_(division, periodKey);
    MrrDashboard_addDept_(division, deptMeta);

    if (!division.data[periodKey]) division.data[periodKey] = {};
    if (!division.data[periodKey][timestampKey]) division.data[periodKey][timestampKey] = {};
    division.data[periodKey][timestampKey][deptKey] = MrrDashboard_extractMetrics_(payload);
  });

  Object.keys(divisions).forEach(function(divisionKey) {
    var division = divisions[divisionKey];
    MrrDashboard_sortDivision_(division);
    MrrDashboard_fillTotalsAndDeals_(division, dealsByDivisionDate[divisionKey] || {});
  });

  var orderedDivisions = ['SS', 'BO'].filter(function(key) {
    return !!divisions[key];
  }).map(function(key) {
    return divisions[key];
  });

  return {
    source: 'snapshot',
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    defaultDivision: divisions.BO ? 'BO' : (orderedDivisions[0] ? orderedDivisions[0].key : ''),
    totalKey: MRR_DASHBOARD_TOTAL_KEY,
    metrics: MrrDashboard_metricDefs_(),
    divisions: orderedDivisions
  };
}

function MrrDashboard_buildDeptCatalog_() {
  var byDeptKey = {};
  var byAnyKey = {};
  var configMap = getDeptConfigMap_();

  Object.keys(configMap || {}).forEach(function(deptKey) {
    var cfg = configMap[deptKey] || {};
    var divisionKey = MrrDashboard_divisionKey_(cfg);
    if (['SS', 'BO'].indexOf(divisionKey) === -1) return;
    var meta = {
      key: deptKey,
      label: String(cfg.label || cfg.departmentName || deptKey).trim(),
      divisionKey: divisionKey,
      divisionLabel: divisionKey
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
  });

  return { byDeptKey: byDeptKey, byAnyKey: byAnyKey };
}

function MrrDashboard_divisionKey_(cfg) {
  var division = String(cfg && cfg.division || '').trim().toUpperCase();
  if (division === 'SS' || division === 'BO') return division;

  var sfSheetKey = String(cfg && cfg.sfSheetKey || '').trim().toUpperCase();
  if (sfSheetKey === 'SS' || sfSheetKey === 'SSCS') return 'SS';
  if (sfSheetKey === 'BO') return 'BO';
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

  division.weeks.forEach(function(week) {
    var dateDeals = dealsByDate[week.date] || {};
    var weekDeals = {};
    var totalDeals = [];
    division.depts.forEach(function(dept) {
      var rows = (dateDeals[dept.key] || []).slice().sort(function(a, b) {
        return Math.abs(b.mrr) - Math.abs(a.mrr);
      });
      weekDeals[dept.key] = rows;
      totalDeals = totalDeals.concat(rows);
    });
    weekDeals[MRR_DASHBOARD_TOTAL_KEY] = totalDeals.sort(function(a, b) {
      return Math.abs(b.mrr) - Math.abs(a.mrr);
    });
    division.deals[week.key] = weekDeals;
  });
}

function MrrDashboard_extractMetrics_(payload) {
  var metrics = MrrDashboard_emptyMetrics_();
  MrrDashboard_metricDefs_().forEach(function(def) {
    metrics[def.key] = MrrDashboard_metricValue_(payload && payload[def.key]);
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

function MrrDashboard_metricValue_(value) {
  if (typeof value === 'number') return MrrDashboard_num_(value);
  return MrrDashboard_num_(value && value.net);
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

function MrrDashboard_normKey_(value) {
  return String(value || '').trim().toUpperCase();
}
