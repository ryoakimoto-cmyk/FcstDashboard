function MrrDashboard_getBoData_() {
  var deptKeys = MrrDashboard_getDivisionDeptKeys_('BO');
  var weekLabels = {};
  var datesSeen = {};
  var data = {};

  deptKeys.forEach(function(deptKey) {
    (FcstSnapshot_getSnapshotDates(deptKey) || []).forEach(function(dateStr) {
      if (dateStr) datesSeen[dateStr] = true;
    });
  });

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
      var metric = MrrDashboard_buildBoDeptMetric_(deptKey, dateStr);
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
    weeks: weeks,
    weekLabels: weekLabels,
    depts: deptLabels,
    data: data
  };
}

function MrrDashboard_buildBoDeptMetric_(deptKey, snapshotDate) {
  var fcstSnapshot = FcstSnapshot_getDataByDate(deptKey, snapshotDate) || {};
  var totalMember = MrrDashboard_findDepartmentTotalMember_(fcstSnapshot.members || [], deptKey);
  var periodKey = MrrDashboard_pickSnapshotPeriodKey_(fcstSnapshot.periodOptions || [], snapshotDate);
  var periodMetric = totalMember && periodKey ? (totalMember[periodKey] || null) : null;
  var keyDealsData = MrrDashboard_getBoKeyDeals_(deptKey, snapshotDate);

  return {
    target: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.target),
    actual: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.confirmed),
    expectedMrr: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.expectedMrr),
    fcst: MrrDashboard_getBreakdownNet_(periodMetric && periodMetric.fcstCommit),
    keyDeal: MrrDashboard_formatLegacyDeals_(keyDealsData),
    keyDealsData: keyDealsData
  };
}

function MrrDashboard_getBoKeyDeals_(deptKey, snapshotDate) {
  var snapshot = OppListSnapshot_getByDate(deptKey, snapshotDate) || {};
  var deptLabel = MrrDashboard_getDeptLabel_(deptKey);
  var deals = (snapshot.rows || []).filter(function(row) {
    return !!(row && row.keyDeal);
  }).map(function(row) {
    return {
      dept: deptLabel,
      company: String(row.dealName || ''),
      mrr: Number(row.mrr) || 0,
      phase: String(row.phase || ''),
      yomi: Number(row.fcstCommit) || 0,
      oppId: String(row.oppId || ''),
      completedMonth: String(row.completedMonth || ''),
      owner: String(row.subOwner || '')
    };
  });

  deals.sort(function(left, right) {
    return Math.abs(right.mrr) - Math.abs(left.mrr);
  });
  return deals;
}

function MrrDashboard_getDivisionDeptKeys_(division) {
  return Object.keys(DEPT_CONFIG).filter(function(deptKey) {
    return DEPT_CONFIG[deptKey] && DEPT_CONFIG[deptKey].division === division;
  });
}

function MrrDashboard_getDeptLabel_(deptKey) {
  return (DEPT_CONFIG[deptKey] && DEPT_CONFIG[deptKey].label) || String(deptKey || '');
}

function MrrDashboard_findDepartmentTotalMember_(members, deptKey) {
  return (members || []).find(function(member) {
    return SharedAppState_isDepartmentTotal_(member, deptKey);
  }) || null;
}

function MrrDashboard_pickSnapshotPeriodKey_(periodOptions, snapshotDate) {
  var monthKeys = [];
  (periodOptions || []).forEach(function(option) {
    if (!option) return;
    if (/^\d{4}-\d{2}$/.test(String(option.key || '')) && monthKeys.indexOf(option.key) === -1) {
      monthKeys.push(option.key);
    }
    (option.months || []).forEach(function(monthKey) {
      if (/^\d{4}-\d{2}$/.test(String(monthKey || '')) && monthKeys.indexOf(monthKey) === -1) {
        monthKeys.push(monthKey);
      }
    });
  });
  monthKeys.sort();
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
  return {
    target: 0,
    actual: 0,
    expectedMrr: 0,
    fcst: 0,
    keyDeal: '',
    keyDealsData: []
  };
}

function MrrDashboard_accumulateMetric_(target, metric) {
  if (!target || !metric) return;
  target.target += Number(metric.target) || 0;
  target.actual += Number(metric.actual) || 0;
  target.expectedMrr += Number(metric.expectedMrr) || 0;
  target.fcst += Number(metric.fcst) || 0;
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
