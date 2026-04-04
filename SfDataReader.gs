function SfDataReader_getAggregated(users, targets) {
  const sheet = SfDataReader_getSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 57);
  const rows = lastRow <= 2 ? [] : sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const membersByName = {};

  rows.forEach(function(row) {
    const date = SfDataReader_parseDate_(row[0]);
    if (!date) return;
    if (date.getFullYear() !== FCST_TARGET_YEAR) return;

    const month = date.getMonth() + 1;
    if (FCST_TARGET_MONTHS.indexOf(month) === -1) return;

    const userName = String(row[3] || '').trim();
    if (!userName) return;

    const user = users[userName] || {};
    const member = membersByName[userName] || SfDataReader_createMember_(userName, user.group || '', user.dept || '', false);
    const blockKey = 'M' + month;
    const metric = member[blockKey];
    const bucket = row[24] === 'New+Exp' ? 'newExp' : row[24] === 'Churn' ? 'churn' : null;

    SfDataReader_addBreakdownValue_(metric.fcstCommit, bucket, row[38]);
    metric.fcstMax += SfDataReader_toNumber_(row[40]);
    SfDataReader_addBreakdownValue_(metric.confirmed, bucket, row[10]);
    SfDataReader_addBreakdownValue_(metric.expectedMrr, bucket, row[37]);
    SfDataReader_addBreakdownValue_(metric.received, bucket, row[53]);
    SfDataReader_addBreakdownValue_(metric.debtMgmt, bucket, row[54]);
    SfDataReader_addBreakdownValue_(metric.debtMgmtLite, bucket, row[55]);
    SfDataReader_addBreakdownValue_(metric.expense, bucket, row[56]);

    if (SfDataReader_toBoolean_(row[32])) {
      metric.keyDeals.push({
        company: SfDataReader_stripLegalForm_(String(row[25] || '').trim()),
        monthlyMrr: SfDataReader_toNumber_(row[10]),
        phase: String(row[30] || '').trim(),
        fcst: SfDataReader_toNumber_(row[38]),
        oppId: String(row[21] || '').trim(),
      });
    }

    membersByName[userName] = member;
  });

  const members = Object.keys(membersByName).map(function(name) {
    const member = membersByName[name];
    SfDataReader_applyTargets_(member, targets);
    member.Q = SfDataReader_sumMetrics_([member.M5, member.M6, member.M7]);
    member.sortOrder = users[name] ? users[name].sortOrder || 0 : 0;
    return member;
  });

  const groupTotals = SfDataReader_buildGroupTotals_(members, targets);
  const overallTotal = SfDataReader_buildOverallTotal_(members, targets);
  const sortedMembers = members.sort(function(a, b) {
    return (a.sortOrder || 0) - (b.sortOrder || 0) || a.name.localeCompare(b.name, 'ja');
  });
  const sortedGroupTotals = groupTotals.sort(function(a, b) {
    return a.name.localeCompare(b.name, 'ja');
  });

  return {
    lastUpdated: SfDataReader_extractLastUpdated_(sheet.getRange(1, 1).getDisplayValue()),
    members: overallTotal ? [overallTotal].concat(sortedGroupTotals, sortedMembers).map(SfDataReader_finalizeMember_) : sortedGroupTotals.concat(sortedMembers).map(SfDataReader_finalizeMember_),
  };
}

function SfDataReader_getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SF_DATA_SHEET_NAME);
  if (!sheet) throw new Error('SFデータ更新シートが見つかりません');
  return sheet;
}

function SfDataReader_parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value || '').trim();
  if (!text || !/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;

  const date = new Date(text + 'T00:00:00');
  if (isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function SfDataReader_createMember_(name, group, dept, isTotal) {
  return {
    name: name,
    group: group,
    dept: dept,
    isTotal: isTotal,
    sortOrder: 0,
    Q: SfDataReader_createMetric_(),
    M5: SfDataReader_createMetric_(),
    M6: SfDataReader_createMetric_(),
    M7: SfDataReader_createMetric_(),
  };
}

function SfDataReader_createMetric_() {
  return {
    target: 0,
    fcstCommit: SfDataReader_createBreakdown_(),
    fcstMax: 0,
    confirmed: SfDataReader_createBreakdown_(),
    expectedMrr: SfDataReader_createBreakdown_(),
    received: SfDataReader_createBreakdown_(),
    debtMgmt: SfDataReader_createBreakdown_(),
    debtMgmtLite: SfDataReader_createBreakdown_(),
    expense: SfDataReader_createBreakdown_(),
    keyDeals: [],
  };
}

function SfDataReader_createBreakdown_() {
  return { net: 0, newExp: 0, churn: 0 };
}

function SfDataReader_addBreakdownValue_(breakdown, bucket, value) {
  const amount = SfDataReader_toNumber_(value);
  breakdown.net += amount;
  if (bucket) breakdown[bucket] += amount;
}

function SfDataReader_buildGroupTotals_(members, targets) {
  const totals = {};

  members.forEach(function(member) {
    const key = member.group || '';
    if (!key) return;

    if (!totals[key]) {
      totals[key] = SfDataReader_createMember_(key + '合計', key, member.dept || '', true);
    }

    totals[key].M5 = SfDataReader_sumMetrics_([totals[key].M5, member.M5]);
    totals[key].M6 = SfDataReader_sumMetrics_([totals[key].M6, member.M6]);
    totals[key].M7 = SfDataReader_sumMetrics_([totals[key].M7, member.M7]);
  });

  return Object.keys(totals).map(function(key) {
    const member = totals[key];
    SfDataReader_applyTargets_(member, targets);
    member.Q = SfDataReader_sumMetrics_([member.M5, member.M6, member.M7]);
    return member;
  });
}

function SfDataReader_buildOverallTotal_(members, targets) {
  if (!members.length) return null;

  const total = SfDataReader_createMember_('BOAM合計', 'BOAM', '', true);
  members.forEach(function(member) {
    total.M5 = SfDataReader_sumMetrics_([total.M5, member.M5]);
    total.M6 = SfDataReader_sumMetrics_([total.M6, member.M6]);
    total.M7 = SfDataReader_sumMetrics_([total.M7, member.M7]);
  });

  SfDataReader_applyTargets_(total, targets);
  total.Q = SfDataReader_sumMetrics_([total.M5, total.M6, total.M7]);
  return total;
}

function SfDataReader_applyTargets_(member, targets) {
  FCST_TARGET_MONTHS.forEach(function(month) {
    const blockKey = 'M' + month;
    const ym = String(FCST_TARGET_YEAR) + String(month).padStart(2, '0');
    const nameKey = member.name + '|' + ym;
    const groupKey = member.group + '|' + ym;
    member[blockKey].target = targets.hasOwnProperty(nameKey) ? targets[nameKey] : targets[groupKey] || 0;
  });
}

function SfDataReader_sumMetrics_(metrics) {
  return metrics.reduce(function(sum, metric) {
    sum.target += metric.target || 0;
    sum.fcstCommit.net += metric.fcstCommit.net || 0;
    sum.fcstCommit.newExp += metric.fcstCommit.newExp || 0;
    sum.fcstCommit.churn += metric.fcstCommit.churn || 0;
    sum.fcstMax += metric.fcstMax || 0;
    sum.confirmed.net += metric.confirmed.net || 0;
    sum.confirmed.newExp += metric.confirmed.newExp || 0;
    sum.confirmed.churn += metric.confirmed.churn || 0;
    sum.expectedMrr.net += metric.expectedMrr.net || 0;
    sum.expectedMrr.newExp += metric.expectedMrr.newExp || 0;
    sum.expectedMrr.churn += metric.expectedMrr.churn || 0;
    sum.received.net += metric.received.net || 0;
    sum.received.newExp += metric.received.newExp || 0;
    sum.received.churn += metric.received.churn || 0;
    sum.debtMgmt.net += metric.debtMgmt.net || 0;
    sum.debtMgmt.newExp += metric.debtMgmt.newExp || 0;
    sum.debtMgmt.churn += metric.debtMgmt.churn || 0;
    sum.debtMgmtLite.net += metric.debtMgmtLite.net || 0;
    sum.debtMgmtLite.newExp += metric.debtMgmtLite.newExp || 0;
    sum.debtMgmtLite.churn += metric.debtMgmtLite.churn || 0;
    sum.expense.net += metric.expense.net || 0;
    sum.expense.newExp += metric.expense.newExp || 0;
    sum.expense.churn += metric.expense.churn || 0;
    sum.keyDeals = sum.keyDeals.concat(metric.keyDeals || []);
    return sum;
  }, SfDataReader_createMetric_());
}

function SfDataReader_finalizeMember_(member) {
  member.M5 = SfDataReader_finalizeMetric_(member.M5);
  member.M6 = SfDataReader_finalizeMetric_(member.M6);
  member.M7 = SfDataReader_finalizeMetric_(member.M7);
  member.Q = SfDataReader_finalizeMetric_(member.Q);
  delete member.sortOrder;
  return member;
}

function SfDataReader_finalizeMetric_(metric) {
  function finalizeBreakdown(b) {
    return { net: b.net, newExp: b.newExp, churn: b.churn };
  }
  return {
    target: metric.target,
    fcstCommit: finalizeBreakdown(metric.fcstCommit),
    fcstMax: metric.fcstMax,
    confirmed: finalizeBreakdown(metric.confirmed),
    expectedMrr: finalizeBreakdown(metric.expectedMrr),
    received: finalizeBreakdown(metric.received),
    debtMgmt: finalizeBreakdown(metric.debtMgmt),
    debtMgmtLite: finalizeBreakdown(metric.debtMgmtLite),
    expense: finalizeBreakdown(metric.expense),
    keyDeals: SfDataReader_unique_(metric.keyDeals.map(function(kd) {
      return JSON.stringify(kd);
    })).map(function(s) { return JSON.parse(s); }),
  };
}

function SfDataReader_unique_(values) {
  const map = {};
  return values.filter(function(value) {
    if (!value || map[value]) return false;
    map[value] = true;
    return true;
  });
}

function SfDataReader_extractLastUpdated_(title) {
  const text = String(title || '').trim();
  const match = text.match(/\d{4}[\/-]\d{1,2}[\/-]\d{1,2}/);
  if (!match) {
    return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  const normalized = match[0].replace(/\//g, '-');
  const date = new Date(normalized + 'T00:00:00');
  if (isNaN(date)) {
    return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function SfDataReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function SfDataReader_toBoolean_(value) {
  if (value === true || value === false) return value;
  return /^(true|yes|1|y)$/i.test(String(value || '').trim());
}

function SfDataReader_stripLegalForm_(name) {
  return String(name || '')
    .replace(/株式会社/g, '')
    .replace(/合同会社/g, '')
    .replace(/有限会社/g, '')
    .replace(/合名会社/g, '')
    .replace(/合資会社/g, '')
    .replace(/一般社団法人/g, '')
    .replace(/一般財団法人/g, '')
    .replace(/公益社団法人/g, '')
    .replace(/公益財団法人/g, '')
    .replace(/[（(](株|有|同|資|名)[)）]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
