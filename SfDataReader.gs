function SfDataReader_getAggregated(deptKey, contextOrUsers, legacyTargets) {
  var context = SfDataReader_normalizeContext_(deptKey, contextOrUsers, legacyTargets);
  var sheet = SfDataReader_getSheet_(deptKey);
  if (!sheet) {
    return { lastUpdated: '', periodOptions: [], members: [] };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 57);
  var headers = lastRow >= 2 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] : [];
  if (deptKey === 'SSSMBCS') headers = normalizeSSCSHeaders_(headers);
  var headerMap = SfDataReader_buildHeaderMap_(headers);
  var rows = lastRow <= 2 ? [] : sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  var membersByName = {};
  var monthKeyMap = {};

  rows.forEach(function(row) {
    var date = SfDataReader_parseDate_(SfDataReader_valueByKeys_(row, headerMap, ['完了月', '計上月', 'Close Month'], 0));
    if (!date || !FcstPeriods_isSupportedDate_(date)) return;

    var monthKey = FcstPeriods_formatMonthKey_(date);
    monthKeyMap[monthKey] = true;

    var sourceName = SfDataReader_formatCell_(SfDataReader_valueByKeys_(row, headerMap, ['サブオーナー', '担当者', '担当名'], 3)).trim();
    if (!sourceName) return;

    var user = SfDataReader_resolveUser_(context, sourceName, monthKey);
    var member = membersByName[sourceName] || SfDataReader_createMember_(
      sourceName,
      user.displayName || sourceName,
      user.group || '',
      user.dept || '',
      false
    );
    member.displayName = user.displayName || member.displayName || member.name;
    member.group = user.group || member.group || '';
    member.groupCode = user.groupCode || member.groupCode || member.group || '';
    member.dept = user.dept || member.dept || '';
    member.totalKind = SHARED_TOTAL_KIND.INDIVIDUAL;
    member.sortOrder = user.sortOrder || member.sortOrder || 0;

    var metric = SfDataReader_getOrCreateMonthMetric_(member, monthKey);
    var typeValue = SfDataReader_formatCell_(SfDataReader_valueByKeys_(row, headerMap, ['種別', '案件種別'], 24)).trim();
    var bucket = typeValue === 'New+Exp' ? 'newExp' : typeValue === 'Churn' ? 'churn' : null;
    var fcstCommitValue = SfDataReader_valueByKeys_(row, headerMap, ['FCST(コミット)(換算値)', 'FCST(コミット)', 'FCST（コミット）'], 38);

    SfDataReader_addBreakdownValue_(metric.fcstCommit, bucket, fcstCommitValue);
    metric.fcstMin += SfDataReader_toNumber_(SfDataReader_valueByKeys_(row, headerMap, ['FCST(MIN)(換算値)', 'FCST(MIN)', 'FCST（MIN）'], 39));
    metric.fcstMax += SfDataReader_toNumber_(SfDataReader_valueByKeys_(row, headerMap, ['FCST(MAX)(換算値)', 'FCST(MAX)', 'FCST（MAX）'], 40));
    SfDataReader_addBreakdownValue_(metric.confirmed, bucket, SfDataReader_valueByKeys_(row, headerMap, ['MRR', '受注MRR'], 10));
    SfDataReader_addBreakdownValue_(metric.expectedMrr, bucket, SfDataReader_valueByKeys_(row, headerMap, ['Expected MRR', 'ExpectedMRR'], 37));
    SfDataReader_addBreakdownValue_(metric.received, bucket, SfDataReader_valueByKeys_(row, headerMap, ['受領'], 53));
    SfDataReader_addBreakdownValue_(metric.debtMgmt, bucket, SfDataReader_valueByKeys_(row, headerMap, ['債権管理', '債権管理回収'], 54));
    SfDataReader_addBreakdownValue_(metric.debtMgmtLite, bucket, SfDataReader_valueByKeys_(row, headerMap, ['債権管理Lite', '債権管理 Lite'], 55));
    SfDataReader_addBreakdownValue_(metric.expense, bucket, SfDataReader_valueByKeys_(row, headerMap, ['経費'], 56));

    if (SfDataReader_toBoolean_(SfDataReader_valueByKeys_(row, headerMap, ['Key Deal フラグ', 'KeyDeal'], 32))) {
      metric.keyDeals.push({
        company: SfDataReader_stripLegalForm_(SfDataReader_formatCell_(SfDataReader_valueByKeys_(row, headerMap, ['取引先名', '会社名'], 25)).trim()),
        monthlyMrr: SfDataReader_toNumber_(SfDataReader_valueByKeys_(row, headerMap, ['MRR', '受注MRR'], 10)),
        phase: SfDataReader_formatCell_(SfDataReader_valueByKeys_(row, headerMap, ['フェーズ'], 30)).trim(),
        fcst: SfDataReader_toNumber_(fcstCommitValue),
        oppId: SfDataReader_formatCell_(SfDataReader_valueByKeys_(row, headerMap, ['ID', '案件ID'], 21)).trim()
      });
    }

    membersByName[sourceName] = member;
  });

  Object.keys(context.targets || {}).forEach(function(key) {
    var match = String(key).match(/\|(\d{6})$/);
    if (!match) return;
    monthKeyMap[match[1].slice(0, 4) + '-' + match[1].slice(4, 6)] = true;
  });

  var periodOptions = FcstPeriods_buildDefinitionsFromMonthKeys_(Object.keys(monthKeyMap));
  var members = Object.keys(membersByName).map(function(sourceName) {
    var member = membersByName[sourceName];
    SfDataReader_applyTargets_(member, context.targets, periodOptions);
    SfDataReader_buildPeriods_(member, periodOptions);
    return member;
  });

  var groupTotals = SfDataReader_buildGroupTotals_(members, periodOptions);
  var overallTotal = SfDataReader_buildOverallTotal_(members, periodOptions, deptKey);
  var sortedMembers = members.sort(function(a, b) {
    return (a.sortOrder || 0) - (b.sortOrder || 0) || (a.displayName || a.name).localeCompare((b.displayName || b.name), 'ja');
  });
  var sortedGroupTotals = groupTotals.sort(function(a, b) {
    return (a.displayName || a.name).localeCompare((b.displayName || b.name), 'ja');
  });

  return {
    lastUpdated: SfDataReader_extractLastUpdated_(sheet.getRange(1, 1)),
    periodOptions: periodOptions,
    members: (overallTotal ? [overallTotal].concat(sortedGroupTotals, sortedMembers) : sortedGroupTotals.concat(sortedMembers))
      .map(function(member) {
        return SfDataReader_finalizeMember_(member, periodOptions);
      })
  };
}

function SfDataReader_normalizeContext_(deptKey, contextOrUsers, legacyTargets) {
  if (legacyTargets !== undefined) {
    return {
      users: AssignmentMaster_normalizeUserMap_(deptKey, contextOrUsers || {}),
      monthlyUsers: {},
      targets: SfDataReader_normalizeLegacyTargets_(legacyTargets)
    };
  }
  return AssignmentMaster_normalizeContext_(deptKey, contextOrUsers || {});
}

function SfDataReader_normalizeLegacyTargets_(targets) {
  var normalized = {};
  Object.keys(targets || {}).forEach(function(key) {
    normalized[key] = SfDataReader_toBreakdown_(targets[key]);
  });
  return normalized;
}

function SfDataReader_resolveUser_(context, sourceName, monthKey) {
  var monthly = (context.monthlyUsers || {})[sourceName + '|' + monthKey];
  if (monthly) return monthly;
  return (context.users || {})[sourceName] || {};
}

function SfDataReader_getSheet_(deptKey) {
  return getSfDataSheet_(deptKey);
}

function SfDataReader_buildHeaderMap_(headers) {
  var map = {};
  (headers || []).forEach(function(header, idx) {
    var raw = String(header || '').trim();
    if (!raw) return;
    map[raw] = idx;
    map[SfDataReader_normalize_(raw)] = idx;
  });
  return map;
}

function SfDataReader_valueByKeys_(row, headerMap, keys, fallbackIndex) {
  for (var i = 0; i < (keys || []).length; i++) {
    var key = keys[i];
    var normalized = SfDataReader_normalize_(key);
    if (headerMap.hasOwnProperty(key)) return row[headerMap[key]];
    if (headerMap.hasOwnProperty(normalized)) return row[headerMap[normalized]];
  }
  if (typeof fallbackIndex === 'number' && fallbackIndex >= 0 && fallbackIndex < row.length) {
    return row[fallbackIndex];
  }
  return '';
}

function SfDataReader_normalize_(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[()（）]/g, '')
    .toLowerCase();
}

function SfDataReader_parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var text = String(value || '').trim();
  if (!text || !/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  var date = new Date(text + 'T00:00:00');
  if (isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function SfDataReader_formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value === null || value === undefined ? '' : String(value);
}

function SfDataReader_createMember_(name, displayName, group, dept, isTotal) {
  return {
    name: name,
    displayName: displayName || name,
    group: group,
    dept: dept,
    isTotal: isTotal,
    sortOrder: 0,
    monthMap: {}
  };
}

function SfDataReader_getOrCreateMonthMetric_(member, monthKey) {
  if (!member.monthMap[monthKey]) member.monthMap[monthKey] = SfDataReader_createMetric_();
  return member.monthMap[monthKey];
}

function SfDataReader_createMetric_() {
  return {
    target: SfDataReader_createBreakdown_(),
    fcstCommit: SfDataReader_createBreakdown_(),
    fcstMin: 0,
    fcstMax: 0,
    confirmed: SfDataReader_createBreakdown_(),
    expectedMrr: SfDataReader_createBreakdown_(),
    received: SfDataReader_createBreakdown_(),
    debtMgmt: SfDataReader_createBreakdown_(),
    debtMgmtLite: SfDataReader_createBreakdown_(),
    expense: SfDataReader_createBreakdown_(),
    keyDeals: []
  };
}

function SfDataReader_createBreakdown_() {
  return { net: 0, newExp: 0, churn: 0 };
}

function SfDataReader_addBreakdownValue_(breakdown, bucket, value) {
  var amount = SfDataReader_toNumber_(value);
  breakdown.net += amount;
  if (bucket) breakdown[bucket] += amount;
}

function SfDataReader_applyTargets_(member, targets, periodOptions) {
  (periodOptions || []).forEach(function(option) {
    (option.months || []).forEach(function(monthKey) {
      var ym = monthKey.replace('-', '');
      var metric = SfDataReader_getOrCreateMonthMetric_(member, monthKey);
      metric.target = SfDataReader_toBreakdown_(targets[member.name + '|' + ym]);
    });
  });
}

function SfDataReader_buildGroupTotals_(members, periodOptions) {
  var totals = {};

  members.forEach(function(member) {
    var key = member.groupCode || member.group || '';
    if (!key) return;
    if (!totals[key]) {
      totals[key] = SfDataReader_createMember_(key + 'グループ', member.group || key, member.group || key, member.dept || '', true);
      totals[key].groupCode = key;
      totals[key].totalKind = SHARED_TOTAL_KIND.GROUP;
    }
    Object.keys(member.monthMap || {}).forEach(function(monthKey) {
      totals[key].monthMap[monthKey] = SfDataReader_sumMetrics_([
        totals[key].monthMap[monthKey] || SfDataReader_createMetric_(),
        member.monthMap[monthKey]
      ]);
    });
  });

  return Object.keys(totals).map(function(key) {
    var member = totals[key];
    SfDataReader_buildPeriods_(member, periodOptions);
    return member;
  });
}

function SfDataReader_buildOverallTotal_(members, periodOptions, deptKey) {
  if (!members.length) return null;
  var total = SfDataReader_createMember_(SHARED_ALL_GROUP_LABEL, SHARED_ALL_GROUP_LABEL, SHARED_ALL_GROUP_LABEL, deptKey || '', true);
  total.groupCode = String(deptKey || '');
  total.totalKind = SHARED_TOTAL_KIND.DEPARTMENT;
  members.forEach(function(member) {
    Object.keys(member.monthMap || {}).forEach(function(monthKey) {
      total.monthMap[monthKey] = SfDataReader_sumMetrics_([
        total.monthMap[monthKey] || SfDataReader_createMetric_(),
        member.monthMap[monthKey]
      ]);
    });
  });
  SfDataReader_buildPeriods_(total, periodOptions);
  return total;
}

function SfDataReader_buildPeriods_(member, periodOptions) {
  (periodOptions || []).forEach(function(option) {
    var months = option.months || [];
    months.forEach(function(monthKey) {
      member[monthKey] = member.monthMap[monthKey] || SfDataReader_createMetric_();
    });
    member[option.key] = SfDataReader_sumMetrics_(months.map(function(monthKey) {
      return member.monthMap[monthKey] || SfDataReader_createMetric_();
    }));
  });
}

function SfDataReader_sumMetrics_(metrics) {
  return (metrics || []).reduce(function(sum, metric) {
    metric = metric || SfDataReader_createMetric_();
    sum.target.net += (metric.target && metric.target.net) || 0;
    sum.target.newExp += (metric.target && metric.target.newExp) || 0;
    sum.target.churn += (metric.target && metric.target.churn) || 0;
    sum.fcstCommit.net += metric.fcstCommit.net || 0;
    sum.fcstCommit.newExp += metric.fcstCommit.newExp || 0;
    sum.fcstCommit.churn += metric.fcstCommit.churn || 0;
    sum.fcstMin += metric.fcstMin || 0;
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

function SfDataReader_finalizeMember_(member, periodOptions) {
  (periodOptions || []).forEach(function(option) {
    (option.months || []).forEach(function(monthKey) {
      member[monthKey] = SfDataReader_finalizeMetric_(member[monthKey] || SfDataReader_createMetric_());
    });
    member[option.key] = SfDataReader_finalizeMetric_(member[option.key] || SfDataReader_createMetric_());
  });
  delete member.sortOrder;
  delete member.monthMap;
  return member;
}

function SfDataReader_finalizeMetric_(metric) {
  function finalizeBreakdown(breakdown) {
    return {
      net: (breakdown && breakdown.net) || 0,
      newExp: (breakdown && breakdown.newExp) || 0,
      churn: (breakdown && breakdown.churn) || 0
    };
  }
  return {
    target: finalizeBreakdown(metric.target),
    fcstCommit: finalizeBreakdown(metric.fcstCommit),
    fcstMin: metric.fcstMin || 0,
    fcstMax: metric.fcstMax || 0,
    confirmed: finalizeBreakdown(metric.confirmed),
    expectedMrr: finalizeBreakdown(metric.expectedMrr),
    received: finalizeBreakdown(metric.received),
    debtMgmt: finalizeBreakdown(metric.debtMgmt),
    debtMgmtLite: finalizeBreakdown(metric.debtMgmtLite),
    expense: finalizeBreakdown(metric.expense),
    keyDeals: SfDataReader_unique_((metric.keyDeals || []).map(function(keyDeal) {
      return JSON.stringify(keyDeal);
    })).map(function(text) {
      return JSON.parse(text);
    })
  };
}

function SfDataReader_toBreakdown_(value) {
  if (!value && value !== 0) return SfDataReader_createBreakdown_();
  if (typeof value === 'number') {
    return { net: value, newExp: 0, churn: 0 };
  }
  return {
    net: SfDataReader_toNumber_(value.net),
    newExp: SfDataReader_toNumber_(value.newExp),
    churn: SfDataReader_toNumber_(value.churn)
  };
}

function SfDataReader_unique_(values) {
  var map = {};
  return (values || []).filter(function(value) {
    if (!value || map[value]) return false;
    map[value] = true;
    return true;
  });
}

function SfDataReader_extractLastUpdated_(cellOrText) {
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
    .replace(/有限会社/g, '')
    .replace(/合同会社/g, '')
    .replace(/一般社団法人/g, '')
    .replace(/一般財団法人/g, '')
    .replace(/公益社団法人/g, '')
    .replace(/公益財団法人/g, '')
    .replace(/[()（）]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

