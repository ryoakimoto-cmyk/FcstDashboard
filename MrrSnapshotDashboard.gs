function MrrDashboard_getSnapshotData_(selection, options) {
  var selectedDivision = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  var opts = options || {};
  var requestedPeriodKey = MrrDashboard_normalizePeriodKey_(opts.periodKey);
  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var dateBatch = { dateSet: {}, dates: [], hasMore: false, oldestDate: '' };
  var fcst = { dates: {}, metricsByDeptDate: {} };
  if (!opts.currentOnly) {
    dateBatch = MrrDashboard_getFcstSnapshotDateBatch_(deptKeys, opts.beforeDate || '', opts.limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT);
    fcst = MrrDashboard_readFcstSnapshots_(deptKeys, dateBatch.dateSet);
  }
  var currentCache = (opts.currentOnly || opts.includeCurrent)
    ? MrrDashboard_readCurrentCacheMap_(deptKeys)
    : null;
  var periodOptions = (opts.currentOnly || opts.includeCurrent)
    ? MrrDashboard_getCurrentPeriodOptions_(deptKeys, currentCache)
    : [];
  if (!periodOptions.length) {
    periodOptions = MrrDashboard_getSnapshotPeriodOptions_(fcst);
  }
  var selectedPeriod = MrrDashboard_resolvePeriodKey_(periodOptions, requestedPeriodKey);
  var selectedQuarterKey = MrrDashboard_getQuarterKeyForPeriod_(periodOptions, selectedPeriod);
  var weeks = dateBatch.dates.slice().sort();
  var weekLabels = {};
  var data = {};
  var diagnostics = { current: [] };
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
      var metric = MrrDashboard_buildDeptMetricFromContext_(deptKey, dateStr, fcst, selectedPeriod);
      rowMap[deptLabel] = metric;
      MrrDashboard_accumulateMetric_(total, metric);
    });

    rowMap[totalKey] = total;
    data[dateStr] = rowMap;
  });

  if (opts.currentOnly || opts.includeCurrent) {
    MrrDashboard_addCurrentData_(deptKeys, weeks, weekLabels, data, totalKey, diagnostics, selectedPeriod, currentCache);
  }

  return {
    division: selectedDivision,
    divisionLabel: MrrDashboard_getDivisionLabel_(selectedDivision),
    choices: MrrDashboard_getDivisionChoices_(),
    totalDeptKey: totalKey,
    allLabel: allLabel,
    metricDefinitions: MrrDashboard_getMetricDefinitions_(),
    periodOptions: periodOptions,
    selectedPeriod: selectedPeriod,
    selectedQuarterKey: selectedQuarterKey,
    weeks: weeks,
    weekLabels: weekLabels,
    depts: deptLabels,
    data: data,
    diagnostics: diagnostics,
    history: {
      beforeDate: opts.beforeDate || '',
      limit: opts.limit || MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT,
      hasMore: dateBatch.hasMore,
      oldestDate: dateBatch.oldestDate || ''
    }
  };
}

function getMrrDashboardDeals(selection, dateStr, deptLabel, periodKey) {
  var selectedDivision = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  var requestedDate = String(dateStr || '').trim();
  var requestedPeriodKey = MrrDashboard_normalizePeriodKey_(periodKey);
  var isLive = requestedDate === MRR_DASHBOARD_LIVE_KEY;
  var snapshotDate = isLive ? MRR_DASHBOARD_LIVE_KEY : MrrDashboard_normalizeSnapshotDate_(requestedDate);
  if (!snapshotDate) return { deals: [] };

  var divisionKeys = MrrDashboard_getSelectedDivisionKeys_(selectedDivision);
  var deptKeys = MrrDashboard_getDeptKeysForDivisions_(divisionKeys);
  var scopedDeptKeys = MrrDashboard_filterDeptKeysByLabel_(deptKeys, deptLabel);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + 'deals:' + selectedDivision + ':' + snapshotDate + ':' + (deptLabel || 'all') + ':period:' + (requestedPeriodKey || 'default');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var periodFilterByDept = isLive
    ? MrrDashboard_getCurrentPeriodFilterByDept_(scopedDeptKeys, requestedPeriodKey)
    : MrrDashboard_getSnapshotPeriodFilterByDept_(scopedDeptKeys, snapshotDate, requestedPeriodKey);
  var deals = isLive
    ? MrrDashboard_readCurrentOppDeals_(scopedDeptKeys, periodFilterByDept)
    : MrrDashboard_readOppSnapshotDeals_(scopedDeptKeys, snapshotDate, periodFilterByDept);
  var result = { deals: deals };
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
}

function MrrDashboard_addCurrentData_(deptKeys, weeks, weekLabels, data, totalKey, diagnostics, periodKey, currentCache) {
  var liveRows = {};
  var total = MrrDashboard_emptyMetric_();
  var hasLive = false;
  currentCache = currentCache || MrrDashboard_readCurrentCacheMap_(deptKeys);

  (deptKeys || []).forEach(function(deptKey) {
    var metric = MrrDashboard_buildCurrentDeptMetric_(deptKey, diagnostics, currentCache, periodKey);
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

function MrrDashboard_readCurrentCacheMap_(deptKeys) {
  var result = {};
  var missing = [];

  (deptKeys || []).forEach(function(deptKey) {
    try {
      var live = CacheLayer_read(deptKey, 'initData', { skipSharedSheet: true });
      if (live) {
        result[deptKey] = { live: live, source: 'script_cache', error: '' };
        return;
      }
    } catch (e) {
      result[deptKey] = {
        live: null,
        source: '',
        error: String(e && e.message ? e.message : e)
      };
    }
    missing.push(deptKey);
  });

  if (missing.length) {
    var bulk = {};
    var bulkError = '';
    try {
      bulk = AggregatedCache_readMany(missing);
    } catch (e1) {
      bulkError = String(e1 && e1.message ? e1.message : e1);
      bulk = {};
    }
    missing.forEach(function(deptKey) {
      if (bulk && bulk[deptKey]) {
        result[deptKey] = { live: bulk[deptKey], source: 'aggregated_cache', error: '' };
        try { CacheLayer_write(deptKey, 'initData', bulk[deptKey], { persistToSheet: false }); } catch (e2) {}
      } else if (!result[deptKey]) {
        result[deptKey] = { live: null, source: '', error: bulkError };
      }
    });
  }

  return result;
}

function MrrDashboard_buildCurrentDeptMetric_(deptKey, diagnostics, currentCache, periodKey) {
  var validated = MrrDashboard_getValidatedCurrentInitData_(deptKey, currentCache, periodKey);
  if (!validated.ok) {
    MrrDashboard_addCurrentDiagnostic_(diagnostics, validated);
    return null;
  }

  var periodMetric = MrrDashboard_buildCurrentPeriodMetric_(validated.live, validated.member, validated.periodKey);
  MrrDashboard_addCurrentDiagnostic_(diagnostics, {
    deptKey: deptKey,
    status: 'ok',
    periodKey: validated.periodKey,
    repaired: !!validated.repaired,
    source: validated.source || '',
    membersCount: validated.membersCount
  });
  return MrrDashboard_buildMetricFromPeriodMetric_(periodMetric, validated.periodKey);
}

function MrrDashboard_getValidatedCurrentInitData_(deptKey, currentCache, periodKey) {
  var cached = MrrDashboard_readCurrentInitData_(deptKey, currentCache);
  var checked = MrrDashboard_validateCurrentInitData_(deptKey, cached.live, periodKey);
  checked.source = cached.source || '';
  if (!checked.ok && cached.error) checked.error = cached.error;
  return checked;
}

function MrrDashboard_readCurrentInitData_(deptKey, currentCache) {
  var cached = currentCache && currentCache[deptKey];
  return cached || { live: null, source: '', error: '' };
}

function MrrDashboard_getCurrentPeriodOptions_(deptKeys, currentCache) {
  for (var i = 0; i < (deptKeys || []).length; i++) {
    var cached = MrrDashboard_readCurrentInitData_(deptKeys[i], currentCache);
    var live = cached && cached.live;
    if (live && Array.isArray(live.periodOptions) && live.periodOptions.length) {
      return MrrDashboard_clonePeriodOptions_(live.periodOptions);
    }
  }
  return [];
}

function MrrDashboard_getSnapshotPeriodOptions_(fcstContext) {
  var monthSet = {};
  var quarterSet = {};
  var metrics = fcstContext && fcstContext.metricsByDeptDate || {};
  Object.keys(metrics).forEach(function(deptKey) {
    Object.keys(metrics[deptKey] || {}).forEach(function(dateStr) {
      ((metrics[deptKey][dateStr] && metrics[deptKey][dateStr].periodKeys) || []).forEach(function(periodKey) {
        periodKey = String(periodKey || '').trim();
        if (/^\d{4}-\d{2}$/.test(periodKey)) {
          monthSet[periodKey] = true;
        } else if (/^\d+Q[1-4]$/.test(periodKey)) {
          quarterSet[periodKey] = true;
        }
      });
    });
  });

  var monthKeys = Object.keys(monthSet).sort();
  var definitions = monthKeys.length ? FcstPeriods_buildDefinitionsFromMonthKeys_(monthKeys) : [];
  if (definitions.length) return definitions;

  return Object.keys(quarterSet)
    .sort(function(a, b) { return a < b ? -1 : a > b ? 1 : 0; })
    .map(function(key) { return FcstPeriods_getQuarterDefinitionByKey_(key); })
    .filter(function(option) { return !!option; });
}

function MrrDashboard_clonePeriodOptions_(periodOptions) {
  return (periodOptions || []).map(function(option) {
    return {
      key: String(option && option.key || ''),
      label: String(option && option.label || ''),
      months: (option && option.months || []).map(function(monthKey) {
        return String(monthKey || '');
      }).filter(function(monthKey) {
        return /^\d{4}-\d{2}$/.test(monthKey);
      })
    };
  }).filter(function(option) {
    return !!option.key;
  });
}

function MrrDashboard_resolvePeriodKey_(periodOptions, periodKey) {
  var requested = MrrDashboard_normalizePeriodKey_(periodKey);
  var validKeys = FcstPeriods_expandKeys_(periodOptions || []);
  if (requested) {
    return validKeys.indexOf(requested) !== -1 ? requested : '';
  }
  return MrrDashboard_getDefaultPeriodKey_(periodOptions);
}

function MrrDashboard_getDefaultPeriodKey_(periodOptions) {
  var todayMonthKey = MrrDashboard_getTodayMonthKey_();
  var currentQuarterKey = FcstPeriods_getQuarterKeyFromMonthKey_(todayMonthKey);
  if (MrrDashboard_findPeriodOption_(periodOptions, currentQuarterKey).key) {
    return currentQuarterKey;
  }
  return periodOptions && periodOptions[0] ? String(periodOptions[0].key || '') : '';
}

function MrrDashboard_getTodayMonthKey_() {
  var today = new Date();
  return Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM');
}

function MrrDashboard_getQuarterKeyForPeriod_(periodOptions, periodKey) {
  var key = MrrDashboard_normalizePeriodKey_(periodKey);
  for (var i = 0; i < (periodOptions || []).length; i++) {
    var option = periodOptions[i] || {};
    if (option.key === key || (option.months || []).indexOf(key) !== -1) {
      return String(option.key || '');
    }
  }
  return '';
}

function MrrDashboard_buildPeriodFilter_(periodOptions, periodKey) {
  var resolved = MrrDashboard_resolvePeriodKey_(periodOptions, periodKey);
  var months = {};
  if (/^\d{4}-\d{2}$/.test(resolved)) {
    months[resolved] = true;
  } else {
    (MrrDashboard_findPeriodOption_(periodOptions, resolved).months || []).forEach(function(monthKey) {
      months[monthKey] = true;
    });
  }
  return {
    periodKey: resolved,
    months: months
  };
}

function MrrDashboard_findPeriodOption_(periodOptions, periodKey) {
  var key = MrrDashboard_normalizePeriodKey_(periodKey);
  for (var i = 0; i < (periodOptions || []).length; i++) {
    if (String(periodOptions[i] && periodOptions[i].key || '') === key) return periodOptions[i] || {};
  }
  return {};
}

function MrrDashboard_periodFilterMatches_(filter, completedMonth) {
  var monthKey = MrrDashboard_normalizeCompletedMonthKey_(completedMonth);
  return !!(filter && filter.periodKey && filter.months && filter.months[monthKey]);
}

function MrrDashboard_normalizeCompletedMonthKey_(completedMonth) {
  if (completedMonth instanceof Date && !isNaN(completedMonth)) {
    return Utilities.formatDate(completedMonth, 'Asia/Tokyo', 'yyyy-MM');
  }
  var text = String(completedMonth || '').trim();
  var match = text.match(/^(\d{4})[-\/](\d{1,2})(?:[-\/]\d{1,2})?$/);
  if (!match) return text;
  return match[1] + '-' + String(Number(match[2])).padStart(2, '0');
}

function MrrDashboard_validateCurrentInitData_(deptKey, live, periodKey) {
  var result = {
    deptKey: deptKey,
    ok: false,
    status: 'error',
    reason: '',
    error: '',
    source: '',
    live: live || null,
    member: null,
    periodKey: '',
    membersCount: Array.isArray(live && live.members) ? live.members.length : 0
  };
  if (!live) {
    result.reason = 'empty_payload';
    return result;
  }
  if (live.error) {
    result.reason = 'source_error';
    result.error = String(live.error || '');
    return result;
  }
  if (!Array.isArray(live.members)) {
    result.reason = 'members_not_array';
    return result;
  }
  if (!Array.isArray(live.periodOptions)) {
    result.reason = 'period_options_not_array';
    return result;
  }

  var member = MrrDashboard_findCurrentDepartmentTotalMember_(deptKey, live.members || []);
  if (!member) {
    result.reason = 'department_total_missing';
    return result;
  }

  var resolvedPeriodKey = MrrDashboard_resolvePeriodKey_(live.periodOptions, periodKey);
  if (!resolvedPeriodKey) {
    result.reason = 'period_missing';
    return result;
  }
  if (!member[resolvedPeriodKey]) {
    result.reason = 'period_metric_missing';
    result.periodKey = resolvedPeriodKey;
    return result;
  }

  result.status = 'ok';
  result.ok = true;
  result.reason = '';
  result.member = member;
  result.periodKey = resolvedPeriodKey;
  return result;
}

function MrrDashboard_addCurrentDiagnostic_(diagnostics, item) {
  if (!diagnostics) return;
  if (!Array.isArray(diagnostics.current)) diagnostics.current = [];
  diagnostics.current.push({
    deptKey: String(item && item.deptKey || ''),
    status: String(item && item.status || 'error'),
    reason: String(item && item.reason || ''),
    error: String(item && item.error || ''),
    initialReason: String(item && item.initialReason || ''),
    periodKey: String(item && item.periodKey || ''),
    repaired: !!(item && item.repaired),
    source: String(item && item.source || ''),
    membersCount: Number(item && item.membersCount || 0) || 0
  });
}

function MrrDashboard_findCurrentDepartmentTotalMember_(deptKey, members) {
  for (var i = 0; i < (members || []).length; i++) {
    var member = members[i] || {};
    if (SharedAppState_isDepartmentTotal_(member, deptKey)) return member;
  }
  return null;
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

function MrrDashboard_getSnapshotPeriodFilterByDept_(deptKeys, snapshotDate, periodKey) {
  var dateSet = {};
  dateSet[snapshotDate] = true;
  var fcst = MrrDashboard_readFcstSnapshots_(deptKeys, dateSet);
  var periodOptions = MrrDashboard_getSnapshotPeriodOptions_(fcst);
  var requestedPeriodKey = MrrDashboard_normalizePeriodKey_(periodKey);
  var resolvedPeriodKey = MrrDashboard_resolvePeriodKey_(periodOptions, requestedPeriodKey);
  var result = {};
  (deptKeys || []).forEach(function(deptKey) {
    var snapshotEntry = fcst &&
      fcst.metricsByDeptDate &&
      fcst.metricsByDeptDate[deptKey] &&
      fcst.metricsByDeptDate[deptKey][snapshotDate];
    var deptPeriodKey = snapshotEntry && resolvedPeriodKey && snapshotEntry.periods[resolvedPeriodKey]
      ? resolvedPeriodKey
      : '';
    result[deptKey] = MrrDashboard_buildPeriodFilter_(periodOptions, deptPeriodKey);
  });
  return result;
}

function MrrDashboard_readOppSnapshotDeals_(deptKeys, snapshotDate, periodFilterByDept) {
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

    var targetFilter = periodFilterByDept && periodFilterByDept[deptKey];
    if (!MrrDashboard_periodFilterMatches_(targetFilter, legacyRow.completedMonth)) return;

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

function MrrDashboard_buildDeptMetricFromContext_(deptKey, snapshotDate, fcstContext, periodKey) {
  var snapshotEntry = fcstContext &&
    fcstContext.metricsByDeptDate &&
    fcstContext.metricsByDeptDate[deptKey] &&
    fcstContext.metricsByDeptDate[deptKey][snapshotDate];
  var resolvedPeriodKey = snapshotEntry && periodKey && snapshotEntry.periods[periodKey]
    ? periodKey
    : '';
  var periodMetric = resolvedPeriodKey && snapshotEntry ? snapshotEntry.periods[resolvedPeriodKey] : null;

  return MrrDashboard_buildMetricFromPeriodMetric_(periodMetric, resolvedPeriodKey);
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

function MrrDashboard_formatWeekLabel_(snapshotDate) {
  if (String(snapshotDate || '') === MRR_DASHBOARD_LIVE_KEY) return '現在';
  var match = String(snapshotDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(snapshotDate || '');
  return Number(match[2]) + '/' + Number(match[3]);
}

function MrrDashboard_getCurrentPeriodFilterByDept_(deptKeys, periodKey) {
  var currentCache = MrrDashboard_readCurrentCacheMap_(deptKeys);
  var result = {};
  (deptKeys || []).forEach(function(deptKey) {
    var validated = MrrDashboard_getValidatedCurrentInitData_(deptKey, currentCache, periodKey);
    if (!validated.ok) {
      throw new Error('MRR current FCST source failed: ' + deptKey + ' / ' + (validated.reason || validated.error || 'invalid payload'));
    }
    result[deptKey] = MrrDashboard_buildPeriodFilter_(validated.live.periodOptions, validated.periodKey);
  });
  return result;
}

function MrrDashboard_readCurrentOppDeals_(deptKeys, periodFilterByDept) {
  var deals = [];
  var sheetContexts = MrrDashboard_getCurrentOppSheetContexts_(deptKeys);
  sheetContexts.forEach(function(context) {
    var sheet = context.sheet;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 3 || lastCol < 1) return;

    var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    if (context.deptKeys.indexOf('SSSMBCS') !== -1) headers = normalizeSSCSHeaders_(headers);
    var headerMap = OppListReader_buildHeaderMap_(headers);
    var values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    var deptUserMaps = MrrDashboard_getDeptUserMaps_(context.deptKeys);
    var unionFilter = MrrDashboard_mergePeriodFilters_(periodFilterByDept, context.deptKeys);

    values.forEach(function(row) {
      if (MrrDashboard_isTrueBoolean_(OppListReader_valueByKeys_(row, headerMap, ['KeyDeal_最新'])) !== true) return;

      var resolvedDeptKey = MrrDashboard_resolveCurrentOppDeptKey_(row, headerMap, context.deptKeys, deptUserMaps);
      if (!resolvedDeptKey && !context.includeUnresolvedRows) return;

      var completedMonth = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['完了予定月']));
      var targetFilter = resolvedDeptKey && periodFilterByDept && periodFilterByDept[resolvedDeptKey]
        ? periodFilterByDept[resolvedDeptKey]
        : unionFilter;
      if (!MrrDashboard_periodFilterMatches_(targetFilter, completedMonth)) return;

      var oppId = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['ID', '案件ID'])).trim();
      var dealName = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['案件名'])).trim();
      if (!oppId || !dealName) return;
      var owner = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['サブオーナー', '担当者', 'ユーザー'])).trim();
      var sourceDept = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['担当部署'])).trim();

      deals.push({
        dept: resolvedDeptKey ? MrrDashboard_getDeptLabel_(resolvedDeptKey) : (sourceDept || context.division || ''),
        company: dealName,
        mrr: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['MRR', '受注MRR', '金額（LK＋新ソリューション）(換算値)', '月額(換算値)'])),
        phase: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['フェーズ_変換', 'フェーズ'])),
        yomi: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_最新', 'FCST(コミット)(換算値)'])),
        oppId: oppId,
        completedMonth: completedMonth,
        owner: owner
      });
    });
  });

  deals.sort(function(left, right) {
    return Math.abs(right.mrr) - Math.abs(left.mrr);
  });
  return deals;
}

function MrrDashboard_getCurrentOppSheetContexts_(deptKeys) {
  var bySheetName = {};
  (deptKeys || []).forEach(function(deptKey) {
    var cfg = DEPT_CONFIG[deptKey];
    if (!cfg || !cfg.sfSheet) return;
    if (!bySheetName[cfg.sfSheet]) {
      bySheetName[cfg.sfSheet] = {
        sheetName: cfg.sfSheet,
        sheet: getSharedSheet(cfg.sfSheet),
        deptKeys: [],
        division: cfg.division || ''
      };
    }
    bySheetName[cfg.sfSheet].deptKeys.push(deptKey);
  });

  return Object.keys(bySheetName).map(function(sheetName) {
    var context = bySheetName[sheetName];
    if (!context.sheet) throw new Error('MRR current Opp source sheet missing: ' + sheetName);
    context.includeUnresolvedRows = context.deptKeys.length > 1;
    return context;
  });
}

function MrrDashboard_getDeptUserMaps_(deptKeys) {
  var maps = {};
  (deptKeys || []).forEach(function(deptKey) {
    maps[deptKey] = {};
    OppListReader_getDeptUserNames_(deptKey).forEach(function(name) {
      name = String(name || '').trim();
      if (name) maps[deptKey][name] = true;
    });
  });
  return maps;
}

function MrrDashboard_resolveCurrentOppDeptKey_(row, headerMap, deptKeys, deptUserMaps) {
  var owner = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['サブオーナー', '担当者', 'ユーザー'])).trim();
  for (var i = 0; i < (deptKeys || []).length; i++) {
    var deptKey = deptKeys[i];
    if (owner && deptUserMaps && deptUserMaps[deptKey] && deptUserMaps[deptKey][owner]) return deptKey;
  }

  var sourceDept = OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['担当部署'])).trim();
  for (var j = 0; j < (deptKeys || []).length; j++) {
    var key = deptKeys[j];
    if (sourceDept === key || sourceDept.indexOf(key) === 0 || sourceDept === MrrDashboard_getDeptLabel_(key)) return key;
  }

  return '';
}

function MrrDashboard_mergePeriodFilters_(periodFilterByDept, deptKeys) {
  var merged = { periodKey: '', months: {} };
  (deptKeys || []).forEach(function(deptKey) {
    var filter = periodFilterByDept && periodFilterByDept[deptKey];
    if (!filter) return;
    if (!merged.periodKey) merged.periodKey = filter.periodKey || '';
    Object.keys(filter.months || {}).forEach(function(monthKey) {
      merged.months[monthKey] = true;
    });
  });
  return merged;
}

function MrrDashboard_isTrueBoolean_(value) {
  return value === true || String(value || '').trim().toUpperCase() === 'TRUE';
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
