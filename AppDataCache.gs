function AppDataCache_getInitData(deptKey) {
  return AppDataCache_getData_(deptKey, 'initData');
}

function AppDataCache_refreshInitData(deptKey) {
  return AppDataCache_refreshData_(deptKey, 'initData');
}

function AppDataCache_getOpportunities(deptKey) {
  return AppDataCache_getData_(deptKey, 'oppList');
}

function AppDataCache_refreshOpportunities(deptKey) {
  return AppDataCache_refreshData_(deptKey, 'oppList');
}

function AppDataCache_getData_(deptKey, dataKey) {
  var profile = AppDataCache_getProfile_(dataKey);
  var cached = CacheLayer_read(deptKey, dataKey, { skipSharedSheet: true });
  if (cached) return cached;

  if (profile.readAggregatedSnapshot) {
    var aggregated = AggregatedCache_read(deptKey);
    if (aggregated) {
      CacheLayer_write(deptKey, dataKey, aggregated, { persistToSheet: false });
      return aggregated;
    }
  }

  return AppDataCache_refreshData_(deptKey, dataKey);
}

function AppDataCache_refreshData_(deptKey, dataKey) {
  var profile = AppDataCache_getProfile_(dataKey);
  var result = dataKey === 'initData'
    ? AppDataCache_refreshInitDataLive_(deptKey)
    : AppDataCache_refreshOpportunitiesLive_(deptKey);

  if (profile.cacheAfterRefresh) {
    CacheLayer_write(deptKey, dataKey, result, { persistToSheet: profile.persistToSheet });
  }
  return result;
}

function AppDataCache_warmDept_(deptKey) {
  var warmed = { deptKey: deptKey };
  try {
    warmed.initData = AppDataCache_refreshData_(deptKey, 'initData');
  } catch (e) {
    warmed.initDataError = String(e && e.message ? e.message : e);
  }
  try {
    warmed.oppList = AppDataCache_refreshData_(deptKey, 'oppList');
  } catch (e) {
    warmed.oppListError = String(e && e.message ? e.message : e);
  }
  return warmed;
}

function AppDataCache_getProfile_(dataKey) {
  if (dataKey === 'initData') {
    return {
      readAggregatedSnapshot: true,
      cacheAfterRefresh: false,
      persistToSheet: false
    };
  }
  if (dataKey === 'oppList') {
    return {
      readAggregatedSnapshot: false,
      cacheAfterRefresh: true,
      persistToSheet: false
    };
  }
  throw new Error('Unsupported shared cache profile: ' + dataKey);
}

function AppDataCache_refreshInitDataLive_(deptKey) {
  return AggregatedCache_refresh(deptKey);
}

function AppDataCache_refreshOpportunitiesLive_(deptKey) {
  var result = OppListReader_getLiveRows(deptKey);
  var latestSnapshot = null;

  result.snapshotDates = OppListSnapshot_getSnapshotDates(deptKey);
  if (result.snapshotDates.length) {
    latestSnapshot = OppListSnapshot_getByDate(deptKey, result.snapshotDates[0]);
  }
  result.previousRows = latestSnapshot ? (latestSnapshot.rows || []) : [];
  result.latestSnapshotData = latestSnapshot;
  return result;
}
