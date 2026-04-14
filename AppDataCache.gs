function AppDataCache_getInitData(deptKey) {
  var cached = CacheLayer_read(deptKey, 'initData', { skipSharedSheet: true });
  if (cached) return cached;

  var aggregated = AggregatedCache_read(deptKey);
  if (aggregated) {
    CacheLayer_write(deptKey, 'initData', aggregated, { persistToSheet: false });
    return aggregated;
  }

  return AppDataCache_refreshInitData(deptKey);
}

function AppDataCache_refreshInitData(deptKey) {
  return AggregatedCache_refresh(deptKey);
}

function AppDataCache_getOpportunities(deptKey) {
  var cached = CacheLayer_read(deptKey, 'oppList', { skipSharedSheet: true });
  if (cached) return cached;
  return AppDataCache_refreshOpportunities(deptKey);
}

function AppDataCache_refreshOpportunities(deptKey) {
  var result = OppListReader_getLiveRows(deptKey);
  var latestSnapshot = null;

  result.snapshotDates = OppListSnapshot_getSnapshotDates(deptKey);
  if (result.snapshotDates.length) {
    latestSnapshot = OppListSnapshot_getByDate(deptKey, result.snapshotDates[0]);
  }
  result.previousRows = latestSnapshot ? (latestSnapshot.rows || []) : [];
  result.latestSnapshotData = latestSnapshot;

  CacheLayer_write(deptKey, 'oppList', result, { persistToSheet: false });
  return result;
}
