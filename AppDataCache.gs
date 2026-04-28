function AppDataCache_getInitData(deptKey) {
  var cached = CacheLayer_read(deptKey, 'initData', { skipSharedSheet: true });
  if (cached) return AppDataCache_attachFcstKeyDeals_(deptKey, cached);

  var aggregated = AggregatedCache_read(deptKey);
  if (aggregated) {
    AppDataCache_attachFcstKeyDeals_(deptKey, aggregated);
    CacheLayer_write(deptKey, 'initData', aggregated, { persistToSheet: false });
    return aggregated;
  }

  return AppDataCache_refreshInitData(deptKey);
}

function AppDataCache_attachFcstKeyDeals_(deptKey, data) {
  FcstSnapshot_attachCurrentKeyDealsToData_(deptKey, data);
  if (data && data.latestSnapshotData && data.latestSnapshotData.date) {
    FcstSnapshot_attachSnapshotKeyDealsToData_(deptKey, data.latestSnapshotData.date, data.latestSnapshotData);
  }
  if (data && data.previousSnapshot && data.previousSnapshot.date) {
    FcstSnapshot_attachSnapshotKeyDealsToData_(deptKey, data.previousSnapshot.date, data.previousSnapshot);
  }
  return data;
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
