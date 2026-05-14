function AppDataCache_getInitData(deptKey, opts) {
  var scope = (opts && opts.scope) || 'fast';
  var cacheLayerKey = scope === 'fast' ? 'initData:fast' : 'initData';
  var aggregatedOpts = scope === 'fast' ? { scope: 'fast' } : undefined;

  var cached = CacheLayer_read(deptKey, cacheLayerKey, { skipSharedSheet: true });
  if (cached) return cached;

  var aggregated = AggregatedCache_read(deptKey, aggregatedOpts);
  if (aggregated) {
    AppDataCache_attachFcstKeyDeals_(deptKey, aggregated);
    CacheLayer_write(deptKey, cacheLayerKey, aggregated, { persistToSheet: false });
    return aggregated;
  }

  return AppDataCache_refreshInitData(deptKey, { scope: scope });
}

function AppDataCache_attachFcstKeyDeals_(deptKey, data) {
  FcstSnapshot_attachCurrentKeyDealsToData_(deptKey, data);
  return data;
}

function AppDataCache_refreshInitData(deptKey, opts) {
  var scope = (opts && opts.scope) || 'fast';
  var cacheLayerKey = scope === 'fast' ? 'initData:fast' : 'initData';
  var refreshOpts = scope === 'fast' ? { scope: 'fast' } : undefined;
  var result = AggregatedCache_refresh(deptKey, refreshOpts);
  AppDataCache_attachFcstKeyDeals_(deptKey, result);
  CacheLayer_write(deptKey, cacheLayerKey, result, { persistToSheet: false });
  return result;
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
