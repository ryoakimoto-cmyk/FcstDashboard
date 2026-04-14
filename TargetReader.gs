function TargetReader_getTargets(deptKey) {
  var sheet = TargetReader_getSheet_(deptKey);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};

  var values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 11)).getValues();
  var filtered = values.filter(function(row) {
    return String(row[10] || '').trim() === deptKey;
  });
  if (!filtered.length) return {};

  return filtered.reduce(function(map, row) {
    var date = TargetReader_parseDate_(row[0]);
    if (!date) return map;
    if (!FcstPeriods_isSupportedDate_(date)) return map;
    if (String(row[8] || '').trim() !== 'Net') return map;

    var orgName = String(row[6] || '').trim();
    if (!orgName) return map;

    map[orgName + '|' + TargetReader_formatYearMonth_(date)] = TargetReader_toNumber_(row[9]);
    return map;
  }, {});
}

function TargetReader_parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  var text = String(value || '').trim();
  if (!text || /.*Q/.test(text)) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;

  var date = new Date(text + 'T00:00:00');
  if (isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function TargetReader_getSheet_(deptKey) {
  return getSharedSheet(TARGET_SHEET_NAME);
}

function TargetReader_formatYearMonth_(date) {
  return String(date.getFullYear()) + String(date.getMonth() + 1).padStart(2, '0');
}

function TargetReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}
