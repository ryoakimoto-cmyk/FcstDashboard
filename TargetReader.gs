function TargetReader_getTargets() {
  const sheet = TargetReader_getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};

  const values = sheet.getRange(1, 1, lastRow, 10).getValues();
  return values.reduce(function(map, row) {
    const date = TargetReader_parseDate_(row[0]);
    if (!date) return map;
    if (date.getFullYear() !== FCST_TARGET_YEAR) return map;

    const month = date.getMonth() + 1;
    if (FCST_TARGET_MONTHS.indexOf(month) === -1) return map;
    if (String(row[8] || '').trim() !== 'Net') return map;

    const orgName = String(row[6] || '').trim();
    if (!orgName) return map;

    map[orgName + '|' + TargetReader_formatYearMonth_(date)] = TargetReader_toNumber_(row[9]);
    return map;
  }, {});
}

function TargetReader_parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value || '').trim();
  if (!text || /期.*Q/.test(text)) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;

  const date = new Date(text + 'T00:00:00');
  if (isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function TargetReader_getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) throw new Error('目標シートが見つかりません');
  return sheet;
}

function TargetReader_formatYearMonth_(date) {
  return String(date.getFullYear()) + String(date.getMonth() + 1).padStart(2, '0');
}

function TargetReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}
