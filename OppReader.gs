function OppReader_getOpportunities() {
  const sheet = OppReader_getOppSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= OPP_HEADER_ROW) return [];

  const headers = sheet.getRange(OPP_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const headerMap = OppReader_buildHeaderMap_(headers);
  const values = sheet.getRange(OPP_HEADER_ROW + 1, 1, lastRow - OPP_HEADER_ROW, lastCol).getValues();

  return values.reduce(function(list, row, idx) {
    const dealName = OppReader_valueByKeys_(row, headerMap, ['dealName', '案件名']);
    if (!dealName) return list;

    list.push({
      rowIndex: OPP_HEADER_ROW + 1 + idx,
      targetMonth: OppReader_valueByKeys_(row, headerMap, ['targetMonth', '対予定月']),
      dept: OppReader_valueByKeys_(row, headerMap, ['dept', '事業部等']),
      type: OppReader_valueByKeys_(row, headerMap, ['type', '種類']),
      phase: OppReader_valueByKeys_(row, headerMap, ['phase', 'フェーズ']),
      forecast: OppReader_valueByKeys_(row, headerMap, ['forecast', 'フォーキャスト']),
      scheduledDate: OppReader_formatCell_(OppReader_valueByKeys_(row, headerMap, ['scheduledDate', '予定日', '受注予定日'])),
      dealName: dealName,
      mrr: OppReader_toNumber_(OppReader_valueByKeys_(row, headerMap, ['mrr', 'MRR'])),
      initialCost: OppReader_toNumber_(OppReader_valueByKeys_(row, headerMap, ['initialCost', '初期費用'])),
      keyDeal: OppReader_toBoolean_(OppReader_valueByKeys_(row, headerMap, ['keyDeal', 'Key Deal', 'KEY DEAL'])),
      fcstCommit: OppReader_pair_(row, headerMap, 'fcstCommit'),
      fcstMin: OppReader_pair_(row, headerMap, 'fcstMin'),
      fcstMax: OppReader_pair_(row, headerMap, 'fcstMax'),
      received: OppReader_pair_(row, headerMap, 'received'),
      debtMgmt: OppReader_pair_(row, headerMap, 'debtMgmt'),
      debtMgmtLite: OppReader_pair_(row, headerMap, 'debtMgmtLite'),
      expense: OppReader_pair_(row, headerMap, 'expense'),
      comment: OppReader_valueByKeys_(row, headerMap, ['comment', 'コメント']),
      notes: OppReader_valueByKeys_(row, headerMap, ['notes', '備考', 'メモ']),
    });
    return list;
  }, []);
}

function OppReader_getOppSheet_() {
  const sheets = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === OPP_SHEET_NAME) return sheets[i];
  }
  throw new Error('案件リストシートが見つかりません');
}

function OppReader_buildHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, idx) {
    const raw = String(header || '').trim();
    if (!raw) return;
    map[raw] = idx;
    map[OppReader_normalize_(raw)] = idx;
  });

  const pairs = FcstWriter_buildOppFieldMap_(headers);
  Object.keys(pairs).forEach(function(field) {
    map[field + '.sfNew'] = pairs[field] - 1;
    map[field + '.sfCurrent'] = pairs[field] - 2;
  });
  return map;
}

function OppReader_pair_(row, headerMap, field) {
  return {
    sfCurrent: OppReader_toNumber_(OppReader_valueByKeys_(row, headerMap, [field + '.sfCurrent'])),
    sfNew: OppReader_toNumber_(OppReader_valueByKeys_(row, headerMap, [field + '.sfNew'])),
  };
}

function OppReader_valueByKeys_(row, headerMap, keys) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const candidates = [key, OppReader_normalize_(key)];
    for (let j = 0; j < candidates.length; j++) {
      const idx = headerMap[candidates[j]];
      if (idx !== undefined) return row[idx];
    }
  }
  return '';
}

function OppReader_normalize_(text) {
  return String(text || '').replace(/\s+/g, '').replace(/[()（）]/g, '').toLowerCase();
}

function OppReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function OppReader_toBoolean_(value) {
  if (value === true || value === false) return value;
  return /^(true|yes|1|y|有|対象)$/i.test(String(value || '').trim());
}

function OppReader_formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value || '';
}
