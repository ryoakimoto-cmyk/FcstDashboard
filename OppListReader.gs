function OppListReader_getLiveRows() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SF_DATA_SHEET_NAME);
  if (!sheet) throw new Error('SFデータ更新シートが見つかりません');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return { rows: [], lastUpdated: '' };

  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var headerMap = OppListReader_buildHeaderMap_(headers);
  var values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  var rows = [];

  values.forEach(function(row) {
    var oppId = OppListReader_valueByKeys_(row, headerMap, ['ID', '案件 ID']);
    var dealName = OppListReader_valueByKeys_(row, headerMap, ['案件名']);
    if (!oppId || !dealName) return;

    oppId = String(oppId).trim();

    var proposalIds = {
      received: OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_受領']),
      debtMgmt: OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_債権管理']),
      debtMgmtLite: OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_債権管理 Lite']),
      expense: OppListReader_valueByKeys_(row, headerMap, ['提案商品ID_経費'])
    };
    Object.keys(proposalIds).forEach(function(key) {
      proposalIds[key] = String(proposalIds[key] || '').trim();
    });

    rows.push({
      oppId: oppId,
      completedMonth: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['完了予定月'])),
      dept: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['担当部署'])),
      type: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['種別'])),
      subOwner: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['ユーザー'])),
      phase: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['フェーズ_変換'])),
      forecast: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['フォーキャスト_変換'])),
      scheduleOrCloseDate: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['予定日 / 確定日'])),
      dealName: OppListReader_formatCell_(dealName),
      allocationPercent: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['パーセント (%)'])),
      mrr: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['月額(換算値)'])),
      initialCost: OppListReader_toNumber_(OppListReader_valueByKeys_(row, headerMap, ['初期費用額(換算値)'])),
      keyDeal: OppListReader_toBoolean_(OppListReader_valueByKeys_(row, headerMap, ['KeyDeal_最新'])),
      fcstCommit: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_最新'])
      ),
      fcstMin: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(MIN)_最新'])
      ),
      fcstMax: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(MAX)_最新'])
      ),
      received: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_受領'])
      ),
      debtMgmt: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_債権管理'])
      ),
      debtMgmtLite: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_債権管理 Lite'])
      ),
      expense: OppListReader_buildAmountPair_(
        OppListReader_valueByKeys_(row, headerMap, ['FCST(コミット)_経費'])
      ),
      proposalProductIds: proposalIds,
      fcstComment: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['FCSTコメント_最新'])),
      firstMeetingDate: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['初回営業日 / ｺﾝﾀｸﾄ日'])),
      summary: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['案件概要'])),
      optionFeatures: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['オプション機能'])),
      initiative: OppListReader_formatCell_(OppListReader_valueByKeys_(row, headerMap, ['施策']))
    });
  });

  return {
    lastUpdated: OppListReader_extractLastUpdated_(sheet.getRange(1, 1).getDisplayValue()),
    rows: rows
  };
}

function OppListReader_buildHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(header, idx) {
    var raw = String(header || '').trim();
    if (!raw) return;
    map[raw] = idx;
    map[OppListReader_normalize_(raw)] = idx;
  });
  return map;
}

function OppListReader_valueByKeys_(row, headerMap, keys) {
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var candidates = [key, OppListReader_normalize_(key)];
    for (var j = 0; j < candidates.length; j++) {
      var idx = headerMap[candidates[j]];
      if (idx !== undefined) return row[idx];
    }
  }
  return '';
}

function OppListReader_buildAmountPair_(sfValue) {
  var sfNum = OppListReader_toNumber_(sfValue);
  return { sfValue: sfNum, draftValue: sfNum };
}

function OppListReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function OppListReader_toBoolean_(value) {
  var v = value;
  if (v === true || v === false) return v;
  return /^(true|yes|1|y|済|true)$/i.test(String(v || '').trim());
}

function OppListReader_formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value === null || value === undefined ? '' : String(value);
}

function OppListReader_normalize_(text) {
  return String(text || '').replace(/\s+/g, '').replace(/[()（）]/g, '').toLowerCase();
}

function OppListReader_extractLastUpdated_(title) {
  var text = String(title || '').trim();
  var match = text.match(/\d{4}[\/-]\d{1,2}[\/-]\d{1,2}/);
  if (!match) return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var normalized = match[0].replace(/\//g, '-');
  var date = new Date(normalized + 'T00:00:00');
  if (isNaN(date)) return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}
