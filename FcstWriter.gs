function FcstWriter_saveFcstAdjusted(params) {
  const snapIdx = Number(params.snapIdx);
  const memberName = String(params.memberName || '').trim();
  const block = String(params.block || '');
  if (!memberName || !BLOCK_COL_STARTS[block]) throw new Error('保存パラメータが不正です');

  const sheet = FcstReader_getFcstSheet_();
  const startRow = SNAPSHOT_START_ROW + snapIdx * ROWS_PER_SNAPSHOT;
  const names = sheet.getRange(startRow, 2, ROWS_PER_SNAPSHOT, 1).getValues();
  let rowIndex = -1;

  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0] || '').trim() === memberName) {
      rowIndex = startRow + i;
      break;
    }
  }
  if (rowIndex === -1) throw new Error('対象メンバーが見つかりません');

  const startCol = BLOCK_COL_STARTS[block];
  sheet.getRange(rowIndex, startCol + BLOCK_OFFSETS.fcstAdjusted.net - 1).setValue(FcstWriter_toYen_(params.net));
  sheet.getRange(rowIndex, startCol + BLOCK_OFFSETS.fcstAdjusted.newExp - 1).setValue(FcstWriter_toYen_(params.newExp));
  sheet.getRange(rowIndex, startCol + BLOCK_OFFSETS.fcstAdjusted.churn - 1).setValue(FcstWriter_toYen_(params.churn));
  return { success: true };
}

function FcstWriter_saveOppSfValue(params) {
  const field = String(params.field || '');
  const rowIndex = Number(params.rowIndex);
  if (!field || !rowIndex) throw new Error('保存パラメータが不正です');

  const sheet = OppReader_getOppSheet_();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(OPP_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const fieldMap = FcstWriter_buildOppFieldMap_(headers);
  const targetCol = fieldMap[field];
  if (!targetCol) throw new Error('SF更新値列が見つかりません: ' + field);

  sheet.getRange(rowIndex, targetCol).setValue(FcstWriter_toYen_(params.value));
  return { success: true };
}

function FcstWriter_buildOppFieldMap_(headers) {
  const parents = {
    fcstCommit: 'FCST(コミット)',
    fcstMin: 'FCST(MIN)',
    fcstMax: 'FCST(MAX)',
    received: '受注',
    debtMgmt: '債務管理',
    debtMgmtLite: '債務管理Lite',
    expense: '費用',
  };
  const map = {};

  Object.keys(parents).forEach(function(field) {
    const parent = parents[field];
    for (let i = 0; i < headers.length - 1; i++) {
      const current = String(headers[i] || '').replace(/\s+/g, '');
      const next = String(headers[i + 1] || '').replace(/\s+/g, '');
      if (current.indexOf(parent.replace(/\s+/g, '')) !== -1 && next.indexOf('SF更新値') !== -1) {
        map[field] = i + 2;
        break;
      }
    }
  });
  return map;
}

function FcstWriter_toYen_(value) {
  return Math.round((Number(value) || 0) * 10000);
}
