function SummaryReader_getSummaryData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) throw new Error('サマリシートが見つかりません');

  // 固定範囲で取得（getLastRow/getLastColumnのAPI呼び出しを省略）
  const values = sheet.getRange(1, 1, 60, 40).getValues();
  // 空行を末尾から除去
  let lastNonEmpty = values.length;
  while (lastNonEmpty > 0 && values[lastNonEmpty - 1].every(function(c) { return c === '' || c === null; })) {
    lastNonEmpty--;
  }
  const rows = values.slice(0, lastNonEmpty);
  return {
    headers: rows.slice(0, Math.min(3, rows.length)),
    rows: rows,
  };
}
