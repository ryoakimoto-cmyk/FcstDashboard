function UserReader_getUsers() {
  const sheet = UserReader_getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return {};

  const values = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
  return values.reduce(function(map, row) {
    const name = String(row[3] || '').trim();
    if (!name) return map;

    map[name] = {
      group: String(row[2] || '').trim(),
      dept: String(row[1] || '').trim(),
      sortOrder: UserReader_toNumber_(row[5]),
    };
    return map;
  }, {});
}

function UserReader_getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SF_USER_SHEET_NAME);
  if (!sheet) throw new Error('SFユーザーシートが見つかりません');
  return sheet;
}

function UserReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}
