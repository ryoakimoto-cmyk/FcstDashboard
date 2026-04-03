function NotesSheet_getNotes() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(NOTES_SHEET_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};
  var values = sheet.getRange(1, 1, lastRow, 3).getValues();
  return values.reduce(function(map, row) {
    var name = String(row[0] || '').trim();
    var period = String(row[1] || '').trim();
    var note = String(row[2] || '');
    if (name && period) map[name + '|' + period] = note;
    return map;
  }, {});
}

function NotesSheet_saveNote(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(NOTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTES_SHEET_NAME);
  }
  var lastRow = sheet.getLastRow();
  var now = new Date();
  if (lastRow > 0) {
    var data = sheet.getRange(1, 1, lastRow, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === p.name && String(data[i][1]).trim() === p.period) {
        sheet.getRange(i + 1, 3).setValue(p.note);
        sheet.getRange(i + 1, 4).setValue(now);
        return { ok: true };
      }
    }
  }
  sheet.appendRow([p.name, p.period, p.note, now]);
  return { ok: true };
}
