var MRR_SHEET_ID = '1dEITXn1wafXDjwtGFSUpAqohnjVLf7dBi888F7OrRL4';

function mrrDashboard_doGet_() {
  return HtmlService.createHtmlOutputFromFile('mrr-index')
    .setTitle('MRR進捗ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMrrDashboardData() {
  var ss = SpreadsheetApp.openById(MRR_SHEET_ID);
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();

  // Row 0 = header, skip it
  var dataRows = values.slice(1);

  var weeksOrder = [];
  var weeksSeen = {};
  var deptsOrder = [];
  var deptsSeen = {};
  var data = {};

  var currentMonth = null;
  var currentWeek = null;

  dataRows.forEach(function(row) {
    var month = row[0]; // Col A: 月
    var week  = row[1]; // Col B: 週
    var dept  = String(row[2] || '').trim(); // Col C: 部署
    var target      = Number(row[3]) || 0;  // Col D: 目標
    var actual      = Number(row[4]) || 0;  // Col E: 実績
    var expectedMrr = Number(row[5]) || 0;  // Col F: 期待MRR
    var fcst        = Number(row[6]) || 0;  // Col G: FCST
    var keyDeal     = String(row[7] || ''); // Col H: Key Deal

    // Fill-down: SS rows have month/week, sub-rows don't
    if (month !== '' && month !== null) currentMonth = month;
    if (week  !== '' && week  !== null) currentWeek  = week;

    if (!currentMonth || !currentWeek || !dept) return;

    var weekKey = currentMonth + '月W' + currentWeek;

    if (!weeksSeen[weekKey]) {
      weeksSeen[weekKey] = true;
      weeksOrder.push(weekKey);
    }
    if (!data[weekKey]) data[weekKey] = {};

    // Track non-SS departments
    if (dept !== 'SS' && !deptsSeen[dept]) {
      deptsSeen[dept] = true;
      deptsOrder.push(dept);
    }

    data[weekKey][dept] = {
      target:      Math.round(target),
      actual:      Math.round(actual),
      expectedMrr: Math.round(expectedMrr),
      fcst:        Math.round(fcst),
      keyDeal:     keyDeal
    };
  });

  return {
    weeks: weeksOrder,
    depts: deptsOrder,
    data:  data
  };
}
