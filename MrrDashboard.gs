var MRR_SHEET_ID = '1dEITXn1wafXDjwtGFSUpAqohnjVLf7dBi888F7OrRL4';

function mrrDashboard_doGet_() {
  return HtmlService.createHtmlOutputFromFile('mrr-index')
    .setTitle('MRR進捗ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMrrDashboardData(division) {
  var normalizedDivision = String(division || 'SS').toUpperCase();
  if (normalizedDivision === 'BO') {
    return MrrDashboard_getBoData_();
  }
  return MrrDashboard_getSsData_();
}

function MrrDashboard_getSsData_() {
  var ss = SpreadsheetApp.openById(MRR_SHEET_ID);
  var sheet = ss.getSheets()[0];
  var values = sheet.getDataRange().getValues();
  var dataRows = values.slice(1);

  var weeksOrder = [];
  var weeksSeen = {};
  var weekLabels = {};
  var deptsOrder = [];
  var deptsSeen = {};
  var data = {};

  var currentMonth = null;
  var currentWeek = null;

  dataRows.forEach(function(row) {
    var month = row[0];
    var week = row[1];
    var dept = String(row[2] || '').trim();
    var target = Number(row[3]) || 0;
    var actual = Number(row[4]) || 0;
    var expectedMrr = Number(row[5]) || 0;
    var fcst = Number(row[6]) || 0;
    var keyDeal = String(row[7] || '');

    if (month !== '' && month !== null) currentMonth = month;
    if (week !== '' && week !== null) currentWeek = week;

    if (!currentMonth || !currentWeek || !dept) return;

    var weekKey = String(currentMonth) + '月W' + String(currentWeek);
    if (!weeksSeen[weekKey]) {
      weeksSeen[weekKey] = true;
      weeksOrder.push(weekKey);
      weekLabels[weekKey] = weekKey;
    }
    if (!data[weekKey]) data[weekKey] = {};

    if (dept !== 'SS' && !deptsSeen[dept]) {
      deptsSeen[dept] = true;
      deptsOrder.push(dept);
    }

    data[weekKey][dept] = {
      target: Math.round(target),
      actual: Math.round(actual),
      expectedMrr: Math.round(expectedMrr),
      fcst: Math.round(fcst),
      keyDeal: keyDeal,
      keyDealsData: []
    };
  });

  return {
    division: 'SS',
    totalDeptKey: 'SS',
    allLabel: '全事業部',
    weeks: weeksOrder,
    weekLabels: weekLabels,
    depts: deptsOrder,
    data: data
  };
}
