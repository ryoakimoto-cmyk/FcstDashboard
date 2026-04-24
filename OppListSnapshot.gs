function OppListSnapshot_createWeekly(deptKey) {
  return OppHistory_createWeekly_(deptKey);
}

function OppListSnapshot_getSnapshotDates(deptKey) {
  return OppHistory_getSnapshotDates_(deptKey);
}

function OppListSnapshot_getByDate(deptKey, dateStr) {
  return OppHistory_getByDate_(deptKey, dateStr);
}

function OppListSnapshot_setupWeeklyTrigger() {
  OppHistory_ensureInfrastructure_();
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'createOppSnapshot') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('createOppSnapshot')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();
  return { ok: true };
}
