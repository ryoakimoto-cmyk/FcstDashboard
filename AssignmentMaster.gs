function AssignmentMaster_getContext(deptKey) {
  return AssignmentMaster_normalizeContext_(deptKey, MonthlyMasterReader_getContext(deptKey));
}

function AssignmentMaster_normalizeContext_(deptKey, rawContext) {
  var context = rawContext || {};
  return {
    users: AssignmentMaster_normalizeUserMap_(deptKey, context.users || {}),
    monthlyUsers: AssignmentMaster_normalizeUserMap_(deptKey, context.monthlyUsers || {}),
    targets: context.targets || {}
  };
}

function AssignmentMaster_normalizeUserMap_(deptKey, users) {
  var normalized = {};
  Object.keys(users || {}).forEach(function(key) {
    normalized[key] = AssignmentMaster_normalizeUser_(deptKey, users[key]);
  });
  return normalized;
}

function AssignmentMaster_normalizeUser_(deptKey, user) {
  var source = user || {};
  var groupCode = String(source.groupCode || source.group || '').trim();
  var groupLabel = String(source.groupLabel || source.group || groupCode).trim();
  var displayName = String(source.displayName || source.name || source.sourceUserName || '').trim();
  return {
    sourceUserName: String(source.sourceUserName || '').trim(),
    displayName: displayName,
    dept: String(source.dept || deptKey || '').trim(),
    group: groupLabel,
    groupCode: groupCode,
    sortOrder: Number(source.sortOrder || 0) || 0
  };
}
