var SHARED_TOTAL_KIND = {
  INDIVIDUAL: 'individual',
  GROUP: 'group',
  DEPARTMENT: 'department'
};

var SHARED_ALL_GROUP_LABEL = '部全体';

function SharedAppState_isDepartmentTotal_(member, deptKey) {
  if (!member || !member.isTotal) return false;
  if (member.totalKind) {
    if (member.totalKind !== SHARED_TOTAL_KIND.DEPARTMENT) return false;
    return !deptKey || String(member.dept || '') === String(deptKey);
  }
  var group = String(member.group || '');
  return group === SHARED_ALL_GROUP_LABEL || (!!deptKey && group === String(deptKey));
}

function SharedAppState_isGroupTotal_(member) {
  if (!member || !member.isTotal) return false;
  if (member.totalKind) return member.totalKind === SHARED_TOTAL_KIND.GROUP;
  return !SharedAppState_isDepartmentTotal_(member, member.dept || '');
}

function SharedAppState_getMemberGroupLabel_(member) {
  if (!member) return '';
  if (SharedAppState_isDepartmentTotal_(member, member.dept || '')) return SHARED_ALL_GROUP_LABEL;
  return String(member.group || member.groupCode || '');
}
