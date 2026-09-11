/**
 * 대상자·알림·열람 규칙 (AUTH-01, AUTH-04, 6.1).
 */

/** 알림 대상 구성원 목록 (6.1 텔레그램 알림 열). 체크 없음 -> 등록자 본인 */
function notifyTargets_(s) {
  var members = getMembers_();
  if (s.all) return members;
  if (s.checked.length) return members.filter(function (m) { return s.checked.indexOf(m.name) >= 0; });
  return members.filter(function (m) { return m.name === s.registrant; });
}

/**
 * 열람 가능 여부 (텔레그램·웹용. 시트에는 미적용 - D-04 확정).
 * D-03 기본값 안1: 체크 없는 일정은 전 구성원 열람.
 */
function canView_(s, member) {
  if (!member) return false;
  if (member.role === '관리자') return true;
  if (s.all) return true;
  if (s.registrant === member.name) return true;
  if (s.checked.length === 0) return true;
  return s.checked.indexOf(member.name) >= 0;
}

/** 대상자 표시 문자열 */
function targetLabel_(s) {
  if (s.all) return '전부';
  if (s.checked.length) return s.checked.join(', ');
  return s.registrant || '본인';
}
