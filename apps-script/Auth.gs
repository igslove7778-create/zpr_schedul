/**
 * 대상자·알림·열람 규칙 (AUTH-01, AUTH-04, 6.1).
 */

/** 대상자 이름 목록. 우선순위: '전부' 체크 > 구성원 체크 열 > '대상자' 칸 텍스트 > 등록자 본인 */
function targetNames_(s) {
  var names = getMembers_().map(function (m) { return m.name; });
  if (s.all) return names;
  if (s.checked.length) return s.checked;
  if (s.target) {
    // 구성원 체크 열이 아직 없거나(구성원 열 동기화 전) 체크가 안 된 경우에도 '대상자' 칸 텍스트를 그대로 따른다
    var r = parseTargetNames_(s.target, names);
    if (r.targets[0] === '전부') return names;
    if (r.targets.length) return r.targets;
  }
  return s.registrant ? [s.registrant] : [];
}

/** 알림 대상 구성원 목록 (6.1 텔레그램 알림 열). 체크 없음 -> 등록자 본인 */
function notifyTargets_(s) {
  var names = targetNames_(s);
  return getMembers_().filter(function (m) { return names.indexOf(m.name) >= 0; });
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
  if (!s.checked.length && !s.target) return true;
  return targetNames_(s).indexOf(member.name) >= 0;
}

/** 대상자 표시 문자열 */
function targetLabel_(s) {
  if (s.all) return '전부';
  var names = targetNames_(s);
  return names.length ? names.join(', ') : (s.registrant || '본인');
}

/** 본인이 등록한 일정이거나 관리자면 수정·삭제 가능 (AUTH-06 + 통화 요구사항: 본인 일정만 수정 가능) */
function canEdit_(s, member) {
  if (!member) return false;
  if (member.role === '관리자') return true;
  return !!s.registrant && s.registrant === member.name;
}

/**
 * [일정] 시트 '대상자' 열의 텍스트를 체크 칸(전부/구성원 열)에 반영한다 (편의 기능: 이름만 적으면 자동 체크).
 * [구성원]에 없는 이름이 섞여 있으면 셀을 강조 표시하고 체크는 건드리지 않는다.
 * sh/hm 은 호출 측에서 이미 읽어둔 [일정] 시트와 헤더맵을 그대로 넘긴다.
 */
function applyTargetColumn_(sh, hm, row, s) {
  if (!hm[COL.TARGET]) return;
  var targetCell = sh.getRange(row, hm[COL.TARGET]);
  var text = (s.target || '').trim();
  if (!text) { targetCell.setBackground(null).clearNote(); return; }

  // 이름은 [일정] 시트 헤더(리마인드 오른쪽 열) + [구성원] 시트 기준으로 인식한다
  var names = checkColumnNames_(hm);
  var r = parseTargetNames_(text, names);
  if (r.unknown.length) {
    targetCell.setBackground('#F8CBAD').setNote('인식 못한 이름: ' + r.unknown.join(', ') + '\n([일정] 시트 이름 열 또는 [구성원] 시트에 없음)');
    log_('오류', s.id, CHANNEL.SHEET, s.registrant || getOperatorName_(), '실패', '대상자 인식 불가: ' + r.unknown.join(', '));
  } else {
    targetCell.setBackground(null).clearNote();
  }
  if (!r.targets.length) return;

  var allOn = r.targets.indexOf('전부') >= 0;
  if (hm[COL.ALL]) sh.getRange(row, hm[COL.ALL]).setValue(allOn ? CHECK_MARK : '');
  setChecks_(sh, hm, row, names, function (n) { return allOn || r.targets.indexOf(n) >= 0; });
}

/** 사람 체크 열에 표시/해제. on(name) 이 true 인 열만 표시 */
function setChecks_(sh, hm, row, names, on) {
  names.forEach(function (n) {
    if (!hm[n]) return;
    var cell = sh.getRange(row, hm[n]);
    var want = on(n);
    var has = isChecked_(cell.getValue());
    if (want && !has) cell.setValue(CHECK_MARK);
    else if (!want && has) cell.setValue('');
  });
}

/**
 * 사람 체크 열을 직접 체크/해제했을 때: 체크된 이름들을 '대상자' 칸에 써 준다.
 * 전원이 체크되면 '대상자'=전부, '전부' 칸도 체크. 아무도 없으면 '대상자' 비움.
 */
function syncTargetFromChecks_(sh, hm, row, s) {
  if (!hm[COL.TARGET]) return;
  var names = checkColumnNames_(hm);
  var values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  var checked = names.filter(function (n) { return hm[n] && isChecked_(values[hm[n] - 1]); });
  var allOn = names.length > 0 && checked.length === names.length;
  sh.getRange(row, hm[COL.TARGET]).setValue(allOn ? '전부' : checked.join(', ')).setBackground(null).clearNote();
  if (hm[COL.ALL]) {
    var allCell = sh.getRange(row, hm[COL.ALL]);
    if (allOn && !isChecked_(allCell.getValue())) allCell.setValue(CHECK_MARK);
    else if (!allOn && isChecked_(allCell.getValue())) allCell.setValue('');
  }
}

/** '전부' 칸을 직접 체크/해제했을 때: 체크면 전원 표시, 해제면 전원 해제 후 '대상자' 텍스트대로 다시 표시 */
function applyAllColumn_(sh, hm, row, s) {
  var names = checkColumnNames_(hm);
  if (s.all) { setChecks_(sh, hm, row, names, function () { return true; }); return; }
  setChecks_(sh, hm, row, names, function () { return false; });
  applyTargetColumn_(sh, hm, row, s);
}
