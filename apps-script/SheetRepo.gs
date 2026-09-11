/**
 * 구글시트 원장 접근 (SYNC-01, 3장). 열은 문자가 아닌 헤더명으로 참조한다.
 */

function ss_() { return SpreadsheetApp.getActive(); }

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name + ' (메뉴 > 일정관리 > 초기 설정 실행)');
  return sh;
}

/** 헤더명 -> 1-based 열 번호 */
function headerMap_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function (h, i) { var k = String(h).trim(); if (k) map[k] = i + 1; });
  map.__headers = headers.map(function (h) { return String(h).trim(); });
  return map;
}

// ---------- 구성원 ----------

/** 사용여부=Y 인 구성원 목록 [{name,email,tgId,role,active,code,row}] */
function getMembers_(includeInactive) {
  var sh = sheet_(SHEET.MEMBER);
  if (sh.getLastRow() < 2) return [];
  var hm = headerMap_(sh);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var out = [];
  rows.forEach(function (r, i) {
    var name = String(r[hm[MEMBER_COL.NAME] - 1] || '').trim();
    if (!name) return;
    var active = String(r[hm[MEMBER_COL.ACTIVE] - 1] || 'Y').trim().toUpperCase() !== 'N';
    if (!active && !includeInactive) return;
    out.push({
      name: name,
      email: String(r[hm[MEMBER_COL.EMAIL] - 1] || '').trim(),
      tgId: String(r[hm[MEMBER_COL.TG_ID] - 1] || '').trim(),
      role: String(r[hm[MEMBER_COL.ROLE] - 1] || '일반').trim(),
      active: active,
      code: String(r[hm[MEMBER_COL.CODE] - 1] || '').trim(),
      calId: hm[MEMBER_COL.CAL_ID] ? String(r[hm[MEMBER_COL.CAL_ID] - 1] || '').trim() : '',
      row: i + 2
    });
  });
  return out;
}

function findMemberByTgId_(tgId) {
  var id = String(tgId);
  return getMembers_().filter(function (m) { return m.tgId === id; })[0] || null;
}

function findMemberByName_(name) {
  return getMembers_(true).filter(function (m) { return m.name === name; })[0] || null;
}

function getAdmins_() {
  return getMembers_().filter(function (m) { return m.role === '관리자'; });
}

/** 시트 직접 입력 건의 등록자 (설정 '담당자명' 또는 첫 관리자) */
function getOperatorName_() {
  var s = getSettings_()['담당자명'];
  if (s) return s;
  var a = getAdmins_()[0];
  return a ? a.name : '담당자';
}

/**
 * 시트를 편집한 사람의 이름. onEdit 이벤트의 편집자 이메일을 [구성원] '구글 계정' 과 대조한다.
 * 개인 Gmail 은 구글 보안 정책상 편집자 이메일이 비어 올 수 있다 -> 그때는 getOperatorName_() 로 대체.
 */
function getEditorName_(e) {
  var email = '';
  try { if (e && e.user) email = String(e.user.getEmail() || ''); } catch (err) { /* 권한 없음 */ }
  // 주의: Session.getActiveUser() 는 설치형 트리거 안에서 편집자가 아니라 트리거 설치자를 돌려줄 수 있어 쓰지 않는다
  email = email.trim().toLowerCase();
  editorDiag_(e, email);
  if (email) {
    var m = getMembers_(true).filter(function (x) { return x.email.toLowerCase() === email; })[0];
    if (m) return m.name;
  }
  return getOperatorName_();
}

/** (진단용) 편집 이벤트에서 구글이 넘겨주는 사용자 정보를 [로그]에 남긴다. 원인 확인 후 제거 예정 */
function editorDiag_(e, email) {
  var a = '', f = '';
  try { a = Session.getActiveUser().getEmail(); } catch (x) { a = 'ERR'; }
  try { f = Session.getEffectiveUser().getEmail(); } catch (y) { f = 'ERR'; }
  log_('진단', '', CHANNEL.SHEET, '시스템', '정보', 'e.user=' + (email || '(빈값)') + ' / active=' + (a || '(빈값)') + ' / effective(트리거소유자)=' + (f || '(빈값)'));
}

// ---------- 일정 ----------

/** 일정 시트의 행 하나를 객체로 */
function rowToSchedule_(hm, rowIdx, values, members) {
  var get = function (col) { var c = hm[col]; return c ? values[c - 1] : ''; };
  var dateVal = get(COL.DATE);
  var date = '';
  if (dateVal instanceof Date) date = Utilities.formatDate(dateVal, 'Asia/Seoul', 'yyyy-MM-dd');
  else if (dateVal) date = String(dateVal).trim();
  var timeVal = get(COL.TIME);
  var time = '';
  if (timeVal instanceof Date) time = Utilities.formatDate(timeVal, 'Asia/Seoul', 'HH:mm');
  else if (timeVal !== '' && timeVal !== null) time = String(timeVal).trim();
  var checked = [];
  members.forEach(function (n) { if (hm[n] && isChecked_(values[hm[n] - 1])) checked.push(n); });
  return {
    row: rowIdx,
    date: date, time: time,
    title: String(get(COL.TITLE) || '').trim(),
    memo: String(get(COL.MEMO) || '').trim(),
    target: String(get(COL.TARGET) || '').trim(),
    alarm: String(get(COL.ALARM) || '').trim(),
    all: isChecked_(get(COL.ALL)),
    checked: checked,
    id: String(get(COL.ID) || '').trim(),
    registrant: String(get(COL.REGISTRANT) || '').trim(),
    channel: String(get(COL.CHANNEL) || '').trim(),
    calId: String(get(COL.CAL_ID) || '').trim(),
    updated: String(get(COL.UPDATED) || '').trim(),
    status: String(get(COL.STATUS) || '').trim(),
    remind: String(get(COL.REMIND) || '').trim()
  };
}

function memberNames_() {
  return getMembers_(true).map(function (m) { return m.name; });
}

/**
 * [일정] 시트의 사람 체크 열 이름 목록.
 * '리마인드'(마지막 시스템 열) 오른쪽에 있는 헤더는 전부 사람 이름으로 본다 (시트에 직접 추가한 열 포함)
 * + [구성원] 시트 이름. 순서는 시트 헤더 순.
 */
function checkColumnNames_(hm) {
  var headers = hm.__headers || [];
  var start = hm[COL.REMIND] || hm[COL.ID] || 0; // 1-based; 이 열 오른쪽부터
  var known = {};
  Object.keys(COL).forEach(function (k) { known[COL[k]] = true; });
  var out = [];
  headers.forEach(function (h, i) {
    if (i + 1 <= start || !h || known[h]) return;
    if (out.indexOf(h) < 0) out.push(h);
  });
  memberNames_().forEach(function (n) { if (hm[n] && out.indexOf(n) < 0) out.push(n); });
  return out;
}

function getSchedule_(rowIdx) {
  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  var values = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
  return rowToSchedule_(hm, rowIdx, values, checkColumnNames_(hm));
}

function getAllSchedules_() {
  var sh = sheet_(SHEET.SCHEDULE);
  if (sh.getLastRow() < 2) return [];
  var hm = headerMap_(sh);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var members = checkColumnNames_(hm);
  return rows.map(function (r, i) { return rowToSchedule_(hm, i + 2, r, members); })
    .filter(function (s) { return s.date || s.title; });
}

function setCell_(sh, hm, rowIdx, col, value) {
  if (!hm[col]) throw new Error('열이 없습니다: ' + col);
  sh.getRange(rowIdx, hm[col]).setValue(value);
}

/**
 * 일정 행 추가 (텔레그램·웹 등록용).
 * s = { date:'yyyy-MM-dd', time:'HH:mm'|'', title, memo, targets:['전부']|[이름...], registrant, channel }
 * 반환 { id, row }
 */
function appendSchedule_(s) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SHEET.SCHEDULE);
    var hm = headerMap_(sh);
    var row = new Array(sh.getLastColumn()).fill('');
    var id = newId_();
    var p = s.date.split('-').map(Number);
    row[hm[COL.DATE] - 1] = new Date(p[0], p[1] - 1, p[2]);
    row[hm[COL.TIME] - 1] = s.time || '';
    row[hm[COL.TITLE] - 1] = s.title;
    row[hm[COL.MEMO] - 1] = s.memo || '';
    row[hm[COL.ALARM] - 1] = ALARM.PENDING;
    if (hm[COL.TARGET]) row[hm[COL.TARGET] - 1] = s.targets.indexOf('전부') >= 0 ? '전체' : s.targets.join(' ');
    if (s.targets.indexOf('전부') >= 0) row[hm[COL.ALL] - 1] = 'o';
    else s.targets.forEach(function (n) { if (hm[n]) row[hm[n] - 1] = 'o'; });
    row[hm[COL.ID] - 1] = id;
    row[hm[COL.REGISTRANT] - 1] = s.registrant;
    row[hm[COL.CHANNEL] - 1] = s.channel;
    row[hm[COL.UPDATED] - 1] = nowString_();
    row[hm[COL.STATUS] - 1] = STATUS.NORMAL;
    sh.appendRow(row);
    var rowIdx = sh.getLastRow();
    sh.getRange(rowIdx, hm[COL.DATE]).setNumberFormat('MM"월" dd"일"');
    // appendRow 는 '18:30' 을 시간값으로 자동 변환하므로 텍스트 서식 지정 후 다시 기록
    sh.getRange(rowIdx, hm[COL.TIME]).setNumberFormat('@').setValue(s.time || '');
    return { id: id, row: rowIdx };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 이체내역 ----------

/** 오늘(day, 1~31)이 이체일인 [이체내역] 행을 읽어 표시용 문자열 배열로 반환 */
function getTodayTransfers_(day) {
  var sh = ss_().getSheetByName(SHEET.TRANSFER);
  if (!sh || sh.getLastRow() < 2) return [];
  var hm = headerMap_(sh);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var out = [];
  rows.forEach(function (r) {
    var dayVal = r[hm[TRANSFER_COL.DAY] - 1];
    if (Number(dayVal) !== day) return;
    var holder = String(r[hm[TRANSFER_COL.HOLDER] - 1] || '').trim();
    if (!holder) return;
    var bank = String(r[hm[TRANSFER_COL.BANK] - 1] || '').trim();
    var account = String(r[hm[TRANSFER_COL.ACCOUNT] - 1] || '').trim();
    var amount = r[hm[TRANSFER_COL.AMOUNT] - 1];
    var memo = String(r[hm[TRANSFER_COL.MEMO] - 1] || '').trim();
    var amountText = amount ? Number(amount).toLocaleString('ko-KR') + '원' : '';
    out.push('- ' + holder + (bank || account ? ' / ' + bank + ' ' + account : '') + (amountText ? ' / ' + amountText : '') + (memo ? ' (' + memo + ')' : ''));
  });
  return out;
}

// ---------- 로그 (ADM-03) ----------

function log_(kind, scheduleId, channel, actor, result, detail) {
  try {
    var sh = ss_().getSheetByName(SHEET.LOG);
    if (!sh) return;
    sh.appendRow([nowString_(), kind, scheduleId || '', channel || '', actor || '', result || '', String(detail || '').slice(0, 500)]);
  } catch (e) {
    console.error('log_ failed: ' + e);
  }
}
