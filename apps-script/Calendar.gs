/**
 * 구글 캘린더 연동 (SYNC-02, D-08 구성원별 캘린더).
 * - 팀 캘린더('제피로스 팀 일정') 에는 모든 일정을 넣는다 (기존 동작 유지).
 * - 대상자마다 본인 명의 캘린더('제피로스 일정 - 이름') 에도 같은 이벤트를 넣는다.
 *   구성원별 캘린더는 이 스크립트를 실행하는 관리자 계정에 만들어지고, [구성원] '구글 계정' 으로 공유(쓰기)된다.
 *   -> 구성원은 본인 구글 캘린더(앱·위젯)에서 바로 볼 수 있다.
 * - [일정] '캘린더ID' 열에 "팀:이벤트ID;성도형:이벤트ID;…" 형태로 저장해 재실행해도 중복 생성하지 않고 갱신한다.
 *   (예전 형식인 이벤트ID 하나만 있으면 팀 캘린더 이벤트로 본다)
 * - 대상자에서 빠진 사람의 이벤트는 지우고, [일정] 상태가 '삭제'면 전부 지운다.
 * - '종료시간' 열이 없으면 자동으로 추가한다. 비워두면 기본일정길이분(설정)만큼 길이로 처리.
 * - 10분마다 자동 실행(메뉴 '캘린더 자동 동기화 켜기') + 메뉴 '캘린더 지금 동기화'.
 *
 * 주의: 캘린더는 만든 계정만 접근할 수 있으므로, 자동 동기화 트리거와 캘린더 생성은 같은 관리자 계정에서 한다.
 */

var END_TIME_HEADER_ = '종료시간';
var TEAM_KEY_ = '팀';
var MEMBER_CAL_PREFIX_ = '제피로스 일정 - ';

// ---------- 캘린더 준비 ----------

function getTeamCalendarId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TEAM_CAL_ID');
  if (id) {
    try { CalendarApp.getCalendarById(id).getName(); return id; } catch (e) { /* 재생성 */ }
  }
  var cal = CalendarApp.createCalendar('제피로스 팀 일정');
  props.setProperty('TEAM_CAL_ID', cal.getId());
  return cal.getId();
}

/** [구성원] 시트에 '캘린더ID' 헤더가 없으면 맨 끝열 다음에 추가 */
function ensureMemberCalColumn_() {
  var sh = sheet_(SHEET.MEMBER);
  var hm = headerMap_(sh);
  if (hm[MEMBER_COL.CAL_ID]) return hm;
  var col = Math.max(sh.getLastColumn(), 1) + 1;
  sh.getRange(1, col).setValue(MEMBER_COL.CAL_ID).setFontWeight('bold').setBackground('#DCE6F1');
  return headerMap_(sh);
}

/** 캘린더를 구성원 구글 계정에 공유 (고급 Calendar 서비스 필요). 실패해도 동기화는 계속한다 */
function shareCalendarWith_(calId, email, name) {
  if (!email) return '이메일 없음';
  if (typeof Calendar === 'undefined' || !Calendar.Acl) return '고급 Calendar 서비스 미사용';
  try {
    var acl = Calendar.Acl.list(calId);
    var already = (acl.items || []).some(function (a) { return a.scope && a.scope.type === 'user' && String(a.scope.value).toLowerCase() === email.toLowerCase(); });
    if (already) return '공유됨';
    Calendar.Acl.insert({ role: 'writer', scope: { type: 'user', value: email } }, calId);
    return '공유 완료';
  } catch (e) {
    log_('오류', '', CHANNEL.CALENDAR, name, '실패', '캘린더 공유(' + email + '): ' + e);
    return '공유 실패';
  }
}

/**
 * 구성원 본인 캘린더 ID. 없으면 만들고 [구성원] 시트 '캘린더ID' 에 기록 + 구글 계정에 공유.
 * 구글 계정이 없는 구성원은 캘린더를 만들지 않는다 (null).
 */
function getMemberCalendarId_(member, createIfMissing) {
  if (!member || !member.email) return null;
  if (member.calId) {
    try { CalendarApp.getCalendarById(member.calId).getName(); return member.calId; } catch (e) { /* 접근 불가 -> 재생성 */ }
  }
  if (!createIfMissing) return null;
  var cal = CalendarApp.createCalendar(MEMBER_CAL_PREFIX_ + member.name);
  var id = cal.getId();
  var hm = ensureMemberCalColumn_();
  sheet_(SHEET.MEMBER).getRange(member.row, hm[MEMBER_COL.CAL_ID]).setValue(id);
  member.calId = id;
  var shared = shareCalendarWith_(id, member.email, member.name);
  log_('설정', '', CHANNEL.CALENDAR, member.name, '성공', '구성원 캘린더 생성 (' + shared + '): ' + id);
  return id;
}

/** 메뉴: 구글 계정이 있는 사용중 구성원 전원의 캘린더를 만들고 공유한다 */
function setupMemberCalendars() {
  ensureMemberCalColumn_();
  var lines = [];
  getMembers_().forEach(function (m) {
    if (!m.email) { lines.push(m.name + ': 구글 계정 없음 (건너뜀)'); return; }
    var id = getMemberCalendarId_(m, true);
    var shared = shareCalendarWith_(id, m.email, m.name);
    lines.push(m.name + ': ' + shared);
  });
  SpreadsheetApp.getUi().alert('구성원 캘린더\n\n' + lines.join('\n') + '\n\n각 구성원의 구글 캘린더 "다른 캘린더" 목록에 "' + MEMBER_CAL_PREFIX_ + '이름" 이 나타납니다.');
}

// ---------- 캘린더ID 셀 형식: "팀:id;이름:id;…" ----------

function parseCalIds_(str) {
  var out = {};
  String(str || '').split(';').forEach(function (p) {
    p = p.trim();
    if (!p) return;
    var i = p.indexOf(':');
    if (i < 0) { out[TEAM_KEY_] = p; return; } // 예전 형식 (이벤트ID 하나)
    out[p.slice(0, i).trim()] = p.slice(i + 1).trim();
  });
  return out;
}

function formatCalIds_(map) {
  return Object.keys(map).filter(function (k) { return map[k]; }).map(function (k) { return k + ':' + map[k]; }).join(';');
}

// ---------- 동기화 ----------

/** [일정] 시트에 '종료시간' 헤더가 없으면 맨 끝열 다음에 추가 */
function ensureEndTimeColumn_() {
  var sh = sheet_(SHEET.SCHEDULE);
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(END_TIME_HEADER_) >= 0) return;
  sh.getRange(1, lastCol + 1).setValue(END_TIME_HEADER_).setFontWeight('bold');
}

/** 이벤트 생성 또는 갱신. 반환: 이벤트ID ('' 이면 실패) */
function upsertEvent_(cal, evId, title, desc, start, end, allDay) {
  var ev = null;
  if (evId) { try { ev = cal.getEventById(evId); } catch (e) { ev = null; } }
  if (ev) {
    ev.setTitle(title);
    ev.setDescription(desc);
    if (allDay) ev.setAllDayDate(start); else ev.setTime(start, end);
    return ev.getId();
  }
  var created = allDay
    ? cal.createAllDayEvent(title, start, { description: desc })
    : cal.createEvent(title, start, end, { description: desc });
  return created.getId();
}

function deleteEvent_(cal, evId) {
  if (!cal || !evId) return;
  try { cal.getEventById(evId).deleteEvent(); } catch (e) { /* 이미 없음 */ }
}

function syncCalendar_() {
  var sh = sheet_(SHEET.SCHEDULE);
  if (!sh) return;
  ensureEndTimeColumn_();
  var hm = headerMap_(sh);
  if (!hm[COL.DATE] || !hm[COL.TITLE] || !hm[COL.CAL_ID]) return;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var settings = getSettings_();
  var defaultMin = Number((settings && settings['기본일정길이분']) || 60) || 60;
  var members = getMembers_();
  var byName = {};
  members.forEach(function (m) { byName[m.name] = m; });
  var checkNames = checkColumnNames_(hm);

  var calCache = {};
  function calFor(key) {
    if (calCache[key] !== undefined) return calCache[key];
    var id = null;
    if (key === TEAM_KEY_) id = getTeamCalendarId_();
    else if (byName[key]) id = getMemberCalendarId_(byName[key], true);
    var cal = null;
    if (id) { try { cal = CalendarApp.getCalendarById(id); } catch (e) { cal = null; } }
    calCache[key] = cal;
    return cal;
  }

  var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (var r = 0; r < data.length; r++) {
    var values = data[r];
    var sheetRow = r + 2;
    var s = rowToSchedule_(hm, sheetRow, values, checkNames);
    var ids = parseCalIds_(s.calId);
    var calCell = sh.getRange(sheetRow, hm[COL.CAL_ID]);

    // 삭제 또는 빈 행: 남은 이벤트 전부 제거
    if (!s.title || s.status === STATUS.DELETED) {
      if (Object.keys(ids).length) {
        Object.keys(ids).forEach(function (k) { deleteEvent_(calFor(k), ids[k]); });
        calCell.setValue('');
      }
      continue;
    }
    if (!s.date) continue;
    var p = s.date.split('-').map(Number);
    var dateObj = new Date(p[0], p[1] - 1, p[2]);
    if (isNaN(dateObj.getTime())) continue;

    // 시간 계산
    var endStr = hm[END_TIME_HEADER_] ? values[hm[END_TIME_HEADER_] - 1] : '';
    if (endStr instanceof Date) endStr = Utilities.formatDate(endStr, 'Asia/Seoul', 'HH:mm');
    endStr = String(endStr || '').trim();
    var start, end, allDay = false;
    if (s.time) {
      var tp = s.time.split(':').map(Number);
      start = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), tp[0] || 0, tp[1] || 0);
      if (endStr) {
        var ep = endStr.split(':').map(Number);
        end = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), ep[0] || 0, ep[1] || 0);
        if (end.getTime() <= start.getTime()) end = new Date(start.getTime() + defaultMin * 60000);
      } else {
        end = new Date(start.getTime() + defaultMin * 60000);
      }
    } else {
      allDay = true;
      start = dateObj;
    }
    var title = String(s.title);
    var desc = (s.memo || '') + '\n대상: ' + targetLabel_(s) + (s.registrant ? '\n등록: ' + s.registrant : '');

    // 넣어야 할 캘린더: 팀 + 대상자 중 구글 계정이 있는 구성원
    var wanted = [TEAM_KEY_];
    targetNames_(s).forEach(function (n) { if (byName[n] && byName[n].email && wanted.indexOf(n) < 0) wanted.push(n); });

    var next = {};
    var changed = false;
    // 대상자에서 빠진 사람의 이벤트 제거
    Object.keys(ids).forEach(function (k) {
      if (wanted.indexOf(k) < 0) { deleteEvent_(calFor(k), ids[k]); changed = true; }
    });
    // 생성·갱신
    wanted.forEach(function (k) {
      var cal = calFor(k);
      if (!cal) { if (ids[k]) next[k] = ids[k]; return; }
      try {
        var newId = upsertEvent_(cal, ids[k], title, desc, start, end, allDay);
        next[k] = newId;
        if (newId !== ids[k]) changed = true;
      } catch (e) {
        if (ids[k]) next[k] = ids[k];
        log_('오류', s.id, CHANNEL.CALENDAR, '시스템', '실패', '캘린더 동기화(' + k + '): ' + e);
      }
    });
    if (changed) calCell.setValue(formatCalIds_(next));
  }
}

function syncCalendarNow() {
  syncCalendar_();
  SpreadsheetApp.getActive().toast('캘린더 동기화 완료', '일정관리');
}

/** 메뉴: 10분마다 syncCalendar_ 를 실행하는 트리거를 현재 계정 소유로 설치하고 바로 한 번 동기화 */
function enableCalendarSync() {
  setupCalendarTrigger_();
  syncCalendar_();
  var me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) { /* 무시 */ }
  log_('설정', '', CHANNEL.CALENDAR, me || '시스템', '성공', '캘린더 자동 동기화 트리거 설치');
  SpreadsheetApp.getActive().toast('캘린더 자동 동기화를 ' + (me || '내 계정') + ' 소유로 켰습니다 (10분마다).', '일정관리', 8);
}

function setupCalendarTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncCalendar_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncCalendar_').timeBased().everyMinutes(10).create();
}

/** 편집기에서 직접 실행하는 초기 설정 (메뉴 '캘린더 자동 동기화 켜기' 와 같음) */
function aaaSetupCalendar() {
  var calId = getTeamCalendarId_();
  ensureEndTimeColumn_();
  setupCalendarTrigger_();
  syncCalendar_();
  Logger.log('팀 캘린더 ID: ' + calId);
  Logger.log('캘린더 링크: https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(calId));
}
