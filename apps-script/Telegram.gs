/**
 * 텔레그램 봇 (REG-02, REG-05, PAR-07, ADM-02, VIEW-06). Webhook 방식, 1:1 대화 (D-10 확정).
 *
 * 보안: Apps Script 웹앱은 익명 접근으로 배포되므로 웹훅 URL 에 비밀 키(?key=)를 붙이고,
 *       발신자 텔레그램 ID 를 [구성원] 시트와 대조하는 이중 검증을 한다 (명세 7장 보안).
 */

var TG_API = 'https://api.telegram.org/bot';

function tgToken_() {
  var t = getProp_('BOT_TOKEN');
  if (!t) throw new Error('스크립트 속성 BOT_TOKEN 이 없습니다.');
  return t;
}

/** 텔레그램 API 호출. 실패 시 재시도 (SYNC-07) */
function tgCall_(method, payload) {
  var retries = Number(getSettings_()['재시도횟수'] || 3);
  var lastErr = null;
  for (var i = 0; i <= retries; i++) {
    try {
      var res = UrlFetchApp.fetch(TG_API + tgToken_() + '/' + method, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      var body = JSON.parse(res.getContentText());
      if (body.ok) return body.result;
      lastErr = body.description || res.getContentText();
      // 403(차단)·400(잘못된 chat) 은 재시도해도 소용없음
      if (res.getResponseCode() === 403 || res.getResponseCode() === 400) break;
    } catch (e) {
      lastErr = String(e);
    }
    Utilities.sleep(500 * (i + 1));
  }
  throw new Error('Telegram ' + method + ' 실패: ' + lastErr);
}

function tgSend_(chatId, text) {
  return tgCall_('sendMessage', { chat_id: String(chatId), text: text, disable_web_page_preview: true });
}

// ---------- Webhook ----------

function doGet(e) {
  return HtmlService.createHtmlOutput('ZPR schedule bot: ok');
}

function doPost(e) {
  try {
    var key = getProp_('WEBHOOK_KEY');
    if (!key || !e || !e.parameter || e.parameter.key !== key) {
      log_('오류', '', CHANNEL.TELEGRAM, '', '실패', '웹훅 키 불일치');
      return HtmlService.createHtmlOutput('forbidden');
    }
    var update = JSON.parse(e.postData.contents);
    if (isDuplicateUpdate_(update.update_id)) return HtmlService.createHtmlOutput('dup');
    var msg = update.message || update.edited_message;
    if (msg && msg.text && msg.chat && msg.chat.type === 'private') handleMessage_(msg);
  } catch (err) {
    log_('오류', '', CHANNEL.TELEGRAM, '', '실패', 'doPost: ' + err + ' ' + (err.stack || ''));
  }
  return HtmlService.createHtmlOutput('ok');
}

/** 텔레그램은 응답이 없으면 같은 update 를 재전송한다. update_id 로 중복 처리 방지 */
function isDuplicateUpdate_(updateId) {
  if (updateId === undefined) return false;
  var cache = CacheService.getScriptCache();
  var k = 'upd_' + updateId;
  if (cache.get(k)) return true;
  cache.put(k, '1', 21600);
  return false;
}

function handleMessage_(msg) {
  var chatId = msg.chat.id;
  var tgId = String(msg.from.id);
  var text = String(msg.text || '').trim();
  var member = findMemberByTgId_(tgId);

  // 계정 연결 (ADM-02)
  if (/^\/start/.test(text)) {
    if (member) { tgSend_(chatId, '[연결 완료] ' + member.name + '님, 이미 연결되어 있습니다.\n\n' + getSettings_()['봇사용법']); return; }
    tgSend_(chatId, '제피로스 일정관리 봇입니다.\n관리자에게 받은 6자리 인증코드를 입력해 주세요.');
    return;
  }
  if (!member) {
    if (/^\d{6}$/.test(text)) { linkAccount_(chatId, tgId, msg.from, text); return; }
    tgSend_(chatId, '[등록 불가] 등록된 구성원이 아닙니다. 관리자에게 문의하세요.\n(인증코드가 있으면 6자리 숫자를 입력해 주세요)');
    log_('오류', '', CHANNEL.TELEGRAM, tgId, '거부', '미등록 사용자: ' + text);
    return;
  }

  // 조회 명령 (VIEW-06, 선택)
  if (/^(오늘|내일|이번주|이번 주)\s*일정$/.test(text) || text === '/today' || text === '/week') {
    tgSend_(chatId, listSchedulesText_(chatId, member, text));
    return;
  }

  // 삭제 (REG-06)
  if (/^(취소|삭제)\s+\d+/.test(text)) { cancelFromTelegram_(chatId, member, text); return; }

  // 수정 (REG-06)
  if (/^(변경|수정)\s+\d+/.test(text)) { editFromTelegram_(chatId, member, text); return; }

  // 등록 (REG-02)
  if (/^일정(\s|$)/.test(text)) { registerFromTelegram_(chatId, member, text); return; }

  tgSend_(chatId, '[사용법]\n' + getSettings_()['봇사용법'] +
    '\n\n조회: "오늘 일정", "이번주 일정"' +
    '\n수정: 목록 번호 확인 후 "변경 2 15:00" 처럼 바꿀 내용만 입력' +
    '\n삭제: "취소 2" 또는 "삭제 2"');
}

/** 목록에서 보여준 번호(idx, 1부터) -> {row, id}. 목록을 먼저 조회해야 유효하다 */
function resolveListedSchedule_(chatId, idx) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('list_' + chatId);
  if (!raw) return null;
  var arr = JSON.parse(raw);
  return arr[idx - 1] || null;
}

function cancelFromTelegram_(chatId, member, text) {
  var c = parseCancelCommand(text);
  if (!c.ok) { tgSend_(chatId, '[삭제 실패] ' + c.message); return; }
  var listed = resolveListedSchedule_(chatId, c.index);
  if (!listed) { tgSend_(chatId, '[삭제 실패] 번호를 찾을 수 없습니다. 먼저 "오늘 일정"이나 "이번주 일정"으로 목록을 확인해 주세요.'); return; }
  var s = getSchedule_(listed.row);
  if (!s.id || s.id !== listed.id || s.status === STATUS.DELETED) { tgSend_(chatId, '[삭제 실패] 이미 삭제됐거나 바뀐 일정입니다. 목록을 다시 확인해 주세요.'); return; }
  if (!canEdit_(s, member)) { tgSend_(chatId, '[삭제 실패] 본인이 등록한 일정만 삭제할 수 있습니다.'); return; }

  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  setCell_(sh, hm, listed.row, COL.STATUS, STATUS.DELETED);
  setCell_(sh, hm, listed.row, COL.UPDATED, nowString_());
  log_('삭제', s.id, CHANNEL.TELEGRAM, member.name, '성공', scheduleText_(s));
  notifyDeleted_(s);
  tgSend_(chatId, '[삭제 완료] ' + scheduleText_(s));
}

function editFromTelegram_(chatId, member, text) {
  var settings = getSettings_();
  var names = getMembers_().map(function (m) { return m.name; });
  var c = parseEditCommand(text, { members: names, pmFrom: Number(settings['오후해석시작시'] || 0) });
  if (!c.ok) { tgSend_(chatId, '[변경 실패] ' + c.message + '\n예) 변경 2 15:00  /  변경 2 9/20 방산IR 제출'); return; }
  var listed = resolveListedSchedule_(chatId, c.index);
  if (!listed) { tgSend_(chatId, '[변경 실패] 번호를 찾을 수 없습니다. 먼저 "오늘 일정"이나 "이번주 일정"으로 목록을 확인해 주세요.'); return; }
  var s = getSchedule_(listed.row);
  if (!s.id || s.id !== listed.id || s.status === STATUS.DELETED) { tgSend_(chatId, '[변경 실패] 이미 삭제됐거나 바뀐 일정입니다. 목록을 다시 확인해 주세요.'); return; }
  if (!canEdit_(s, member)) { tgSend_(chatId, '[변경 실패] 본인이 등록한 일정만 수정할 수 있습니다.'); return; }

  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  var before = formatMonthDay(s.date) + ' ' + (s.time || '종일');
  if (c.dateFound) {
    var p = c.date.split('-').map(Number);
    sh.getRange(listed.row, hm[COL.DATE]).setValue(new Date(p[0], p[1] - 1, p[2]));
  }
  if (c.timeFound) sh.getRange(listed.row, hm[COL.TIME]).setNumberFormat('@').setValue(c.time);
  if (c.title) setCell_(sh, hm, listed.row, COL.TITLE, c.title);
  if (c.targets.length && hm[COL.TARGET]) {
    setCell_(sh, hm, listed.row, COL.TARGET, c.targets.indexOf('전부') >= 0 ? '전체' : c.targets.join(' '));
    applyTargetColumn_(sh, hm, listed.row, getSchedule_(listed.row));
  }
  setCell_(sh, hm, listed.row, COL.UPDATED, nowString_());

  var updated = getSchedule_(listed.row);
  log_('수정', s.id, CHANNEL.TELEGRAM, member.name, '성공', text);
  if (s.alarm === ALARM.SENT) notifyChanged_(updated, before);
  tgSend_(chatId, '[변경 완료] ' + scheduleText_(updated));
}

function linkAccount_(chatId, tgId, from, code) {
  var m = getMembers_().filter(function (x) { return x.code === code; })[0];
  if (!m) {
    tgSend_(chatId, '[연결 실패] 인증코드가 올바르지 않습니다. 관리자에게 확인해 주세요.');
    log_('오류', '', CHANNEL.TELEGRAM, tgId, '실패', '인증코드 불일치: ' + code);
    return;
  }
  var sh = sheet_(SHEET.MEMBER);
  var hm = headerMap_(sh);
  sh.getRange(m.row, hm[MEMBER_COL.TG_ID]).setValue(tgId);
  sh.getRange(m.row, hm[MEMBER_COL.CODE]).setValue('');
  tgSend_(chatId, '[연결 완료] ' + m.name + '님, 텔레그램 계정이 연결되었습니다.\n\n' + getSettings_()['봇사용법']);
  log_('등록', '', CHANNEL.TELEGRAM, m.name, '성공', '텔레그램 계정 연결 (' + (from.username ? '@' + from.username : tgId) + ')');
}

function registerFromTelegram_(chatId, member, text) {
  var settings = getSettings_();
  var names = getMembers_().map(function (m) { return m.name; });
  var r = parseScheduleCommand(text, { members: names, pmFrom: Number(settings['오후해석시작시'] || 0) });
  if (!r.ok) {
    tgSend_(chatId, '[등록 실패] ' + r.message + '\n입력 예) 일정 6/6 18:30 전체회의');
    log_('오류', '', CHANNEL.TELEGRAM, member.name, '실패', r.error + ': ' + text);
    return;
  }
  var targets = r.targets.length ? r.targets : [member.name];
  var res = appendSchedule_({ date: r.date, time: r.time, title: r.title, memo: '', targets: targets, registrant: member.name, channel: CHANNEL.TELEGRAM });
  log_('등록', res.id, CHANNEL.TELEGRAM, member.name, '성공', text);

  var s = getSchedule_(res.row);
  tgSend_(chatId,
    '[일정 등록 완료]\n날짜: ' + formatDateKo(r.date) + '  시간: ' + (r.time || '종일') +
    '\n일정: ' + r.title + '  대상: ' + targetLabel_(s));

  if (settings['즉시알림'] !== 'N') notifyImmediate_(s, member.name);
}

/** VIEW-06: 본인 열람 가능 일정 목록. 표시한 번호는 캐시에 남겨 "변경 N"·"취소 N" 명령이 참조한다 */
function listSchedulesText_(chatId, member, text) {
  var today = new Date();
  var from = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var to = new Date(from);
  var label = '오늘';
  if (/내일/.test(text)) { from.setDate(from.getDate() + 1); to = new Date(from); label = '내일'; }
  else if (/이번\s?주|week/.test(text)) { to.setDate(from.getDate() + (7 - from.getDay())); label = '이번 주'; }
  to.setDate(to.getDate() + 1);
  var fromS = Utilities.formatDate(from, 'Asia/Seoul', 'yyyy-MM-dd');
  var toS = Utilities.formatDate(to, 'Asia/Seoul', 'yyyy-MM-dd');
  var list = getAllSchedules_().filter(function (s) {
    return s.status !== STATUS.DELETED && s.date >= fromS && s.date < toS && canView_(s, member);
  }).sort(function (a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });

  if (!list.length) { CacheService.getScriptCache().remove('list_' + chatId); return '[' + label + ' 일정] 없음'; }

  CacheService.getScriptCache().put('list_' + chatId, JSON.stringify(list.map(function (s) { return { row: s.row, id: s.id }; })), 1800);
  var body = list.map(function (s, i) {
    return (i + 1) + '. ' + formatMonthDay(s.date) + ' ' + (s.time || '종일') + ' ' + s.title + (s.memo ? ' (' + s.memo + ')' : '');
  }).join('\n');
  return '[' + label + ' 일정]\n' + body + '\n\n수정: "변경 번호 ..."  삭제: "취소 번호" (본인 등록 건만)';
}
