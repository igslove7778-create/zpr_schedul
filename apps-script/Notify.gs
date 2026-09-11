/**
 * 텔레그램 알림 발송 (AUTH-03, AUTH-05, VIEW-03). 발송 시점 D-05 기본값: 등록 즉시 + 당일 리마인드 + 사전 리마인드.
 * 결과는 [일정] E열(알람)에 대기 / 발송완료 / 실패 로 기록한다.
 */

function scheduleText_(s) {
  return formatMonthDay(s.date) + ' ' + (s.time || '종일') + ' ' + s.title + (s.memo ? '\n메모: ' + s.memo : '');
}

/** 대상자에게 텍스트 발송. 반환 {sent:[names], failed:[names]} */
function broadcast_(s, text, excludeName) {
  var res = { sent: [], failed: [], skipped: [] };
  notifyTargets_(s).forEach(function (m) {
    if (excludeName && m.name === excludeName) { res.skipped.push(m.name); return; }
    if (!m.tgId) { res.failed.push(m.name + '(미연결)'); return; }
    try { tgSend_(m.tgId, text); res.sent.push(m.name); }
    catch (e) { res.failed.push(m.name); log_('알림', s.id, CHANNEL.TELEGRAM, m.name, '실패', String(e)); }
  });
  return res;
}

function markAlarm_(s, res) {
  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  var status = res.failed.length && !res.sent.length && !res.skipped.length ? ALARM.FAILED : ALARM.SENT;
  setCell_(sh, hm, s.row, COL.ALARM, status);
  log_('알림', s.id, CHANNEL.TELEGRAM, '시스템', status,
    '발송: ' + res.sent.join(',') + (res.failed.length ? ' / 실패: ' + res.failed.join(',') : '') + (res.skipped.length ? ' / 제외: ' + res.skipped.join(',') : ''));
}

/** 등록 즉시 알림. 등록자 본인은 등록 완료 회신을 이미 받았으므로 제외 */
function notifyImmediate_(s, excludeName) {
  var res = broadcast_(s, '[일정 알림] ' + scheduleText_(s) + '\n대상: ' + targetLabel_(s) + '\n등록: ' + (s.registrant || '-'), excludeName);
  markAlarm_(s, res);
}

/** 변경 재알림 (AUTH-05) */
function notifyChanged_(s, before) {
  var text = '[일정 변경] ' + (before ? before + ' → ' : '') + scheduleText_(s) + '\n대상: ' + targetLabel_(s);
  markAlarm_(s, broadcast_(s, text));
}

/** 삭제 알림 (AUTH-05) */
function notifyDeleted_(s) {
  var res = broadcast_(s, '[일정 취소] ' + scheduleText_(s));
  log_('삭제', s.id, CHANNEL.SHEET, '시스템', '성공', '취소 알림: ' + res.sent.join(','));
}

/** 알람=대기 인 행을 모두 발송 (메뉴 / 수동 재처리용) */
function sendPendingNotifications() {
  var n = 0;
  getAllSchedules_().forEach(function (s) {
    if (s.alarm !== ALARM.PENDING || s.status === STATUS.DELETED || !s.id) return;
    notifyImmediate_(s, null); n++;
  });
  SpreadsheetApp.getActive().toast(n + '건 발송 처리', '일정관리');
}

/**
 * 리마인드 (시간 기반 트리거, 10분 간격). 리마인드 열에 발송 이력을 남겨 중복 발송을 막는다.
 *  - 당일 리마인드: 설정 '당일리마인드시각' 이후 첫 실행 시 오늘 일정 전체
 *  - 사전 리마인드: 시간이 있는 일정의 N분 전
 */
function runReminders() {
  var settings = getSettings_();
  var now = new Date();
  var todayS = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
  var nowHM = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  var dayAt = String(settings['당일리마인드시각'] || '').trim();
  var preMin = Number(settings['사전리마인드분'] || 0);
  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);

  getAllSchedules_().forEach(function (s) {
    if (s.status === STATUS.DELETED || !s.id || s.date !== todayS) return;
    var flags = s.remind ? s.remind.split(';') : [];
    var changed = false;

    if (dayAt && nowHM >= dayAt && flags.indexOf('당일') < 0) {
      var r = broadcast_(s, '[오늘 일정] ' + scheduleText_(s) + '\n대상: ' + targetLabel_(s));
      log_('알림', s.id, CHANNEL.TELEGRAM, '시스템', '당일', '발송: ' + r.sent.join(',') + (r.failed.length ? ' / 실패: ' + r.failed.join(',') : ''));
      flags.push('당일'); changed = true;
    }
    if (preMin > 0 && s.time && flags.indexOf('사전') < 0) {
      var p = s.time.split(':').map(Number);
      var start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), p[0], p[1]);
      var diffMin = (start.getTime() - now.getTime()) / 60000;
      if (diffMin <= preMin && diffMin > -10) {
        var r2 = broadcast_(s, '[' + preMin + '분 전] ' + scheduleText_(s));
        log_('알림', s.id, CHANNEL.TELEGRAM, '시스템', '사전', '발송: ' + r2.sent.join(',') + (r2.failed.length ? ' / 실패: ' + r2.failed.join(',') : ''));
        flags.push('사전'); changed = true;
      }
    }
    if (changed) setCell_(sh, hm, s.row, COL.REMIND, flags.join(';'));
  });
}

/** 관리자 텔레그램 통보 (SYNC-07) */
function notifyAdmins_(text) {
  getAdmins_().forEach(function (a) { if (a.tgId) { try { tgSend_(a.tgId, '[관리자] ' + text); } catch (e) { } } });
}
