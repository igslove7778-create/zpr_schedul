function zNoticeFromRow_(row, index, map) {
  function value(name) { return map[name] ? row[map[name] - 1] : ''; }
  return {
    row: index,
    // 날짜는 Date 객체를 그대로 보존해야 9월 24일 같은 입력이 오늘 날짜로 바뀌지 않는다.
    writtenAt: value(ZEPHYRUS.noticeCol.writtenAt),
    // 제목 칸이 시트의 시간 값이면 사람이 읽는 '오후 2:00' 형식으로 바꾼다.
    title: zNoticeTimeText_(value(ZEPHYRUS.noticeCol.title)),
    content: String(value(ZEPHYRUS.noticeCol.content) || '').trim(),
    alarm: String(value(ZEPHYRUS.noticeCol.alarm) || '').trim(),
    id: String(value(ZEPHYRUS.noticeCol.id) || '').trim(),
    registrant: String(value(ZEPHYRUS.noticeCol.registrant) || '').trim(),
    status: String(value(ZEPHYRUS.noticeCol.status) || '').trim(),
    teamEvent: String(value(ZEPHYRUS.noticeCol.teamEvent) || '').trim()
  };
}

function zSetNotice_(sheet, rowNumber, map, name, value) {
  if (!map[name]) throw new Error('공지사항 시트에 열이 없습니다: ' + name);
  sheet.getRange(rowNumber, map[name]).setValue(value);
}

function zNoticeDateText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, ZEPHYRUS.kst, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  var normalized = zNormalizeDate_(text);
  if (normalized) return normalized;

  // 공지사항에서는 "24일" 또는 "24"처럼 날짜만 적어도 현재 연/월로 처리한다.
  var match = text.match(/^(\d{1,2})\s*일?$/);
  if (!match) return '';
  var now = new Date();
  var year = Utilities.formatDate(now, ZEPHYRUS.kst, 'yyyy');
  var month = Utilities.formatDate(now, ZEPHYRUS.kst, 'M');
  return zDateTextFromParts_(year, month, match[1]);
}

function zAnnouncementText_(notice) {
  var dateText = zNoticeDateText_(notice.writtenAt) || zToday_();
  return '[공지사항]\n날짜: ' + dateText + '\n' + notice.title + '\n\n' + notice.content;
}

function zNoticeMarker_(noticeId) {
  return '[ZEPHYRUS_NOTICE:' + noticeId + ']';
}

function zMarkerNoticeId_(description) {
  var match = String(description || '').match(/\[ZEPHYRUS_NOTICE:([^\]]+)\]/);
  return match ? match[1] : '';
}

// A notice is visible in the public team calendar and in every dedicated
// calendar.  It begins one minute later with a one-minute popup reminder, so
// Google Calendar can deliver a phone/app alert without calendar defaults.
function zCreateNoticeEvent_(calendarId, notice) {
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Apps Script 왼쪽 “서비스”의 Calendar API가 필요합니다.');
  }
  var ymd = zNoticeDateText_(notice.writtenAt) || zToday_();
  var parts = ymd.split('-').map(Number);
  var start = new Date(parts[0], parts[1] - 1, parts[2], 9, 0, 0);
  var end = new Date(start.getTime() + 30 * 60000);

  // 같은 공지ID가 이미 이 캘린더에 있으면 새로 만들지 않는다.
  var marker = zNoticeMarker_(notice.id);
  var searchFrom = new Date(start.getTime() - 86400000).toISOString();
  var searchUntil = new Date(end.getTime() + 86400000).toISOString();
  try {
    var existing = Calendar.Events.list(calendarId, {
      timeMin: searchFrom,
      timeMax: searchUntil,
      singleEvents: true,
      showDeleted: false,
      maxResults: 250
    });
    var found = (existing.items || []).filter(function(item) {
      return String(item.description || '').indexOf(marker) !== -1;
    })[0] || null;
    if (found) return found;
  } catch (ignore) {}

  var resource = {
    summary: '[공지] ' + notice.title,
    description: '공지일: ' + ymd + '\n' + notice.content + '\n\n' + marker,
    start: { dateTime: start.toISOString(), timeZone: ZEPHYRUS.kst },
    end: { dateTime: end.toISOString(), timeZone: ZEPHYRUS.kst },
    reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 1 }] }
  };
  return Calendar.Events.insert(resource, calendarId, { sendUpdates: 'none' });
}

function zCreateMemberNoticeEvent_(member, notice) {
  return zCreateNoticeEvent_(member.calendarId, notice);
}

function zCreateTeamNoticeEvent_(notice) {
  return zCreateNoticeEvent_(zTeamCalendarId_(), notice);
}

function zPublishNotice_(notice) {
  var result = { telegramSent: [], calendarSent: [], teamEventId: '', failed: [] };
  try {
    result.teamEventId = String(zCreateTeamNoticeEvent_(notice).id || '');
  } catch (teamError) {
    result.failed.push('공용캘린더');
    zLog_('오류', notice.id, '공용캘린더', '실패', '공지 등록: ' + teamError);
  }
  zMembers_().forEach(function(member) {
    if (member.telegramId) {
      try {
        zSendTelegram_(member.telegramId, zAnnouncementText_(notice));
        result.telegramSent.push(member.name);
      } catch (error) {
        result.failed.push(member.name + '(텔레그램)');
        zLog_('오류', notice.id, '텔레그램', '실패', '공지 ' + member.name + ': ' + error);
      }
    }
    if (member.calendarId) {
      try {
        zCreateMemberNoticeEvent_(member, notice);
        result.calendarSent.push(member.name);
      } catch (error2) {
        result.failed.push(member.name + '(개인캘린더)');
        zLog_('오류', notice.id, '개인캘린더', '실패', '공지 ' + member.name + ': ' + error2);
      }
    }
  });
  return result;
}

function zHandleNoticeEdit_(event) {
  // 공지 편집끼리만 잠근다. 일반 일정 동기화의 ScriptLock과 충돌하지 않게
  // DocumentLock을 사용해 공지 등록이 느려지거나 건너뛰는 현상을 막는다.
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;
  try {
    var sheet = event.range.getSheet();
    var map = zHeaders_(sheet);
    for (var row = event.range.getRow(); row < event.range.getRow() + event.range.getNumRows(); row++) {
      // 락을 잡은 뒤 다시 읽어야 앞선 실행이 기록한 공지ID를 볼 수 있다.
      var values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      var notice = zNoticeFromRow_(values, row, map);
      if (!notice.title || !notice.content) continue;
      if (notice.id) {
        zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.updated, zNowText_());
        if (notice.status === ZEPHYRUS.status.normal) zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.status, '수정됨');
        zSaveNoticeSnapshot_(notice);
        continue;
      }

      notice.id = '공지-' + zId_();
      var enteredNoticeDate = notice.writtenAt;
      notice.writtenAt = zNoticeDateText_(enteredNoticeDate) || zToday_();
      notice.registrant = zRegistrantFromSheetEdit_(event, { registrant: '' });
      if (enteredNoticeDate === '' || enteredNoticeDate === null || enteredNoticeDate === undefined) {
        var todayObject = zDateObject_(notice.writtenAt);
        zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.writtenAt, todayObject || notice.writtenAt);
        if (map[ZEPHYRUS.noticeCol.writtenAt]) sheet.getRange(row, map[ZEPHYRUS.noticeCol.writtenAt]).setNumberFormat('m"월" d"일"');
      }

      // ID를 먼저 기록/flush해서 뒤늦게 들어온 다른 onEdit 실행이 재발행하지 못하게 한다.
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.id, notice.id);
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.registrant, notice.registrant);
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.updated, zNowText_());
      SpreadsheetApp.flush();
      zSaveNoticeSnapshot_(notice);

      var sent = zPublishNotice_(notice);
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.teamEvent, sent.teamEventId);
      var summary = '텔레그램 ' + sent.telegramSent.length + '명 / 공용캘린더 ' + (sent.teamEventId ? '1건' : '실패') + ' / 개인캘린더 ' + sent.calendarSent.length + '명';
      if (sent.failed.length) summary += ' / 실패 ' + sent.failed.join(', ');
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.alarm, summary);
      zSetNotice_(sheet, row, map, ZEPHYRUS.noticeCol.status, sent.failed.length ? '일부 실패' : ZEPHYRUS.status.normal);
      zSaveNoticeSnapshot_(notice);
    }
  } finally {
    lock.releaseLock();
  }
}

// One-time recovery for notices created before the public-calendar delivery
// was added.  It adds only the missing public event; Telegram and personal
// calendar notices are never sent a second time.
function zSyncNoticeTeamCalendar() {
  var sheet = zEnsureNoticeSheet_();
  var map = zHeaders_(sheet);
  var added = 0;
  var failed = [];
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(function(row, index) {
      var notice = zNoticeFromRow_(row, index + 2, map);
      if (!notice.id || !notice.title || !notice.content || notice.teamEvent) return;
      try {
        var event = zCreateTeamNoticeEvent_(notice);
        zSetNotice_(sheet, notice.row, map, ZEPHYRUS.noticeCol.teamEvent, String(event.id || ''));
        added++;
      } catch (error) {
        failed.push(notice.title);
        zLog_('오류', notice.id, '공용캘린더', '실패', '기존 공지 반영: ' + error);
      }
    });
  }
  if (failed.length) throw new Error('공용캘린더에 반영하지 못한 공지: ' + failed.join(', '));
  zNotice_(added + '건의 기존 공지를 공용캘린더에 반영했습니다.');
}

