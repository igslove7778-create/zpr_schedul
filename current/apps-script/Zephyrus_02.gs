function zEnsureSheet_(spreadsheet, name, headers) {
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#DCE6F1');
    return sheet;
  }
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  headers.forEach(function(header) {
    if (existing.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header).setFontWeight('bold').setBackground('#DCE6F1');
      existing.push(header);
    }
  });
  return sheet;
}

// The former "캘린더" tab was only a team-calendar shortcut.  It is now
// repurposed as the notice board the managers asked for.  This runs only from
// the dedicated finalisation command, so an ordinary structure check never
// overwrites the old shortcut by surprise.
function zEnsureNoticeSheet_() {
  var ss = zSpreadsheet_();
  var sheet = ss.getSheetByName(ZEPHYRUS.sheet.notice);
  var renamedLegacySheet = false;
  if (!sheet) {
    var legacy = ss.getSheetByName('캘린더');
    if (legacy) {
      legacy.setName(ZEPHYRUS.sheet.notice);
      sheet = legacy;
      renamedLegacySheet = true;
    }
  }
  if (!sheet) sheet = ss.insertSheet(ZEPHYRUS.sheet.notice);

  var headers = [
    ZEPHYRUS.noticeCol.writtenAt, ZEPHYRUS.noticeCol.title,
    ZEPHYRUS.noticeCol.content, ZEPHYRUS.noticeCol.alarm,
    ZEPHYRUS.noticeCol.id, ZEPHYRUS.noticeCol.registrant,
    ZEPHYRUS.noticeCol.updated, ZEPHYRUS.noticeCol.status,
    ZEPHYRUS.noticeCol.teamEvent
  ];
  // The legacy tab contains a large visual hyperlink rather than rows of
  // notice data.  Clearing it is intentional: the user asked to replace this
  // tab with 공지사항.  Existing 공지사항 rows are always preserved.
  if (renamedLegacySheet) sheet.clear();
  sheet = zEnsureSheet_(ss, ZEPHYRUS.sheet.notice, headers);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#DCE6F1');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 190);
  return sheet;
}

function zShouldKeepLog_(kind, result, detail) {
  var text = [kind, result, detail].join(' ');
  // Routine "성공", "반영 완료", webhook receipt and reminder records are
  // intentionally omitted.  The log is kept for failures and events a
  // manager may need to investigate later.
  return String(kind || '') === '오류' || /(실패|오류|거부|삭제|대상자 미반영|공유 실패|일부 실패|건너뜀)/.test(text);
}

function zCleanRoutineLogs_() {
  var sheet = zSheet_(ZEPHYRUS.sheet.log);
  if (sheet.getLastRow() < 2) return 0;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var kept = values.filter(function(row) { return zShouldKeepLog_(row[1], row[4], row[5]); });
  var removed = values.length - kept.length;
  if (!removed) return 0;
  sheet.getRange(2, 1, values.length, 6).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, 6).setValues(kept);
  return removed;
}

function zTrimLogSheetRows_() {
  var sheet = zSheet_(ZEPHYRUS.sheet.log);
  // Keep a small working area beneath the real last log, but remove the
  // thousands of now-empty rows left by an earlier content-only cleanup.
  var desiredRows = Math.max(30, sheet.getLastRow() + 10);
  var extra = sheet.getMaxRows() - desiredRows;
  if (extra > 0) sheet.deleteRows(desiredRows + 1, extra);
  return Math.max(extra, 0);
}

// One-time finishing command: it changes the old calendar-link tab into the
// notice board and removes only old routine-success logs.  It never changes
// schedules, team calendars, personal calendars, files or member records.
function zApplyNoticeAndLogFinalization() {
  zEnsureNoticeSheet_();
  var removed = zCleanRoutineLogs_();
  var trimmed = zTrimLogSheetRows_();
  zNotice_('하단 캘린더 탭을 공지사항으로 바꾸고, 일반 반영 로그 ' + removed + '건과 빈 행 ' + trimmed + '개를 정리했습니다. 공지사항에서는 공지제목과 공지내용을 모두 적으면 됩니다.');
}

function zPrepareSystem() {
  var ss = zSpreadsheet_();
  var scheduleSheet = ss.getSheetByName(ZEPHYRUS.sheet.schedule);
  if (scheduleSheet) zMigrateScheduleLayout_(scheduleSheet);
  scheduleSheet = zEnsureSheet_(ss, ZEPHYRUS.sheet.schedule, [
    ZEPHYRUS.col.date, ZEPHYRUS.col.time, ZEPHYRUS.col.endTime, ZEPHYRUS.col.title, ZEPHYRUS.col.location,
    ZEPHYRUS.col.targets, ZEPHYRUS.col.alarm, ZEPHYRUS.col.file, ZEPHYRUS.col.id, ZEPHYRUS.col.registrant,
    ZEPHYRUS.col.channel, ZEPHYRUS.col.calendarEvent, ZEPHYRUS.col.updated,
    ZEPHYRUS.col.status, ZEPHYRUS.col.reminders
  ]);
  zMigrateScheduleLayout_(scheduleSheet);
  zEnsureSheet_(ss, ZEPHYRUS.sheet.members, [
    ZEPHYRUS.memberCol.name, ZEPHYRUS.memberCol.email, ZEPHYRUS.memberCol.telegramId,
    ZEPHYRUS.memberCol.role, ZEPHYRUS.memberCol.active, ZEPHYRUS.memberCol.code,
    ZEPHYRUS.memberCol.calendarId
  ]);
  zEnsureSheet_(ss, ZEPHYRUS.sheet.settings, ['키', '값', '설명']);
  zEnsureSheet_(ss, ZEPHYRUS.sheet.log, ['일시', '구분', '일정ID', '채널', '결과', '상세']);
  zSetDefaultSettings_();
  zNotice_('시트 구조 점검이 끝났습니다. 기존 일정과 캘린더는 바꾸지 않았습니다.');
}

// Existing workbooks used "메모" in E.  Rename that header in place and put
// the new file column directly before 일정ID, without changing any row data.
function zMigrateScheduleLayout_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return;
  var map = zHeaders_(sheet);
  if (map['메모'] && !map[ZEPHYRUS.col.location]) {
    sheet.getRange(1, map['메모']).setValue(ZEPHYRUS.col.location);
  }
  map = zHeaders_(sheet);
  if (map[ZEPHYRUS.col.file]) return;
  var column = map[ZEPHYRUS.col.id] || sheet.getLastColumn() + 1;
  if (column <= sheet.getLastColumn()) sheet.insertColumnBefore(column);
  sheet.getRange(1, column).setValue(ZEPHYRUS.col.file).setFontWeight('bold').setBackground('#DCE6F1');
}

function zSetDefaultSettings_() {
  var sheet = zSpreadsheet_().getSheetByName(ZEPHYRUS.sheet.settings);
  var defaults = [
    ['즉시알림', 'Y', '새 일정 등록 시 텔레그램 발송 (Y/N)'],
    ['당일리마인드시각', '08:00', '당일 일정 리마인드 시각'],
    ['사전리마인드분', '30', '시간 일정의 사전 리마인드 분'],
    ['기본일정길이분', '60', '종료시간이 없을 때 일정 길이'],
    ['캘린더가져오기일수', '60', '처음 캘린더를 시트로 읽어올 기간']
  ];
  var existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  defaults.forEach(function(row) {
    if (existing.indexOf(row[0]) === -1) sheet.appendRow(row);
  });
}

function zHeaders_(sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var map = {};
  headers.forEach(function(value, index) {
    var name = String(value || '').trim();
    if (name) map[name] = index + 1;
  });
  map.__headers = headers.map(function(value) { return String(value || '').trim(); });
  return map;
}

function zSheet_(name) {
  var sheet = zSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('시트가 없습니다: ' + name + '. 먼저 “시트 구조 점검”을 실행하세요.');
  return sheet;
}

function zSettings_() {
  var settings = { '즉시알림': 'Y', '당일리마인드시각': '08:00', '사전리마인드분': '30', '기본일정길이분': '60', '캘린더가져오기일수': '60' };
  var sheet = zSheet_(ZEPHYRUS.sheet.settings);
  if (sheet.getLastRow() < 2) return settings;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(row) {
    if (row[0] !== '') settings[String(row[0])] = String(row[1] === '' ? settings[String(row[0])] || '' : row[1]);
  });
  return settings;
}

function zMembers_() {
  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  if (sheet.getLastRow() < 2) return [];
  var map = zHeaders_(sheet);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().map(function(row, index) {
    return {
      row: index + 2,
      name: String(row[(map[ZEPHYRUS.memberCol.name] || 1) - 1] || '').trim(),
      email: String(row[(map[ZEPHYRUS.memberCol.email] || 1) - 1] || '').trim(),
      telegramId: String(row[(map[ZEPHYRUS.memberCol.telegramId] || 1) - 1] || '').trim(),
      role: String(row[(map[ZEPHYRUS.memberCol.role] || 1) - 1] || '일반').trim(),
      active: String(row[(map[ZEPHYRUS.memberCol.active] || 1) - 1] || 'Y').trim().toUpperCase() !== 'N',
      code: String(row[(map[ZEPHYRUS.memberCol.code] || 1) - 1] || '').trim(),
      calendarId: String(row[(map[ZEPHYRUS.memberCol.calendarId] || 1) - 1] || '').trim()
    };
  }).filter(function(member) { return member.name && member.active; });
}

function zMemberByEmail_(email) {
  var wanted = String(email || '').trim().toLowerCase();
  if (!wanted) return null;
  return zMembers_().filter(function(member) {
    return String(member.email || '').trim().toLowerCase() === wanted;
  })[0] || null;
}

