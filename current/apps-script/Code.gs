/**
 * Zephyrus schedule recovery
 *
 * This is a safe replacement for the missing Apps Script project. It is meant
 * to be bound to an existing Zephyrus spreadsheet. It never creates a team
 * calendar automatically: TEAM_CAL_ID must be set in Script Properties first.
 */

var ZEPHYRUS = {
  // Public-source defaults deliberately contain no live resource IDs.
  // Configure SPREADSHEET_ID and TEAM_CAL_ID in Script Properties.
  recoverySpreadsheetId: '',
  recoveryTeamCalendarId: '',
  sheet: {
    schedule: '일정',
    members: '구성원',
    settings: '설정',
    log: '로그'
  },
  col: {
    date: '날짜',
    time: '시작시간',
    title: '일정',
    memo: '메모',
    targets: '대상자',
    alarm: '알람',
    id: '일정ID',
    registrant: '등록자',
    channel: '등록채널',
    // The old live sheet already uses this column.  Keeping it avoids making
    // a second set of calendar links while recovering the project.
    calendarEvent: '캘린더ID',
    updated: '수정일시',
    status: '상태',
    reminders: '리마인드'
  },
  memberCol: {
    name: '이름',
    email: '구글 계정',
    // The live sheet already has this header with a space.  Reuse it so
    // existing Telegram connections continue to work.
    telegramId: '텔레그램 ID',
    role: '권한',
    active: '사용여부',
    code: '인증코드',
    calendarId: '캘린더ID'
  },
  status: { normal: '정상', deleted: '삭제' },
  alarm: { pending: '대기', sent: '발송완료', failed: '실패' },
  kst: 'Asia/Seoul',
  markerPrefix: '[ZEPHYRUS:'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('제피로스 복구')
    .addItem('1. 시트 구조 점검', 'zPrepareSystem')
    .addItem('2. 연결 상태 확인', 'zShowStatus')
    .addSeparator()
    .addItem('시트 수정 감지 켜기', 'zInstallEditTrigger')
    .addItem('텔레그램 리마인드 켜기 (1분)', 'zInstallReminderTrigger')
    .addItem('시트에서 캘린더 동기화 켜기 (1분)', 'zInstallCalendarSyncTrigger')
    .addItem('캘린더에서 시트 반영 켜기', 'zInstallCalendarReverseTrigger')
    .addItem('자동 실행 4개 한꺼번에 켜기', 'zInstallCoreTriggers')
    .addSeparator()
    .addItem('지금 시트에서 캘린더로 동기화', 'syncCalendar_')
    .addItem('지금 리마인드 확인', 'runReminders')
    .addItem('인증코드 만들기', 'zGenerateAuthCodes')
    .addItem('텔레그램 웹훅 등록', 'zRegisterWebhook')
    .addSeparator()
    .addItem('직원 전용 캘린더 만들기/공유', 'zProvisionMemberCalendars')
    .addItem('직원 캘린더 양방향 동기화 켜기', 'zEnableMemberCalendarSync')
    .addItem('대상자 체크칸 만들기', 'zPrepareMemberCheckColumns')
    .addToUi();
}

function zProps_() {
  return PropertiesService.getScriptProperties();
}

function zProp_(key) {
  return zProps_().getProperty(key) || '';
}

function zRequireProp_(key) {
  var value = zProp_(key).trim();
  if (!value) {
    throw new Error('스크립트 속성 ' + key + ' 값이 비어 있습니다.');
  }
  return value;
}

function zSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  var id = zProp_('SPREADSHEET_ID') || ZEPHYRUS.recoverySpreadsheetId;
  if (!id) throw new Error('스크립트 속성 SPREADSHEET_ID 값이 비어 있습니다.');
  return SpreadsheetApp.openById(id);
}

function zTeamCalendarId_() {
  var id = zProp_('TEAM_CAL_ID') || ZEPHYRUS.recoveryTeamCalendarId;
  if (!id) throw new Error('스크립트 속성 TEAM_CAL_ID 값이 비어 있습니다.');
  return id;
}

function zTeamCalendar_() {
  var id = zTeamCalendarId_();
  var calendar = CalendarApp.getCalendarById(id);
  // The handover contained one stale calendar id.  If that id is no longer
  // visible, adopt only one calendar with the exact confirmed team name.
  // This never creates a calendar and refuses to guess when there are two.
  if (!calendar) {
    var matches = CalendarApp.getAllCalendars().filter(function(item) {
      return item.getName() === '제피로스 팀 일정';
    });
    // In the captured account, several old self-owned copies have the same
    // name.  The real team calendar is the single shared calendar under
    // "Other calendars", so prefer one exact-name calendar not owned by this
    // account.  We still refuse to guess if more than one shared copy exists.
    var sharedMatches = matches.filter(function(item) { return !item.isOwnedByMe(); });
    var candidate = sharedMatches.length === 1 ? sharedMatches[0] : (matches.length === 1 ? matches[0] : null);
    if (candidate) {
      calendar = candidate;
      zProps_().setProperty('TEAM_CAL_ID', calendar.getId());
    }
  }
  if (!calendar) {
    throw new Error('제피로스 팀 일정 캘린더에 접근할 수 없습니다. 새 캘린더를 만들지 않았습니다.');
  }
  return calendar;
}

function zNowText_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy-MM-dd HH:mm:ss');
}

// This recovery project may be created as a standalone project.  In that
// case there is no spreadsheet UI to show a popup in, so use the execution
// log instead of treating a completed setup as an error.
function zNotice_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (ignore) { console.log(message); }
}

function zToday_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy-MM-dd');
}

function zId_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyyMMddHHmmss') + '-' + Math.random().toString(36).slice(2, 7);
}

function zNormalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, ZEPHYRUS.kst, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (match) return zDateTextFromParts_(match[1], match[2], match[3]);

  // The original sheet contains date-like text such as "9월 16일" and
  // "9-20".  Treat these as dates in the current Korean calendar year so
  // old and new schedules can share one real date column and sort correctly.
  match = text.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/) ||
    text.match(/^(\d{1,2})[-./](\d{1,2})$/);
  if (!match) return '';
  return zDateTextFromParts_(Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy'), match[1], match[2]);
}

function zDateTextFromParts_(year, month, day) {
  var y = Number(year);
  var m = Number(month);
  var d = Number(day);
  // 03:00 UTC is noon in Korea, which avoids a date shifting at midnight.
  var date = new Date(Date.UTC(y, m - 1, d, 3, 0, 0));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return '';
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}

function zDateObject_(ymd) {
  if (ymd instanceof Date && !isNaN(ymd.getTime())) return ymd;
  var normalized = zNormalizeDate_(ymd);
  var parts = normalized.split('-').map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) return null;
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 3, 0, 0));
}

function zNormalizeTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, ZEPHYRUS.kst, 'HH:mm');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return '';
  var hour = Number(match[1]);
  var minute = Number(match[2]);
  if (hour > 23 || minute > 59) return '';
  return ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2);
}

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

function zPrepareSystem() {
  var ss = zSpreadsheet_();
  zEnsureSheet_(ss, ZEPHYRUS.sheet.schedule, [
    ZEPHYRUS.col.date, ZEPHYRUS.col.time, ZEPHYRUS.col.title, ZEPHYRUS.col.memo,
    ZEPHYRUS.col.targets, ZEPHYRUS.col.alarm, ZEPHYRUS.col.id, ZEPHYRUS.col.registrant,
    ZEPHYRUS.col.channel, ZEPHYRUS.col.calendarEvent, ZEPHYRUS.col.updated,
    ZEPHYRUS.col.status, ZEPHYRUS.col.reminders
  ]);
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

function zScheduleFromRow_(row, index, map, members) {
  function value(name) { return map[name] ? row[map[name] - 1] : ''; }
  return {
    row: index,
    date: zNormalizeDate_(value(ZEPHYRUS.col.date)),
    time: zNormalizeTime_(value(ZEPHYRUS.col.time)),
    title: String(value(ZEPHYRUS.col.title) || '').trim(),
    memo: String(value(ZEPHYRUS.col.memo) || '').trim(),
    // 대상자 칸에 이름을 쓰는 기존 방식과, 관리자용 사람별 체크칸을
    // 함께 지원한다. 직원은 이 시트를 사용하지 않아도 된다.
    targets: zTargetsFromRow_(row, map, members),
    alarm: String(value(ZEPHYRUS.col.alarm) || '').trim(),
    id: String(value(ZEPHYRUS.col.id) || '').trim(),
    registrant: String(value(ZEPHYRUS.col.registrant) || '').trim(),
    channel: String(value(ZEPHYRUS.col.channel) || '').trim(),
    calendarEvent: String(value(ZEPHYRUS.col.calendarEvent) || '').trim(),
    updated: String(value(ZEPHYRUS.col.updated) || '').trim(),
    status: String(value(ZEPHYRUS.col.status) || '').trim(),
    reminders: String(value(ZEPHYRUS.col.reminders) || '').trim()
  };
}

function zIsChecked_(value) {
  var mark = String(value || '').trim().toUpperCase();
  // The existing schedule sheet uses the letter O as its circle mark.
  // New checkbox cells use TRUE, so support both without changing old rows.
  return value === true || mark === 'TRUE' || mark === 'Y' || mark === 'O' || mark === '○';
}

function zTargetNames_(text) {
  var value = String(text || '').trim();
  if (value === '전부' || value === '전체') return value;
  return value.split(/[\s,]+/).filter(Boolean);
}

function zTargetsFromRow_(row, map, members) {
  var base = String(map[ZEPHYRUS.col.targets] ? row[map[ZEPHYRUS.col.targets] - 1] : '').trim();
  if (base === '전부' || base === '전체') return base;
  var names = zTargetNames_(base);
  // Reading the 구성원 sheet once per schedule made a simple Telegram
  // "오늘" query take tens of seconds.  A caller that already loaded the
  // roster passes it here, while one-off edit operations retain the fallback.
  (members || zMembers_()).forEach(function(member) {
    if (map[member.name] && zIsChecked_(row[map[member.name] - 1])) names.push(member.name);
  });
  return names.filter(function(name, index) { return names.indexOf(name) === index; }).join(', ');
}

// 이 메뉴는 구성원 이름을 제목으로 한 체크칸을 일정 시트 끝에 만든다.
// 체크된 사람만 해당 일정의 전용 캘린더에 복사된다.
function zPrepareMemberCheckColumns() {
  var members = zMembers_();
  if (!members.length) throw new Error('구성원 시트에 먼저 이름과 구글 계정을 입력하세요.');
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  var added = [];
  members.forEach(function(member) {
    if (map[member.name]) return;
    var column = sheet.getLastColumn() + 1;
    sheet.getRange(1, column).setValue(member.name).setFontWeight('bold').setBackground('#E2F0D9');
    sheet.getRange(2, column, Math.max(1, sheet.getMaxRows() - 1), 1).insertCheckboxes();
    map[member.name] = column;
    added.push(member.name);
  });
  zNotice_(added.length ? '대상자 체크칸을 만들었습니다: ' + added.join(', ') : '대상자 체크칸이 이미 준비되어 있습니다.');
}

function zWriteScheduleMemberChecks_(rowNumber, targetText) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  var targets = zTargetNames_(targetText);
  zMembers_().forEach(function(member) {
    if (!map[member.name]) return;
    var checked = targets === '전부' || targets === '전체' || targets.indexOf(member.name) !== -1;
    sheet.getRange(rowNumber, map[member.name]).setValue(checked);
  });
}

function zSchedules_(members) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  if (sheet.getLastRow() < 2) return [];
  var map = zHeaders_(sheet);
  var roster = members || zMembers_();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().map(function(row, index) {
    return zScheduleFromRow_(row, index + 2, map, roster);
  }).filter(function(schedule) { return schedule.date || schedule.title || schedule.id; });
}

function zSetSchedule_(rowNumber, name, value) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  if (!map[name]) throw new Error('일정 시트에 열이 없습니다: ' + name);
  sheet.getRange(rowNumber, map[name]).setValue(value);
}

// Keep dates as real date values, but display them in the familiar Korean
// format.  Sorting the complete row range keeps every schedule's notes,
// targets, IDs and check marks together.
function zNormalizeAndSortScheduleSheet_() {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var map = zHeaders_(sheet);
  var rowCount = lastRow - 1;
  var dateRange = sheet.getRange(2, map[ZEPHYRUS.col.date], rowCount, 1);
  var dates = dateRange.getValues();
  var normalizedCount = 0;
  var values = dates.map(function(row) {
    var before = row[0];
    var normalized = zNormalizeDate_(before);
    if (!normalized) return [before];
    var date = zDateObject_(normalized);
    if (zNormalizeDate_(before) !== Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd') || !(before instanceof Date)) normalizedCount++;
    return [date];
  });
  dateRange.setValues(values).setNumberFormat('m"월" d"일"');

  var sortBy = [{ column: map[ZEPHYRUS.col.date], ascending: true }];
  if (map[ZEPHYRUS.col.time]) sortBy.push({ column: map[ZEPHYRUS.col.time], ascending: true });
  sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).sort(sortBy);
  return normalizedCount;
}

function zNormalizeAndSortSchedules() {
  var count = zNormalizeAndSortScheduleSheet_();
  zNotice_(count + '개 날짜를 통일하고 일정표를 날짜순으로 정렬했습니다.');
}

function zFindScheduleById_(id) {
  return zSchedules_().filter(function(schedule) { return schedule.id === id; })[0] || null;
}

function zFindScheduleByCalendarEvent_(eventId) {
  return zSchedules_().filter(function(schedule) {
    var ids = zParseCalendarIds_(schedule.calendarEvent);
    return Object.keys(ids).some(function(key) { return ids[key] === String(eventId || ''); });
  })[0] || null;
}

// Existing Zephyrus rows store links like "팀:event-id;홍길동:event-id".
// Older rows can also contain just one event id.  We only manage the team
// part here and deliberately leave any personal-calendar links untouched.
function zParseCalendarIds_(value) {
  var text = String(value || '').trim();
  var ids = {};
  if (!text) return ids;
  if (text.indexOf(';') === -1 && text.indexOf(':') === -1) {
    ids.팀 = text;
    return ids;
  }
  text.split(';').forEach(function(part) {
    var item = String(part || '').trim();
    if (!item) return;
    var at = item.indexOf(':');
    if (at < 0) {
      ids.팀 = item;
      return;
    }
    var key = item.slice(0, at).trim();
    var id = item.slice(at + 1).trim();
    if (key && id) ids[key] = id;
  });
  return ids;
}

function zTeamEventId_(schedule) {
  return zCalendarEventId_(schedule, '팀') || zCalendarEventId_(schedule, 'TEAM') || zCalendarEventId_(schedule, 'team');
}

function zCalendarEventId_(schedule, key) {
  var ids = zParseCalendarIds_(schedule.calendarEvent);
  return String(ids[key] || '');
}

function zCalendarCellWithId_(schedule, key, eventId) {
  var ids = zParseCalendarIds_(schedule.calendarEvent);
  if (eventId) ids[key] = String(eventId); else delete ids[key];
  var keys = Object.keys(ids).filter(function(item) { return ids[item]; });
  if (keys.length === 1 && keys[0] === '팀') return ids.팀;
  return keys.map(function(item) { return item + ':' + ids[item]; }).join(';');
}

function zCalendarCellWithTeamId_(schedule, eventId) {
  return zCalendarCellWithId_(schedule, '팀', eventId);
}

function zSetMemberCalendarId_(member, calendarId) {
  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  var map = zHeaders_(sheet);
  if (!map[ZEPHYRUS.memberCol.calendarId]) throw new Error('구성원 시트에 캘린더ID 열이 없습니다. 시트 구조 점검을 먼저 실행하세요.');
  sheet.getRange(member.row, map[ZEPHYRUS.memberCol.calendarId]).setValue(calendarId || '');
}

function zMemberCalendar_(member) {
  if (!member || !member.calendarId) return null;
  return CalendarApp.getCalendarById(member.calendarId) || null;
}

function zMemberCalendarKey_(member) {
  return String(member && member.name || '').trim();
}

function zAppendSchedule_(input) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  var row = new Array(sheet.getLastColumn()).fill('');
  var id = input.id || zId_();
  var date = zDateObject_(input.date);
  row[map[ZEPHYRUS.col.date] - 1] = date || input.date;
  row[map[ZEPHYRUS.col.time] - 1] = input.time || '';
  row[map[ZEPHYRUS.col.title] - 1] = input.title || '';
  row[map[ZEPHYRUS.col.memo] - 1] = input.memo || '';
  row[map[ZEPHYRUS.col.targets] - 1] = input.targets || '';
  row[map[ZEPHYRUS.col.alarm] - 1] = input.alarm || ZEPHYRUS.alarm.pending;
  row[map[ZEPHYRUS.col.id] - 1] = id;
  row[map[ZEPHYRUS.col.registrant] - 1] = input.registrant || '시스템';
  row[map[ZEPHYRUS.col.channel] - 1] = input.channel || '시트';
  row[map[ZEPHYRUS.col.calendarEvent] - 1] = input.calendarEvent || '';
  row[map[ZEPHYRUS.col.updated] - 1] = zNowText_();
  row[map[ZEPHYRUS.col.status] - 1] = input.status || ZEPHYRUS.status.normal;
  row[map[ZEPHYRUS.col.reminders] - 1] = input.reminders || '';
  sheet.appendRow(row);
  var number = sheet.getLastRow();
  sheet.getRange(number, map[ZEPHYRUS.col.date]).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(number, map[ZEPHYRUS.col.time]).setNumberFormat('@');
  zWriteScheduleMemberChecks_(number, input.targets || input.registrant || '');
  if (!input.deferSort) zNormalizeAndSortScheduleSheet_();
  return zFindScheduleById_(id);
}

function zLog_(kind, scheduleId, channel, result, detail) {
  try {
    zSheet_(ZEPHYRUS.sheet.log).appendRow([zNowText_(), kind, scheduleId || '', channel || '', result || '', String(detail || '').slice(0, 500)]);
  } catch (error) {
    console.log('log failure: ' + error);
  }
}

function zRecipients_(schedule) {
  var members = zMembers_();
  var target = String(schedule.targets || '').trim();
  if (target === '전부' || target === '전체') return members;
  var names = target.split(/[\s,]+/).filter(Boolean);
  var found = members.filter(function(member) { return names.indexOf(member.name) !== -1; });
  if (found.length) return found;
  return members.filter(function(member) { return member.name === schedule.registrant; });
}

function zScheduleText_(schedule) {
  return '[' + schedule.date + (schedule.time ? ' ' + schedule.time : ' 종일') + '] ' + schedule.title + (schedule.memo ? '\n메모: ' + schedule.memo : '');
}

function zMemberCalendarSyncEnabled_() {
  return zProp_('MEMBER_CALENDAR_SYNC_ENABLED') === 'Y';
}

function zScheduleCalendarRecipients_(schedule) {
  var byName = {};
  zRecipients_(schedule).forEach(function(member) { byName[member.name] = member; });
  // A writer should always receive their own copy, even when an administrator
  // changes the extra viewers in the 대상자 cell.
  zMembers_().forEach(function(member) {
    if (member.name === schedule.registrant) byName[member.name] = member;
  });
  return Object.keys(byName).map(function(name) { return byName[name]; })
    .filter(function(member) { return member.calendarId; });
}

function zProvisionMemberCalendars() {
  var created = [];
  var missingEmail = [];
  var inaccessible = [];
  zMembers_().forEach(function(member) {
    if (member.calendarId) {
      if (zMemberCalendar_(member)) return;
      inaccessible.push(member.name);
      return;
    }
    if (!member.email) {
      missingEmail.push(member.name);
      return;
    }
    var calendar = CalendarApp.createCalendar('제피로스 일정 - ' + member.name, {
      timeZone: ZEPHYRUS.kst,
      description: '제피로스 업무 일정 전용 캘린더. 직원 개인 일정과 분리해 사용합니다.'
    });
    calendar.addEditor(member.email);
    zSetMemberCalendarId_(member, calendar.getId());
    created.push(member.name);
  });
  if (missingEmail.length) throw new Error('구글 계정이 비어 있어 전용 캘린더를 만들 수 없는 구성원: ' + missingEmail.join(', '));
  if (inaccessible.length) throw new Error('기존 캘린더ID에는 접근할 수 없습니다: ' + inaccessible.join(', ') + '. 상태 확인에서 기존 캘린더를 확인한 뒤에만 해당 ID를 비우세요.');
  zNotice_(created.length ? created.join(', ') + ' 전용 캘린더를 만들고 해당 직원에게 공유했습니다.' : '새로 만들 전용 캘린더가 없습니다.');
}

function zEnableMemberCalendarSync() {
  var missing = zMembers_().filter(function(member) { return !member.calendarId || !zMemberCalendar_(member); }).map(function(member) { return member.name; });
  if (missing.length) throw new Error('전용 캘린더가 없는 구성원: ' + missing.join(', ') + '. 먼저 zProvisionMemberCalendars를 실행하세요.');
  zProps_().setProperty('MEMBER_CALENDAR_SYNC_ENABLED', 'Y');
  syncCalendar_();
  zNotice_('직원 전용 캘린더 양방향 동기화를 켰습니다. 이제 약 1분 안에 개인 캘린더 변경도 반영됩니다.');
}

function zMemberCalendarSyncKey_(member) {
  return 'MEMBER_CAL_SYNC_' + String(member.calendarId || member.name).replace(/[^a-zA-Z0-9]/g, '_').slice(-160);
}

function zApplyMemberCalendarApiEvent_(member, event) {
  var key = zMemberCalendarKey_(member);
  var scheduleId = zMarkerScheduleId_(event.description);
  var schedule = scheduleId ? zFindScheduleById_(scheduleId) : null;
  if (!schedule) schedule = zFindScheduleByCalendarEvent_(String(event.id || event.iCalUID || ''));

  if (event.status === 'cancelled') {
    if (schedule && schedule.registrant === member.name) {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.status, ZEPHYRUS.status.deleted);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    }
    return;
  }

  var date = zApiDate_(event);
  if (!date || !event.summary) return;
  var ymd = Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd');
  var time = zApiTime_(event);
  var eventId = String(event.id || event.iCalUID || '');

  if (schedule) {
    // Only the owner of a schedule may change its shared source record from
    // their personal calendar.  A recipient's copied event is read-only to
    // the synchronizer, even if Google granted them edit access.
    if (schedule.registrant === member.name || schedule.channel === '개인캘린더') {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.date, date);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.time, time);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.title, String(event.summary));
      zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    }
    if (zCalendarEventId_(schedule, key) !== eventId) {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithId_(schedule, key, eventId));
    }
    return;
  }

  // A new event entered directly in the employee's dedicated calendar is a
  // new private work schedule.  The next push creates its manager-team copy.
  zAppendSchedule_({
    date: ymd,
    time: time,
    title: String(event.summary),
    memo: String(event.description || '').replace(/\[(?:ZEPHYRUS|일정ID):[^\]]+\]/g, '').trim(),
    targets: member.name,
    registrant: member.name,
    channel: '개인캘린더',
    calendarEvent: key + ':' + eventId
  });
}

function zPullOneMemberCalendar_(member) {
  if (!member.calendarId) return;
  var props = zProps_();
  var propKey = zMemberCalendarSyncKey_(member);
  var token = props.getProperty(propKey);
  var nextToken = null;
  for (var attempt = 0; attempt < 2; attempt++) {
    var pageToken = null;
    nextToken = null;
    try {
      do {
        var options = { maxResults: 100, singleEvents: true, showDeleted: true };
        if (pageToken) options.pageToken = pageToken;
        if (token) options.syncToken = token;
        if (!token) {
          var days = Number(zSettings_()['캘린더가져오기일수'] || 60) || 60;
          options.timeMin = new Date(Date.now() - days * 86400000).toISOString();
        }
        var response = Calendar.Events.list(member.calendarId, options);
        (response.items || []).forEach(function(event) { zApplyMemberCalendarApiEvent_(member, event); });
        pageToken = response.nextPageToken;
        if (response.nextSyncToken) nextToken = response.nextSyncToken;
      } while (pageToken);
      break;
    } catch (error) {
      if (token && String(error).indexOf('410') !== -1 && attempt === 0) {
        props.deleteProperty(propKey);
        token = '';
        continue;
      }
      throw error;
    }
  }
  if (nextToken) props.setProperty(propKey, nextToken);
}

function zPullMemberCalendars_() {
  zMembers_().forEach(function(member) {
    if (!member.calendarId) return;
    try { zPullOneMemberCalendar_(member); }
    catch (error) { zLog_('오류', '', '개인캘린더', '실패', member.name + ': ' + error); }
  });
}

function zSyncScheduleToMemberCalendars_(schedule) {
  var recipients = zScheduleCalendarRecipients_(schedule);
  var intended = {};
  recipients.forEach(function(member) { intended[zMemberCalendarKey_(member)] = member; });
  var allMembers = zMembers_().filter(function(member) { return member.calendarId; });
  var current = schedule;

  allMembers.forEach(function(member) {
    var key = zMemberCalendarKey_(member);
    var priorId = zCalendarEventId_(current, key);
    var calendar = zMemberCalendar_(member);
    if (!calendar) {
      zLog_('오류', schedule.id, '개인캘린더', '건너뜀', member.name + ' 전용 캘린더에 접근할 수 없습니다.');
      return;
    }
    if (!intended[key]) {
      if (priorId) {
        try { calendar.getEventById(priorId).deleteEvent(); } catch (ignore) {}
        var removed = zCalendarCellWithId_(current, key, '');
        zSetSchedule_(current.row, ZEPHYRUS.col.calendarEvent, removed);
        current = zFindScheduleById_(current.id) || current;
      }
      return;
    }
    var eventId = zUpsertCalendarEvent_(calendar, current, priorId);
    if (eventId && eventId !== priorId) {
      var updated = zCalendarCellWithId_(current, key, eventId);
      zSetSchedule_(current.row, ZEPHYRUS.col.calendarEvent, updated);
      current = zFindScheduleById_(current.id) || current;
    }
  });
}

function zDeleteScheduleFromMemberCalendars_(schedule) {
  zMembers_().forEach(function(member) {
    if (!member.calendarId) return;
    var eventId = zCalendarEventId_(schedule, zMemberCalendarKey_(member));
    if (!eventId) return;
    try {
      var calendar = zMemberCalendar_(member);
      if (calendar) calendar.getEventById(eventId).deleteEvent();
    } catch (ignore) {}
  });
}

function zTelegramToken_() {
  return zRequireProp_('BOT_TOKEN');
}

function zTelegramCall_(method, payload) {
  var lastError;
  // Retrying a send can create duplicate chat messages when Telegram got the
  // first request but Apps Script lost the response.  Only retry idempotent
  // setup/read operations.
  var attempts = method === 'sendMessage' ? 1 : 3;
  for (var attempt = 0; attempt < attempts; attempt++) {
    try {
      var response = UrlFetchApp.fetch('https://api.telegram.org/bot' + zTelegramToken_() + '/' + method, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      var body = JSON.parse(response.getContentText() || '{}');
      if (body.ok) return body;
      lastError = new Error('Telegram ' + method + ' 실패: ' + response.getContentText());
      // These are permanent request/permission errors, so another attempt
      // would not help.
      if (response.getResponseCode() === 400 || response.getResponseCode() === 403) break;
    } catch (error) {
      lastError = error;
    }
    if (attempt < attempts - 1) Utilities.sleep(600 * (attempt + 1));
  }
  throw lastError || new Error('Telegram ' + method + ' 호출에 실패했습니다.');
}

function zTelegramChunks_(text) {
  var limit = 3600; // Telegram allows 4096 characters; leave room for safety.
  var lines = String(text || '').split('\n');
  var chunks = [];
  var current = '';
  lines.forEach(function(line) {
    var rest = String(line);
    while (rest.length > limit) {
      if (current) {
        chunks.push(current);
        current = '';
      }
      chunks.push(rest.slice(0, limit));
      rest = rest.slice(limit);
    }
    var candidate = current ? current + '\n' + rest : rest;
    if (candidate.length > limit && current) {
      chunks.push(current);
      current = rest;
    } else {
      current = candidate;
    }
  });
  if (current || !chunks.length) chunks.push(current);
  return chunks.filter(function(chunk) { return chunk !== ''; });
}

function zSendTelegram_(chatId, text) {
  var result = null;
  zTelegramChunks_(text).forEach(function(chunk) {
    result = zTelegramCall_('sendMessage', { chat_id: String(chatId), text: chunk });
  });
  return result;
}

function zNotifySchedule_(schedule, title, skipTelegramId) {
  var sent = [];
  var failed = [];
  var recipients = {};
  zRecipients_(schedule).forEach(function(member) { recipients[member.name] = member; });
  // Administrators receive an automatic record of every schedule; ordinary
  // members receive only their own or explicitly shared schedules.
  zMembers_().forEach(function(member) {
    if (member.role === '관리자') recipients[member.name] = member;
  });
  Object.keys(recipients).map(function(name) { return recipients[name]; }).forEach(function(member) {
    if (!member.telegramId) return;
    if (skipTelegramId && String(member.telegramId) === String(skipTelegramId)) return;
    try {
      zSendTelegram_(member.telegramId, title + '\n' + zScheduleText_(schedule));
      sent.push(member.name);
    } catch (error) {
      failed.push(member.name);
      zLog_('오류', schedule.id, '텔레그램', '실패', error);
    }
  });
  return { sent: sent, failed: failed };
}

function zGenerateAuthCodes() {
  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  var map = zHeaders_(sheet);
  var count = 0;
  zMembers_().forEach(function(member) {
    if (member.telegramId || member.code) return;
    sheet.getRange(member.row, map[ZEPHYRUS.memberCol.code]).setValue(String(Math.floor(100000 + Math.random() * 900000)));
    count++;
  });
  zNotice_(count + '명에게 인증코드를 만들었습니다.');
}

// Normally Telegram calls the Apps Script URL directly.  When the optional
// immediate-response proxy is enabled, TELEGRAM_WEBHOOK_URL holds that public
// Worker URL instead.  The schedule data still stays in Apps Script/Sheets.
function zTelegramWebhookUrl_() {
  return String(zProp_('TELEGRAM_WEBHOOK_URL') || zRequireProp_('WEB_APP_URL')).replace(/\/$/, '');
}

function zRegisterWebhook() {
  var baseUrl = zTelegramWebhookUrl_();
  var key = zProp_('WEBHOOK_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    zProps_().setProperty('WEBHOOK_KEY', key);
  }
  zTelegramCall_('setWebhook', { url: baseUrl + '?key=' + encodeURIComponent(key), allowed_updates: ['message'] });
  zNotice_('텔레그램 웹훅을 등록했습니다.');
}

// Recovery diagnostic: checks the outbound Telegram connection without
// creating an event, changing a calendar, or relying on a webhook update.
function zTestTelegramReply() {
  var member = zMembers_().filter(function(item) {
    return item.telegramId;
  })[0] || null;
  if (!member) throw new Error('연결된 구성원의 텔레그램 ID를 찾지 못했습니다.');
  try {
    zSendTelegram_(member.telegramId, '[제피로스 점검] 텔레그램 발송 연결이 정상입니다.');
    zLog_('점검', '', '텔레그램', '성공', member.name + ' 텔레그램 직접 발송 성공');
    zNotice_(member.name + ' 텔레그램으로 점검 문구를 보냈습니다.');
  } catch (error) {
    zLog_('점검', '', '텔레그램', '실패', member.name + ' 직접 발송 오류: ' + error);
    throw error;
  }
}

// One-time fallback for linking an account when Telegram has queued old
// /start updates.  This does not change any calendar or schedule data.
function zBeginTelegramManualLink() {
  zTelegramCall_('deleteWebhook', { drop_pending_updates: true });
  zGenerateAuthCodes();
  zNotice_('이제 텔레그램 봇에 시트의 6자리 인증번호를 한 번 보내고, zFinishTelegramManualLink를 실행하세요.');
}

function zFinishTelegramManualLink() {
  var updates = zTelegramCall_('getUpdates', { timeout: 0 }).result || [];
  var members = zMembers_();
  var found = null;
  updates.forEach(function(update) {
    var message = update && update.message;
    var code = message ? String(message.text || '').replace(/[^0-9]/g, '') : '';
    if (!message || !message.chat || !/^\d{6}$/.test(code)) return;
    var member = members.filter(function(item) {
      return String(item.code || '').replace(/[^0-9]/g, '') === code;
    })[0] || null;
    if (member) found = { update: update, message: message, member: member };
  });
  if (!found) throw new Error('방금 보낸 인증번호를 찾지 못했습니다. 먼저 zBeginTelegramManualLink를 실행한 뒤 번호를 한 번 보내세요.');

  var sheet = zSheet_(ZEPHYRUS.sheet.members);
  var map = zHeaders_(sheet);
  sheet.getRange(found.member.row, map[ZEPHYRUS.memberCol.telegramId]).setValue(String(found.message.chat.id));
  sheet.getRange(found.member.row, map[ZEPHYRUS.memberCol.code]).setValue('');
  zTelegramCall_('getUpdates', { offset: Number(found.update.update_id) + 1, timeout: 0 });
  zRegisterWebhook();
  zSendTelegram_(String(found.message.chat.id), found.member.name + '님 연결이 완료되었습니다.');
  zLog_('인증', '', '텔레그램', '성공', found.member.name + ' 텔레그램 연결 완료');
  zNotice_('텔레그램 연결을 완료했습니다.');
}

function doGet() {
  return ContentService.createTextOutput('Zephyrus webhook is running.');
}

// Kept for backward compatibility with earlier deployments.  The active
// webhook flow below sends the reply through the Bot API instead: Apps Script
// ContentService uses a redirect, and carrying Telegram's inline reply through
// that redirect can intermittently make Telegram retry the same update.
function zWebhookTelegramReply_(chatId, text) {
  return ContentService.createTextOutput(JSON.stringify({
    method: 'sendMessage',
    chat_id: String(chatId),
    text: String(text || '')
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(event) {
  try {
    var suppliedKey = event && event.parameter ? String(event.parameter.key || '') : '';
    if (!suppliedKey || suppliedKey !== zRequireProp_('WEBHOOK_KEY')) {
      zLog_('수신', '', '텔레그램', '거부', '웹훅 키가 없거나 일치하지 않습니다.');
      return ContentService.createTextOutput('forbidden');
    }
    var update = JSON.parse((event.postData && event.postData.contents) || '{}');
    var message = update.message;
    if (!message || !message.chat) return ContentService.createTextOutput('ok');
    var updateId = String(update.update_id || '');
    var lastUpdateId = Number(zProp_('LAST_TELEGRAM_UPDATE_ID') || 0);
    var numericUpdateId = Number(updateId || 0);
    // Telegram update IDs increase in order.  Store the last completed one
    // only after the handler finishes; unlike a temporary cache this cannot
    // make a freshly sent message look like an old duplicate.
    if (numericUpdateId && numericUpdateId <= lastUpdateId) {
      zLog_('수신', '', '텔레그램', '중복', '이미 처리한 업데이트: ' + updateId);
      return ContentService.createTextOutput('ok');
    }
    var chatId = String(message.chat.id);
    zLog_('수신', '', '텔레그램', '시작', '업데이트 수신: ' + (updateId || '번호 없음'));
    // Generate the reply first, then store the completed update ID before
    // sending.  This prevents a network retry from registering one schedule
    // twice, while the actual chat reply is delivered through Telegram's Bot
    // API rather than the fragile inline-webhook response path.
    var replyText = zHandleTelegramMessage_(chatId, String(message.text || '').trim(), true);
    if (numericUpdateId) zProps_().setProperty('LAST_TELEGRAM_UPDATE_ID', String(numericUpdateId));
    zLog_('수신', '', '텔레그램', '성공', '업데이트 처리 완료: ' + (updateId || '번호 없음'));
    if (replyText) {
      try {
        zSendTelegram_(chatId, replyText);
      } catch (sendError) {
        // Return 200 to stop Telegram from re-sending a schedule-registration
        // update.  The user can safely send a new query if a rare send error
        // occurs, and the failure is retained in the log for checking.
        zLog_('오류', '', '텔레그램', '실패', '답장 발신: ' + sendError);
      }
    }
    return ContentService.createTextOutput('ok');
  } catch (error) {
    zLog_('오류', '', '텔레그램', '실패', '웹훅: ' + error);
    return ContentService.createTextOutput('error');
  }
}

function zIsFullViewer_(member) {
  return member && (member.role === '관리자' || member.role === '전체열람');
}

function zIsScheduleVisibleToMember_(schedule, member) {
  if (zIsFullViewer_(member)) return true;
  if (schedule.registrant === member.name) return true;
  var target = String(schedule.targets || '').trim();
  if (target === '전부' || target === '전체') return true;
  return target.split(/[\s,]+/).filter(Boolean).indexOf(member.name) !== -1;
}

function zVisibleSchedulesForMember_(member, date, members) {
  return zSchedules_(members).filter(function(schedule) {
    return schedule.status !== ZEPHYRUS.status.deleted && (!date || schedule.date === date) && zIsScheduleVisibleToMember_(schedule, member);
  });
}

function zTelegramTime_(period, hour, minute) {
  if (hour === '' || hour === undefined || hour === null) return '';
  var h = Number(hour);
  var m = Number(minute || 0);
  if (period === '오후' && h < 12) h += 12;
  if (period === '오전' && h === 12) h = 0;
  // In ordinary Korean work-schedule input, an unqualified 1~7시는 usually
  // intended as afternoon.  오전/오후 or 24-hour input remains explicit.
  if (!period && h > 0 && h < 8) h += 12;
  return zNormalizeTime_(('0' + h).slice(-2) + ':' + ('0' + m).slice(-2));
}

function zParseTelegramSchedule_(text) {
  var source = String(text || '').trim().replace(/^일정\s*/i, '');
  var full = source.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:\s+(오전|오후)?\s*(\d{1,2})(?::(\d{2}))?\s*시?)?\s+(.+)$/);
  if (full) {
    return {
      date: zNormalizeDate_(full[1] + '-' + full[2] + '-' + full[3]),
      time: zTelegramTime_(full[4] || '', full[5], full[6]),
      title: String(full[7] || '').trim()
    };
  }
  var short = source.match(/^(\d{1,2})(?:\s*월\s*|[/.])(\d{1,2})\s*(?:일)?(?:\s+(오전|오후)?\s*(\d{1,2})(?::(\d{2}))?\s*시?)?\s+(.+)$/);
  if (short) {
    return {
      date: zNormalizeDate_(new Date().getFullYear() + '-' + short[1] + '-' + short[2]),
      time: zTelegramTime_(short[3] || '', short[4], short[5]),
      title: String(short[6] || '').trim()
    };
  }
  return null;
}

function zHandleTelegramMessage_(chatId, text, replyViaWebhook) {
  text = String(text || '').replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
  var compactText = text.normalize('NFC').replace(/\s+/g, '');
  var members = zMembers_();
  var member = members.filter(function(item) { return item.telegramId === chatId; })[0] || null;
  function reply(message) {
    if (replyViaWebhook) return String(message);
    zSendTelegram_(chatId, message);
    return '';
  }
  if (text === '/start') {
    return reply('제피로스 일정관리입니다. 받은 6자리 인증코드를 입력하세요.');
  }
  var enteredCode = text.replace(/[^0-9]/g, '');
  if (/^\d{6}$/.test(enteredCode)) {
    var match = members.filter(function(item) {
      return String(item.code || '').replace(/[^0-9]/g, '') === enteredCode;
    })[0] || null;
    if (!match) {
      return reply('인증코드를 찾지 못했습니다.');
    }
    var sheet = zSheet_(ZEPHYRUS.sheet.members);
    var map = zHeaders_(sheet);
    sheet.getRange(match.row, map[ZEPHYRUS.memberCol.telegramId]).setValue(chatId);
    sheet.getRange(match.row, map[ZEPHYRUS.memberCol.code]).setValue('');
    zLog_('인증', '', '텔레그램', '성공', match.name + ' 텔레그램 연결 완료');
    return reply(match.name + '님 연결이 완료되었습니다.');
  }
  if (!member) {
    return reply('/start 후 인증코드를 먼저 입력하세요.');
  }
  // Telegram clients can retain an invisible formatting character even when
  // the bubble visibly says only "오늘".  Treat any otherwise ordinary text
  // containing 오늘 as the today-query command.
  if (compactText === '오늘일정' || compactText === '오늘' || text.indexOf('오늘') !== -1) {
    var items = zVisibleSchedulesForMember_(member, zToday_(), members);
    zLog_('조회', '', '텔레그램', '성공', member.name + ' / 오늘 / ' + items.length + '건');
    return reply(items.length ? items.map(zScheduleText_).join('\n\n') : '오늘 일정이 없습니다.');
  }
  var add = zParseTelegramSchedule_(text);
  if (add && add.date && add.title) {
    var schedule = zAppendSchedule_({
      date: add.date, time: add.time, title: add.title,
      targets: member.name, registrant: member.name, channel: '텔레그램'
    });
    // The sender receives the clear "등록했습니다" reply below.  Do not also
    // send the notification copy to the same chat, which looked like a
    // duplicate response in Telegram.
    zNotifySchedule_(schedule, '[새 일정]', chatId);
    return reply('등록했습니다.\n' + zScheduleText_(schedule));
  }
  return reply("입력 예: 9/15 2시 삼겹살 먹는 날\n또는: 일정 2026-09-15 14:00 회의\n조회: '오늘' 입력");
}

function zEventDescription_(schedule) {
  var lines = [];
  if (schedule.memo) lines.push(schedule.memo);
  if (schedule.targets) lines.push('대상: ' + schedule.targets);
  if (schedule.registrant) lines.push('등록: ' + schedule.registrant);
  // Keep both labels during recovery.  The second label is what the old
  // project wrote, so an existing event can be found even after its cell link
  // was accidentally cleared.
  lines.push('[일정ID:' + schedule.id + ']');
  lines.push(ZEPHYRUS.markerPrefix + schedule.id + ']');
  return lines.join('\n');
}

function zEventTimes_(schedule) {
  var day = zDateObject_(schedule.date);
  if (!day) throw new Error('날짜 형식 오류: ' + schedule.date);
  if (!schedule.time) return { allDay: true, start: day, end: null };
  var parts = schedule.time.split(':').map(Number);
  var start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), parts[0], parts[1]);
  var minutes = Number(zSettings_()['기본일정길이분'] || 60) || 60;
  return { allDay: false, start: start, end: new Date(start.getTime() + minutes * 60000) };
}

function zSameEventTime_(event, times) {
  if (event.isAllDayEvent() !== times.allDay) return false;
  if (times.allDay) {
    return zNormalizeDate_(event.getAllDayStartDate()) === zNormalizeDate_(times.start);
  }
  return event.getStartTime().getTime() === times.start.getTime() && event.getEndTime().getTime() === times.end.getTime();
}

function zFindExistingCalendarEvent_(calendar, schedule, times) {
  var events;
  try { events = calendar.getEventsForDay(times.start); } catch (ignore) { return null; }
  var tagged = null;
  var same = null;
  events.forEach(function(candidate) {
    var description = candidate.getDescription();
    if (zMarkerScheduleId_(description) === schedule.id) {
      if (!tagged) tagged = candidate;
      return;
    }
    if (candidate.getTitle() === schedule.title && zSameEventTime_(candidate, times) && !same) same = candidate;
  });
  return tagged || same;
}

function zUpdateCalendarEvent_(event, schedule, times, description) {
  var changed = false;
  if (event.getTitle() !== schedule.title) {
    event.setTitle(schedule.title);
    changed = true;
  }
  if (event.getDescription() !== description) {
    event.setDescription(description);
    changed = true;
  }
  if (!zSameEventTime_(event, times)) {
    if (times.allDay) event.setAllDayDate(times.start); else event.setTime(times.start, times.end);
    changed = true;
  }
  return changed;
}

function zUpsertCalendarEvent_(calendar, schedule, knownEventId) {
  var times = zEventTimes_(schedule);
  var event = null;
  var eventId = arguments.length >= 3 ? String(knownEventId || '') : zTeamEventId_(schedule);
  if (eventId) {
    try { event = calendar.getEventById(eventId); } catch (ignore) { event = null; }
  }
  var description = zEventDescription_(schedule);
  if (event) {
    try {
      zUpdateCalendarEvent_(event, schedule, times, description);
    } catch (error) {
      // Some old events were created by another account and cannot be edited
      // even though the calendar itself is shared.  Keep their link; never
      // create a duplicate and never let that one event stop every schedule.
      zLog_('오류', schedule.id, '캘린더', '건너뜀', '수정 불가 이벤트: ' + error);
      return '';
    }
  } else if (event = zFindExistingCalendarEvent_(calendar, schedule, times)) {
    try {
      zUpdateCalendarEvent_(event, schedule, times, description);
    } catch (error2) {
      zLog_('오류', schedule.id, '캘린더', '건너뜀', '수정 불가 이벤트: ' + error2);
      return '';
    }
  } else if (times.allDay) {
    event = calendar.createAllDayEvent(schedule.title, times.start, { description: description });
  } else {
    event = calendar.createEvent(schedule.title, times.start, times.end, { description: description });
  }
  return event.getId();
}

function syncCalendar_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    if (zMemberCalendarSyncEnabled_()) zPullMemberCalendars_();
    var calendar = zTeamCalendar_();
    zSchedules_().forEach(function(schedule) {
      if (!schedule.id) return;
      if (schedule.status === ZEPHYRUS.status.deleted) {
        if (zTeamEventId_(schedule)) {
          try { calendar.getEventById(zTeamEventId_(schedule)).deleteEvent(); } catch (ignore) {}
          zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, ''));
        }
        if (zMemberCalendarSyncEnabled_()) zDeleteScheduleFromMemberCalendars_(schedule);
        return;
      }
      if (!schedule.date || !schedule.title) return;
      var eventId = zUpsertCalendarEvent_(calendar, schedule);
      if (!eventId) return;
      if (eventId !== zTeamEventId_(schedule)) {
        zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, eventId));
        zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
        schedule = zFindScheduleById_(schedule.id) || schedule;
      }
      if (zMemberCalendarSyncEnabled_()) zSyncScheduleToMemberCalendars_(schedule);
    });
    zLog_('동기화', '', '캘린더', '성공', zMemberCalendarSyncEnabled_() ? '시트·개인 캘린더·팀 캘린더 반영 완료' : '시트에서 팀 캘린더 반영 완료');
  } catch (error) {
    zLog_('오류', '', '캘린더', '실패', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function zApiDate_(event) {
  if (event.start && event.start.date) return zDateObject_(event.start.date);
  if (event.start && event.start.dateTime) return new Date(event.start.dateTime);
  return null;
}

function zApiTime_(event) {
  if (!event.start || !event.start.dateTime) return '';
  return Utilities.formatDate(new Date(event.start.dateTime), ZEPHYRUS.kst, 'HH:mm');
}

function zMarkerScheduleId_(description) {
  var match = String(description || '').match(/\[(?:ZEPHYRUS|일정ID):([^\]]+)\]/);
  return match ? match[1] : '';
}

function zApplyCalendarApiEvent_(event) {
  var scheduleId = zMarkerScheduleId_(event.description);
  var schedule = scheduleId ? zFindScheduleById_(scheduleId) : null;
  if (!schedule) schedule = zFindScheduleByCalendarEvent_(String(event.id || ''));
  if (!schedule && event.iCalUID) schedule = zFindScheduleByCalendarEvent_(String(event.iCalUID));
  if (event.status === 'cancelled') {
    if (schedule) {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.status, ZEPHYRUS.status.deleted);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    }
    return;
  }
  var date = zApiDate_(event);
  if (!date || !event.summary) return;
  var ymd = Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd');
  var time = zApiTime_(event);
  var eventKey = String(event.id || event.iCalUID || '');
  if (schedule) {
    zSetSchedule_(schedule.row, ZEPHYRUS.col.date, date);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.time, time);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.title, String(event.summary));
    zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, eventKey));
    zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    return;
  }
  // A Calendar event has no "대상자" field.  When the creator is registered
  // in 구성원, treat that person as the registrant so they and administrators
  // receive the same Telegram record as a Telegram-created schedule.
  var creator = zMemberByEmail_(event.creator && event.creator.email || event.organizer && event.organizer.email);
  var imported = zAppendSchedule_({
    date: ymd, time: time, title: String(event.summary), memo: String(event.description || '').replace(/\[ZEPHYRUS:[^\]]+\]/g, '').trim(),
    targets: creator ? creator.name : '', registrant: creator ? creator.name : '캘린더',
    channel: '캘린더', calendarEvent: '팀:' + eventKey, deferSort: true
  });
  if (zSettings_()['즉시알림'] !== 'N') {
    var notified = zNotifySchedule_(imported, '[캘린더 새 일정]');
    zLog_('알림', imported.id, '텔레그램', notified.failed.length ? '일부 실패' : '성공',
      notified.sent.length ? notified.sent.join(', ') + '에게 발송' : '연결된 수신자가 없습니다.');
  }
}

function onCalendarChange_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var calendarId = zTeamCalendarId_();
    var props = zProps_();
    var token = props.getProperty('CALENDAR_SYNC_TOKEN');
    var nextToken = null;
    // A sync token can expire.  Retry once with a bounded full scan in the
    // same lock; do not call this function recursively while holding it.
    for (var attempt = 0; attempt < 2; attempt++) {
      var pageToken = null;
      nextToken = null;
      try {
        do {
          var options = { maxResults: 100, singleEvents: true, showDeleted: true };
          if (pageToken) options.pageToken = pageToken;
          if (token) options.syncToken = token;
          if (!token) {
            var days = Number(zSettings_()['캘린더가져오기일수'] || 60) || 60;
            options.timeMin = new Date(Date.now() - days * 86400000).toISOString();
          }
          var response = Calendar.Events.list(calendarId, options);
          (response.items || []).forEach(zApplyCalendarApiEvent_);
          pageToken = response.nextPageToken;
          if (response.nextSyncToken) nextToken = response.nextSyncToken;
        } while (pageToken);
        break;
      } catch (error) {
        if (token && String(error).indexOf('410') !== -1 && attempt === 0) {
          props.deleteProperty('CALENDAR_SYNC_TOKEN');
          token = '';
          continue;
        }
        throw error;
      }
    }
    if (nextToken) props.setProperty('CALENDAR_SYNC_TOKEN', nextToken);
    zNormalizeAndSortScheduleSheet_();
    zLog_('동기화', '', '캘린더', '성공', '팀 캘린더에서 시트 반영 완료');
  } finally {
    lock.releaseLock();
  }
}

function zDeleteMyTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
  });
}

function zInstallEditTrigger() {
  zDeleteMyTriggers_('handleEdit');
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(zSpreadsheet_()).onEdit().create();
  zNotice_('시트 수정 감지 트리거를 만들었습니다.');
}

function zInstallReminderTrigger() {
  zDeleteMyTriggers_('runReminders');
  ScriptApp.newTrigger('runReminders').timeBased().everyMinutes(1).create();
  zNotice_('리마인드 트리거를 1분 주기로 만들었습니다.');
}

function zInstallCalendarSyncTrigger() {
  zTeamCalendar_();
  zDeleteMyTriggers_('syncCalendar_');
  ScriptApp.newTrigger('syncCalendar_').timeBased().everyMinutes(1).create();
  zNotice_('시트에서 캘린더 동기화 트리거를 1분 주기로 만들었습니다.');
}

function zInstallCalendarReverseTrigger() {
  zTeamCalendar_();
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Apps Script 왼쪽 “서비스”에서 Calendar API를 먼저 추가하세요.');
  }
  zDeleteMyTriggers_('onCalendarChange_');
  ScriptApp.newTrigger('onCalendarChange_').forUserCalendar(zTeamCalendarId_()).onEventUpdated().create();
  zProps_().deleteProperty('CALENDAR_SYNC_TOKEN');
  onCalendarChange_();
  zNotice_('캘린더 변경 감지 트리거를 만들고 기존 최근 일정도 확인했습니다.');
}

function zInstallCoreTriggers() {
  zTeamCalendar_();
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Apps Script 왼쪽 “서비스”에서 Calendar API를 먼저 추가하세요.');
  }
  zDeleteMyTriggers_('handleEdit');
  zDeleteMyTriggers_('runReminders');
  zDeleteMyTriggers_('syncCalendar_');
  zDeleteMyTriggers_('onCalendarChange_');
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(zSpreadsheet_()).onEdit().create();
  ScriptApp.newTrigger('runReminders').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('syncCalendar_').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('onCalendarChange_').forUserCalendar(zTeamCalendarId_()).onEventUpdated().create();
  zProps_().deleteProperty('CALENDAR_SYNC_TOKEN');
  onCalendarChange_();
  zNotice_('자동 실행 4개를 현재 계정으로 켰습니다.');
}

function handleEdit(event) {
  try {
    if (!event || !event.range) return;
    var sheet = event.range.getSheet();
    if (sheet.getName() !== ZEPHYRUS.sheet.schedule || event.range.getRow() < 2) return;
    var map = zHeaders_(sheet);
    for (var row = event.range.getRow(); row < event.range.getRow() + event.range.getNumRows(); row++) {
      var values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      var schedule = zScheduleFromRow_(values, row, map);
      if (!schedule.date || !schedule.title || schedule.status === ZEPHYRUS.status.deleted) continue;
      if (!schedule.id) {
        zSetSchedule_(row, ZEPHYRUS.col.id, zId_());
        zSetSchedule_(row, ZEPHYRUS.col.registrant, schedule.registrant || '시트');
        zSetSchedule_(row, ZEPHYRUS.col.channel, '시트');
        zSetSchedule_(row, ZEPHYRUS.col.status, ZEPHYRUS.status.normal);
        zSetSchedule_(row, ZEPHYRUS.col.alarm, schedule.alarm || ZEPHYRUS.alarm.pending);
        zSetSchedule_(row, ZEPHYRUS.col.updated, zNowText_());
        schedule = zScheduleFromRow_(sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0], row, map);
        if (zSettings_()['즉시알림'] !== 'N') zNotifySchedule_(schedule, '[새 일정]');
      } else {
        zSetSchedule_(row, ZEPHYRUS.col.updated, zNowText_());
      }
    }
    zNormalizeAndSortScheduleSheet_();
  } catch (error) {
    zLog_('오류', '', '시트', '실패', 'handleEdit: ' + error);
  }
}

function runReminders() {
  var settings = zSettings_();
  var today = zToday_();
  var now = new Date();
  var nowTime = Utilities.formatDate(now, ZEPHYRUS.kst, 'HH:mm');
  var morning = String(settings['당일리마인드시각'] || '').trim();
  var beforeMinutes = Number(settings['사전리마인드분'] || 0) || 0;
  zSchedules_().forEach(function(schedule) {
    if (schedule.status === ZEPHYRUS.status.deleted || schedule.date !== today || !schedule.id) return;
    var flags = schedule.reminders ? schedule.reminders.split(';').filter(Boolean) : [];
    var changed = false;
    if (morning && nowTime >= morning && flags.indexOf('당일') === -1) {
      zNotifySchedule_(schedule, '[오늘 일정]');
      flags.push('당일');
      changed = true;
    }
    if (beforeMinutes > 0 && schedule.time && flags.indexOf('사전') === -1) {
      var parts = schedule.time.split(':').map(Number);
      var start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parts[0], parts[1]);
      var remaining = (start.getTime() - now.getTime()) / 60000;
      if (remaining <= beforeMinutes && remaining > -2) {
        zNotifySchedule_(schedule, '[' + beforeMinutes + '분 전 알림]');
        flags.push('사전');
        changed = true;
      }
    }
    if (changed) zSetSchedule_(schedule.row, ZEPHYRUS.col.reminders, flags.join(';'));
  });
}

function zShowStatus() {
  var lines = [];
  try { lines.push('스프레드시트: ' + zSpreadsheet_().getName()); } catch (error) { lines.push('스프레드시트: 오류 - ' + error.message); }
  try { lines.push('팀 캘린더: ' + zTeamCalendar_().getName()); } catch (error2) { lines.push('팀 캘린더: 오류 - ' + error2.message); }
  zMembers_().forEach(function(member) {
    if (!member.calendarId) {
      lines.push('직원 캘린더 - ' + member.name + ': 아직 없음');
      return;
    }
    var calendar = zMemberCalendar_(member);
    lines.push('직원 캘린더 - ' + member.name + ': ' + (calendar ? calendar.getName() : '접근 불가'));
  });
  lines.push('텔레그램 토큰: ' + (zProp_('BOT_TOKEN') ? '입력됨' : '아직 없음'));
  lines.push('현재 계정이 만든 트리거:');
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    lines.push('- ' + trigger.getHandlerFunction() + ' / ' + trigger.getEventType());
  });
  zNotice_(lines.join('\n'));
}
