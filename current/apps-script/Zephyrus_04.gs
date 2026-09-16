function zSetSchedule_(rowNumber, name, value) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  if (!map[name]) throw new Error('일정 시트에 열이 없습니다: ' + name);
  sheet.getRange(rowNumber, map[name]).setValue(value);
}

function zSetScheduleLocation_(rowNumber, value) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var column = zScheduleLocationColumn_(zHeaders_(sheet));
  if (!column) throw new Error('일정 시트에 장소 열이 없습니다. 시트 구조 점검을 먼저 실행하세요.');
  sheet.getRange(rowNumber, column).setValue(value || '');
}

// Keep dates as real date values, but display them in the familiar Korean
// format.  Sorting the complete row range keeps every schedule's notes,
// targets, IDs and check marks together.
function zNormalizeAndSortScheduleSheet_(force) {
  // 연속 편집 중에는 다른 트리거가 행을 움직이지 못하게 한다.
  // 마지막 편집이 끝난 뒤 zEndScheduleEdit_()가 force=true로 딱 한 번 정렬한다.
  if (!force && zActiveScheduleEditCount_() > 0) return 0;
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
    if (
      zNormalizeDate_(before) !== Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd') ||
      !(before instanceof Date)
    ) {
      normalizedCount++;
    }
    return [date];
  });

  dateRange.setValues(values).setNumberFormat('m"월" d"일"');

  // Google Sheets의 일반 오름차순 정렬은 빈 행을 위쪽에 끌어올릴 수 있다.
  // 임시 보조열을 사용해 "실제 일정 = 0 / 빈 행 = 1"로 먼저 정렬한 뒤
  // 날짜와 시작시간 순으로 정렬한다. 그래서 일정 사이에 빈 줄이 끼지 않는다.
  var originalLastColumn = sheet.getLastColumn();
  var helperColumn = originalLastColumn + 1;

  if (sheet.getMaxColumns() < helperColumn) {
    sheet.insertColumnAfter(originalLastColumn);
  }

  var rowValues = sheet.getRange(2, 1, rowCount, originalLastColumn).getValues();
  var helperValues = rowValues.map(function(row) {
    var hasContent = row.some(function(value) {
      return value !== '' && value !== null;
    });
    return [hasContent ? 0 : 1];
  });

  sheet.getRange(2, helperColumn, rowCount, 1).setValues(helperValues);

  var sortBy = [
    { column: helperColumn, ascending: true },
    { column: map[ZEPHYRUS.col.date], ascending: true }
  ];
  if (map[ZEPHYRUS.col.time]) {
    sortBy.push({ column: map[ZEPHYRUS.col.time], ascending: true });
  }

  sheet.getRange(2, 1, rowCount, helperColumn).sort(sortBy);
  sheet.getRange(2, helperColumn, rowCount, 1).clearContent();

  return normalizedCount;
}

function zNormalizeAndSortSchedules() {
  var count = zNormalizeAndSortScheduleSheet_(true);
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
  var locationColumn = zScheduleLocationColumn_(map);
  var id = input.id || zId_();
  var date = zDateObject_(input.date);
  row[map[ZEPHYRUS.col.date] - 1] = date || input.date;
  row[map[ZEPHYRUS.col.time] - 1] = input.time || '';
  if (map[ZEPHYRUS.col.endTime]) row[map[ZEPHYRUS.col.endTime] - 1] = input.endTime || '';
  row[map[ZEPHYRUS.col.title] - 1] = input.title || '';
  if (locationColumn) row[locationColumn - 1] = input.location || '';
  row[map[ZEPHYRUS.col.targets] - 1] = input.targets || '';
  row[map[ZEPHYRUS.col.alarm] - 1] = input.alarm || ZEPHYRUS.alarm.pending;
  if (map[ZEPHYRUS.col.file]) row[map[ZEPHYRUS.col.file] - 1] = input.file || '';
  row[map[ZEPHYRUS.col.id] - 1] = id;
  row[map[ZEPHYRUS.col.registrant] - 1] = input.registrant || '시스템';
  row[map[ZEPHYRUS.col.channel] - 1] = input.channel || '시트';
  row[map[ZEPHYRUS.col.calendarEvent] - 1] = input.calendarEvent || '';
  row[map[ZEPHYRUS.col.updated] - 1] = zNowText_();
  row[map[ZEPHYRUS.col.status] - 1] = input.status || ZEPHYRUS.status.normal;
  row[map[ZEPHYRUS.col.reminders] - 1] = input.reminders || '';
  sheet.appendRow(row);
  SpreadsheetApp.flush();

  // 동시에 여러 사용자가 appendRow를 실행해도 getLastRow()를 믿지 않는다.
  // 고유 일정ID로 방금 추가한 실제 행을 다시 찾아 후속 작업을 한다.
  var inserted = zFindScheduleById_(id);
  var number = inserted ? inserted.row : sheet.getLastRow();

  sheet.getRange(number, map[ZEPHYRUS.col.date]).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(number, map[ZEPHYRUS.col.time]).setNumberFormat('@');
  if (map[ZEPHYRUS.col.endTime]) sheet.getRange(number, map[ZEPHYRUS.col.endTime]).setNumberFormat('@');
  zWriteScheduleMemberChecks_(number, input.targets || input.registrant || '');

  if (!input.deferSort) {
    zNormalizeAndSortScheduleSheet_();
    inserted = zFindScheduleById_(id);
  }
  return inserted || zFindScheduleById_(id);
}

function zLog_(kind, scheduleId, channel, result, detail) {
  if (!zShouldKeepLog_(kind, result, detail)) return;
  try {
    zSheet_(ZEPHYRUS.sheet.log).appendRow([zNowText_(), kind, scheduleId || '', channel || '', result || '', String(detail || '').slice(0, 500)]);
  } catch (error) {
    console.log('log failure: ' + error);
  }
}

function zNoticeTimeText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    var hour = Number(Utilities.formatDate(value, ZEPHYRUS.kst, 'H'));
    var minute = Utilities.formatDate(value, ZEPHYRUS.kst, 'mm');
    var period = hour < 12 ? '오전' : '오후';
    var h12 = hour % 12 || 12;
    return period + ' ' + h12 + ':' + minute;
  }
  var text = String(value || '').trim();
  // 과거 백업에 Google Sheets 시간값이 "Sat Dec 30 1899 14:00:00 ..." 형태로
  // 문자열 저장된 경우에도 사람이 읽는 한국식 시간으로 복구한다.
  var legacy = text.match(/(?:1899|1900).*?(\d{1,2}):(\d{2}):(\d{2})/i);
  if (legacy) {
    var h = Number(legacy[1]);
    var m = legacy[2];
    var p = h < 12 ? '오전' : '오후';
    var h12legacy = h % 12 || 12;
    return p + ' ' + h12legacy + ':' + m;
  }
  return text;
}

