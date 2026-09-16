function zScheduleFromRow_(row, index, map, members) {
  function value(name) { return map[name] ? row[map[name] - 1] : ''; }
  return {
    row: index,
    date: zNormalizeDate_(value(ZEPHYRUS.col.date)),
    time: zNormalizeTime_(value(ZEPHYRUS.col.time)),
    endTime: zNormalizeTime_(value(ZEPHYRUS.col.endTime)),
    title: String(value(ZEPHYRUS.col.title) || '').trim(),
    location: String(value(ZEPHYRUS.col.location) || value('메모') || '').trim(),
    file: String(value(ZEPHYRUS.col.file) || '').trim(),
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

function zScheduleLocationColumn_(map) {
  return map[ZEPHYRUS.col.location] || map['메모'] || 0;
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

function zHasValidTargets_(targetText) {
  var names = zTargetNames_(targetText);
  if (names === '전부' || names === '전체') return true;
  if (!names.length) return false;
  var known = {};
  zMembers_().forEach(function(member) { known[member.name] = true; });
  return names.every(function(name) { return known[name]; });
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

// 사람별 대상자 표시는 체크박스 대신 "○" 또는 빈칸으로만 보여 준다.
// F열 대상자가 기준이고, 이 열은 보기와 간단한 수동 선택을 함께 지원한다.
function zPrepareMemberCheckColumns() {
  var members = zMembers_();
  if (!members.length) throw new Error('구성원 시트에 먼저 이름과 구글 계정을 입력하세요.');
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  var added = [];
  var rowCount = Math.max(1, sheet.getMaxRows() - 1);
  members.forEach(function(member) {
    var column = map[member.name];
    if (!column) {
      column = sheet.getLastColumn() + 1;
      sheet.getRange(1, column).setValue(member.name).setFontWeight('bold').setBackground('#E2F0D9');
      map[member.name] = column;
      added.push(member.name);
    }
    // Old checkbox validation renders TRUE/FALSE when its formatting is lost.
    // Remove it so the sheet always uses only a circle or an empty cell.
    sheet.getRange(2, column, rowCount, 1)
      .clearDataValidations()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(18)
      .setFontWeight('bold');
    sheet.setColumnWidth(column, 42);
  });
  if (sheet.getLastRow() >= 2) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    rows.forEach(function(values, index) {
      zWriteScheduleMemberChecks_(index + 2, String(values[map[ZEPHYRUS.col.targets] - 1] || '').trim());
    });
  }
  zNotice_(added.length ? '대상자 ○ 표시 열을 만들고 기존 표시도 정리했습니다: ' + added.join(', ') : '대상자 표시를 ○ 또는 빈칸으로 정리했습니다.');
}

function zWriteScheduleMemberChecks_(rowNumber, targetText) {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  var map = zHeaders_(sheet);
  var targets = zTargetNames_(targetText);
  zMembers_().forEach(function(member) {
    if (!map[member.name]) return;
    var checked = targets === '전부' || targets === '전체' || targets.indexOf(member.name) !== -1;
    sheet.getRange(rowNumber, map[member.name])
      .setValue(checked ? '○' : '')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(18)
      .setFontWeight('bold');
  });
}

function zRangeTouchesColumn_(range, column) {
  return column >= range.getColumn() && column < range.getColumn() + range.getNumColumns();
}

function zCheckedMemberNames_(values, map, members) {
  return members.filter(function(member) {
    return map[member.name] && zIsChecked_(values[map[member.name] - 1]);
  }).map(function(member) { return member.name; });
}

// F열과 사람별 체크칸은 같은 대상자를 보여준다. F열을 고치면 체크칸을
// 맞추고, 체크칸을 고치면 F열의 이름 목록을 맞춘다. 한 번의 편집에서 둘을
// 동시에 건드린 경우에는 F열 값을 우선한다.
function zSyncTargetInputsForEdit_(sheet, rowNumber, map, values, editedRange, members) {
  var targetColumn = map[ZEPHYRUS.col.targets];
  if (!targetColumn) return false;
  var targetTouched = zRangeTouchesColumn_(editedRange, targetColumn);
  var checksTouched = members.some(function(member) {
    return map[member.name] && zRangeTouchesColumn_(editedRange, map[member.name]);
  });
  if (!targetTouched && !checksTouched) return false;
  if (targetTouched) {
    zWriteScheduleMemberChecks_(rowNumber, String(values[targetColumn - 1] || '').trim());
    return true;
  }
  var targetText = zCheckedMemberNames_(values, map, members).join(', ');
  if (String(values[targetColumn - 1] || '').trim() !== targetText) {
    sheet.getRange(rowNumber, targetColumn).setValue(targetText);
  }
  return true;
}

// 대상자(F열)를 기준으로 사람별 체크칸을 한 번에 다시 맞춘다.
// 기존 일정에도 바로 적용할 수 있는 정리용 함수다.
function zSyncAllScheduleMemberChecks() {
  var sheet = zSheet_(ZEPHYRUS.sheet.schedule);
  if (sheet.getLastRow() < 2) return;
  var map = zHeaders_(sheet);
  if (!map[ZEPHYRUS.col.targets]) throw new Error('일정 시트에 대상자 열이 없습니다.');
  var members = zMembers_();
  if (!members.length) throw new Error('구성원 시트에 먼저 구성원을 등록하세요.');
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach(function(values, index) {
    var targetText = String(values[map[ZEPHYRUS.col.targets] - 1] || '').trim();
    zWriteScheduleMemberChecks_(index + 2, targetText);
  });
  zNotice_('대상자(F열)를 기준으로 사람별 체크 표시를 정리했습니다.');
}

// 설치형 편집 트리거에는 실제 편집자의 계정이 제공될 수 있다.
// 보안 정책상 계정을 알 수 없는 경우에만 예전처럼 "시트"로 남긴다.
function zRegistrantFromSheetEdit_(event, schedule) {
  var existing = String(schedule && schedule.registrant || '').trim();
  if (existing && existing !== '시트') return existing;
  var email = '';
  try {
    email = event && event.user && event.user.getEmail ? String(event.user.getEmail() || '').trim() : '';
  } catch (ignore) {}
  var member = zMemberByEmail_(email);
  return member ? member.name : '시트';
}

// 캘린더 ID 기준으로 실제 캘린더 이름과 현재 관리자 화면의 표시명을
// 함께 바로잡는다. 기존 일정, 공유 권한, 캘린더 ID는 건드리지 않는다.
function zRepairMemberCalendarDisplayNames() {
  if (typeof Calendar === 'undefined' || !Calendar.Calendars || !Calendar.CalendarList) {
    throw new Error('Apps Script 왼쪽 “서비스”에서 Calendar API를 먼저 추가하세요.');
  }
  var repaired = [];
  var failed = [];
  zMembers_().forEach(function(member) {
    if (!member.calendarId) return;
    var name = '제피로스 일정 - ' + member.name;
    try {
      Calendar.Calendars.patch({ summary: name }, member.calendarId);
      Calendar.CalendarList.patch({ summaryOverride: name }, member.calendarId);
      repaired.push(name);
    } catch (error) {
      failed.push(member.name);
      zLog_('오류', '', '개인캘린더', '실패', '표시명 정리 ' + member.name + ': ' + error);
    }
  });
  if (failed.length) throw new Error('표시명을 정리하지 못한 구성원: ' + failed.join(', '));
  zNotice_(repaired.length ? '개인 캘린더 표시명을 구성원 이름으로 정리했습니다.' : '정리할 개인 캘린더가 없습니다.');
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

