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
    if (candidate.getTitle() === zCalendarTitleForSchedule_(schedule) && zSameEventTime_(candidate, times) && !same) same = candidate;
  });
  return tagged || same;
}

function zUpdateCalendarEvent_(event, schedule, times, description) {
  var changed = false;
  var calendarTitle = zCalendarTitleForSchedule_(schedule);
  if (event.getTitle() !== calendarTitle) {
    event.setTitle(calendarTitle);
    changed = true;
  }
  if (event.getLocation() !== (schedule.location || '')) {
    event.setLocation(schedule.location || '');
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


function zDisplayTitleForSchedule_(schedule) {
  var title = String(schedule && schedule.title || '');
  var targets = String(schedule && schedule.targets || '').trim();
  if ((targets === '전부' || targets === '전체') && title.indexOf('[전체일정] ') !== 0) {
    title = '[전체일정] ' + title;
  }
  return title;
}

function zCalendarTitleForSchedule_(schedule) {
  return zDisplayTitleForSchedule_(schedule);
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
    event = calendar.createAllDayEvent(zCalendarTitleForSchedule_(schedule), times.start, { description: description, location: schedule.location || '' });
  } else {
    event = calendar.createEvent(zCalendarTitleForSchedule_(schedule), times.start, times.end, { description: description, location: schedule.location || '' });
  }
  return event.getId();
}

// Keep the complete calendar-event map privately.  When a sheet row is
// removed, this lets us delete the exact team and personal copies without
// scanning unrelated events by title or date.
function zSyncIndexSheet_() {
  var ss = zSpreadsheet_();
  var sheet = ss.getSheetByName(ZEPHYRUS.sheet.syncIndex);
  if (!sheet) sheet = ss.insertSheet(ZEPHYRUS.sheet.syncIndex);
  if (sheet.getLastRow() === 0) sheet.appendRow(['일정ID', '캘린더ID목록', '일정백업JSON']);
  else sheet.getRange(1, 1, 1, 3).setValues([['일정ID', '캘린더ID목록', '일정백업JSON']]);
  sheet.hideSheet();
  return sheet;
}

function zScheduleSnapshot_(schedule) {
  var telegramRecipients = [];
  var seen = {};
  try {
    zRecipients_(schedule).forEach(function(member) {
      var chatId = String(member.telegramId || '').trim();
      if (!chatId || seen[chatId]) return;
      seen[chatId] = true;
      telegramRecipients.push({ name: member.name, chatId: chatId });
    });
  } catch (ignore) {}

  return {
    id: String(schedule.id || ''),
    date: String(schedule.date || ''),
    time: String(schedule.time || ''),
    endTime: String(schedule.endTime || ''),
    title: String(schedule.title || ''),
    location: String(schedule.location || ''),
    targets: String(schedule.targets || ''),
    registrant: String(schedule.registrant || ''),
    channel: String(schedule.channel || ''),
    telegramRecipients: telegramRecipients
  };
}

function zNotifyDeletedSchedule_(snapshot) {
  if (!snapshot || !snapshot.id) return;

  var cache = CacheService.getScriptCache();
  var dedupeKey = 'TG_CANCEL_' + String(snapshot.id);
  if (cache.get(dedupeKey)) return;
  cache.put(dedupeKey, '1', 120);

  var recipients = Array.isArray(snapshot.telegramRecipients) ? snapshot.telegramRecipients : [];
  // 예전 백업에는 telegramRecipients가 없을 수 있으므로 현재 구성원 정보로 한 번 복구한다.
  if (!recipients.length) {
    try {
      recipients = zRecipients_(snapshot).map(function(member) {
        return { name: member.name, chatId: String(member.telegramId || '').trim() };
      }).filter(function(item) { return item.chatId; });
    } catch (ignore) {}
  }

  var seen = {};
  recipients.forEach(function(item) {
    var chatId = String(item.chatId || '').trim();
    if (!chatId || seen[chatId]) return;
    seen[chatId] = true;
    try {
      zSendTelegram_(chatId, '[일정 취소]\n' + zScheduleText_(snapshot));
    } catch (error) {
      zLog_('오류', snapshot.id, '텔레그램', '실패', '일정 취소 알림 ' + String(item.name || '') + ': ' + error);
    }
  });
}

function zSaveScheduleDeleteBackup_(schedule) {
  if (!schedule || !schedule.id) return;
  var indexSheet = zSyncIndexSheet_();
  var rows = indexSheet.getLastRow() < 2 ? [] : indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 3).getValues();
  var targetRow = 0;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '') === String(schedule.id)) {
      targetRow = i + 2;
      break;
    }
  }
  if (!targetRow) targetRow = indexSheet.getLastRow() + 1;
  indexSheet.getRange(targetRow, 1, 1, 3).setValues([[
    String(schedule.id),
    String(schedule.calendarEvent || ''),
    JSON.stringify(zScheduleSnapshot_(schedule))
  ]]);
}

function zSyncDeletedRows_(calendar) {
  var indexSheet = zSyncIndexSheet_();
  var current = {};
  zSchedules_().forEach(function(schedule) {
    if (schedule.id) current[schedule.id] = { calendarEvent: schedule.calendarEvent, snapshot: zScheduleSnapshot_(schedule) };
  });
  var previous = indexSheet.getLastRow() < 2 ? [] : indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 3).getValues();
  var removed = 0;
  var membersByName = {};
  zMembers_().forEach(function(member) { membersByName[member.name] = member; });
  previous.forEach(function(row) {
    var id = String(row[0] || '');
    var ids = zParseCalendarIds_(row[1]);
    if (!id || current[id]) return;

    var snapshot = null;
    try { snapshot = row[2] ? JSON.parse(String(row[2])) : null; } catch (ignore) {}
    if (snapshot) zNotifyDeletedSchedule_(snapshot);

    Object.keys(ids).forEach(function(key) {
      var eventId = String(ids[key] || '');
      var isTeam = key === '팀' || key === 'TEAM' || key === 'team';
      var member = isTeam ? null : membersByName[key];
      var calendarId = isTeam ? zTeamCalendarId_() : (member && member.calendarId);
      if (!calendarId) return;
      if (zDeleteMarkedScheduleEvent_(calendarId, id, eventId, isTeam ? '공용캘린더' : '개인캘린더')) removed++;
    });
  });
  var rows = Object.keys(current).map(function(id) {
    return [id, current[id].calendarEvent || '', JSON.stringify(current[id].snapshot || {})];
  });
  var clearRows = Math.max(indexSheet.getLastRow() - 1, rows.length, 1);
  indexSheet.getRange(2, 1, clearRows, 3).clearContent();
  if (rows.length) indexSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return removed;
}

// Before reverse-syncing personal calendars, remove Zephyrus copies whose
// source schedule row no longer exists.  This does not touch ordinary personal
// events because only events carrying a Zephyrus schedule marker are examined.
