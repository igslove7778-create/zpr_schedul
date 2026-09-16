function zDeleteMarkedScheduleEvent_(calendarId, scheduleId, knownEventId, channel) {
  calendarId = String(calendarId || '').trim();
  scheduleId = String(scheduleId || '').trim();
  knownEventId = String(knownEventId || '').trim();
  if (!calendarId || !scheduleId) return false;

  // 1) Advanced Calendar API ID path.
  if (typeof Calendar !== 'undefined' && Calendar.Events && knownEventId) {
    try {
      var apiEvent = Calendar.Events.get(calendarId, knownEventId);
      if (apiEvent && zMarkerScheduleId_(apiEvent.description) === scheduleId) {
        Calendar.Events.remove(calendarId, knownEventId, { sendUpdates: 'none' });
        return true;
      }
    } catch (ignoreApiId) {}
  }

  // 2) CalendarApp/iCalUID path.
  try {
    var appCalendar = CalendarApp.getCalendarById(calendarId);
    if (appCalendar && knownEventId) {
      var appEvent = appCalendar.getEventById(knownEventId);
      if (appEvent && zMarkerScheduleId_(appEvent.getDescription()) === scheduleId) {
        appEvent.deleteEvent();
        return true;
      }
    }
  } catch (ignoreCalendarAppId) {}

  // 3) Final fallback: locate the event by the exact Zephyrus marker.
  if (typeof Calendar !== 'undefined' && Calendar.Events) {
    try {
      var pageToken = null;
      do {
        var options = {
          q: scheduleId,
          singleEvents: true,
          showDeleted: false,
          maxResults: 250
        };
        if (pageToken) options.pageToken = pageToken;
        var response = Calendar.Events.list(calendarId, options);
        var items = response.items || [];
        for (var i = 0; i < items.length; i++) {
          if (zMarkerScheduleId_(items[i].description) !== scheduleId) continue;
          Calendar.Events.remove(calendarId, items[i].id, { sendUpdates: 'none' });
          return true;
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
    } catch (fallbackError) {
      zLog_('오류', scheduleId, channel || '캘린더', '실패', '표식 기준 일정 삭제 실패: ' + fallbackError);
      return false;
    }
  }
  return false;
}


function zImmediateCalendarSyncForNewSchedule_(schedule) {
  var current = zFindScheduleById_(schedule.id) || schedule;
  if (!current || current.status === ZEPHYRUS.status.deleted || !current.date || !current.title) {
    return { ok: false, detail: '유효한 일정이 아닙니다.' };
  }

  try {
    // 1) 공용(팀) 캘린더 즉시 반영
    var teamCalendar = zTeamCalendar_();
    var teamPriorId = zTeamEventId_(current);
    var teamEventId = zUpsertCalendarEvent_(teamCalendar, current, teamPriorId);

    if (teamEventId && teamEventId !== teamPriorId) {
      zSetSchedule_(
        current.row,
        ZEPHYRUS.col.calendarEvent,
        zCalendarCellWithTeamId_(current, teamEventId)
      );
      current = zFindScheduleById_(current.id) || current;
    }

    // 2) 개인캘린더 즉시 반영.
    // 대상자가 바뀐 경우에는 새 대상자에게 추가하는 것뿐 아니라,
    // 기존 대상자 캘린더의 복사본도 즉시 제거해야 한다.
    if (zMemberCalendarSyncEnabled_()) {
      var recipients = zScheduleCalendarRecipients_(current);
      var intended = {};
      recipients.forEach(function(member) {
        intended[zMemberCalendarKey_(member)] = member;
      });

      var allMembers = zMembers_().filter(function(member) { return member.calendarId; });

      allMembers.forEach(function(member) {
        var key = zMemberCalendarKey_(member);
        var priorId = zCalendarEventId_(current, key);

        // 더 이상 대상자가 아니면 기존 복사본만 제거한다.
        if (!intended[key]) {
          if (priorId) {
            zDeleteMarkedScheduleEvent_(member.calendarId, current.id, priorId, '개인캘린더');
            zSetSchedule_(
              current.row,
              ZEPHYRUS.col.calendarEvent,
              zCalendarCellWithId_(current, key, '')
            );
            current = zFindScheduleById_(current.id) || current;
          }
          return;
        }

        // 현재 대상자는 기존 일정 갱신 또는 새로 생성.
        var calendar = zMemberCalendar_(member);
        if (!calendar) {
          zLog_('오류', current.id, '개인캘린더', '건너뜀', member.name + ' 전용 캘린더 접근 실패');
          return;
        }

        var eventId = zUpsertCalendarEvent_(calendar, current, priorId);

        if (eventId && eventId !== priorId) {
          zSetSchedule_(
            current.row,
            ZEPHYRUS.col.calendarEvent,
            zCalendarCellWithId_(current, key, eventId)
          );
          current = zFindScheduleById_(current.id) || current;
        }
      });
    }

    current = zFindScheduleById_(current.id) || current;
    zSaveScheduleDeleteBackup_(current);
    return { ok: true, detail: '즉시 반영 완료', schedule: current };
  } catch (error) {
    // 실패해도 일정 등록 자체는 살려두고 기존 주기 동기화가 복구한다.
    zLog_('오류', current.id, '즉시캘린더', '실패', error);
    return { ok: false, detail: String(error), schedule: current };
  }
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
        zDeleteMarkedScheduleEvent_(member.calendarId, current.id, priorId, '개인캘린더');
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
    zDeleteMarkedScheduleEvent_(member.calendarId, schedule.id, eventId, '개인캘린더');
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

