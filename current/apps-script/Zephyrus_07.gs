function zProvisionMemberCalendars() {
  var created = [];
  var recovered = [];
  var missingEmail = [];
  var inaccessible = [];
  zMembers_().forEach(function(member) {
    if (member.calendarId) {
      var registered = zMemberCalendar_(member);
      if (registered) { zGrantMemberCalendarAccess_(registered, member.email); return; }
      inaccessible.push(member.name);
      return;
    }
    if (!member.email) {
      missingEmail.push(member.name);
      return;
    }
    var calendar = zOwnedMemberCalendar_(member);
    if (calendar) recovered.push(member.name);
    else {
      calendar = CalendarApp.createCalendar('제피로스 일정 - ' + member.name, {
        timeZone: ZEPHYRUS.kst,
        description: '제피로스 업무 일정 전용 캘린더. 직원 개인 일정과 분리해 사용합니다.'
      });
      created.push(member.name);
    }
    zGrantMemberCalendarAccess_(calendar, member.email);
    zSetMemberCalendarId_(member, calendar.getId());
  });
  if (missingEmail.length) throw new Error('구글 계정이 비어 있어 전용 캘린더를 만들 수 없는 구성원: ' + missingEmail.join(', '));
  if (inaccessible.length) throw new Error('기존 캘린더ID에는 접근할 수 없습니다: ' + inaccessible.join(', ') + '. 상태 확인에서 기존 캘린더를 확인한 뒤에만 해당 ID를 비우세요.');
  var message = [];
  if (created.length) message.push(created.join(', ') + ' 전용 캘린더를 만들었습니다.');
  if (recovered.length) message.push(recovered.join(', ') + ' 기존 전용 캘린더를 연결했습니다.');
  zNotice_(message.length ? message.join(' ') + ' 해당 직원에게 공유했습니다.' : '전용 캘린더와 공유 권한이 이미 준비되어 있습니다.');
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

function zIsExpiredSyncToken_(error) {
  // Calendar API normally returns HTTP 410, but Apps Script can surface only
  // the human-readable message.  Both mean exactly the same thing: discard
  // the old cursor once and perform a bounded full read.
  return /(\b410\b|sync token is no longer valid|full sync is required)/i.test(String(error || ''));
}

function zDeleteOrphanedMarkedEvent_(calendarId, event, scheduleId, channel) {
  if (!event || event.status === 'cancelled' || !event.id) return;
  try {
    Calendar.Events.remove(calendarId, event.id, { sendUpdates: 'none' });
  } catch (error) {
    zLog_('오류', scheduleId, channel, '실패', '삭제된 시트 일정의 캘린더 복사본 정리 실패: ' + error);
  }
}

function zApplyMemberCalendarApiEvent_(member, event) {
  // Company notices are one-way delivery items, not work schedules.  Ignore
  // their marker while polling personal calendars so they never create a
  // duplicate row in the 일정 sheet or flow back to the team calendar.
  if (zMarkerNoticeId_(event.description)) return;
  var key = zMemberCalendarKey_(member);
  var scheduleId = zMarkerScheduleId_(event.description);
  var schedule = scheduleId ? zFindScheduleById_(scheduleId) : null;
  if (!schedule && !scheduleId) schedule = zFindScheduleByCalendarEvent_(String(event.id || event.iCalUID || ''));

  if (event.status === 'cancelled') {
    if (schedule && schedule.registrant === member.name) {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.status, ZEPHYRUS.status.deleted);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    }
    return;
  }

  // A marked event was copied from an existing sheet schedule.  If that row
  // is now gone, it is a deletion—not a new private schedule.  Remove the
  // personal copy rather than importing it back as 김형원/김명재's own work.
  if (scheduleId && !schedule) {
    zDeleteOrphanedMarkedEvent_(member.calendarId, event, scheduleId, '개인캘린더');
    return;
  }

  var date = zApiDate_(event);
  if (!date || !event.summary) return;
  var ymd = Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd');
  var time = zApiTime_(event);
  var endTime = zApiEndTime_(event);
  var eventId = String(event.id || event.iCalUID || '');

  if (schedule) {
    // Only the owner of a schedule may change its shared source record from
    // their personal calendar.  A recipient's copied event is read-only to
    // the synchronizer, even if Google granted them edit access.
    if (schedule.registrant === member.name || schedule.channel === '개인캘린더') {
      zSetSchedule_(schedule.row, ZEPHYRUS.col.date, date);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.time, time);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.endTime, endTime);
      zSetSchedule_(schedule.row, ZEPHYRUS.col.title, String(event.summary));
      zSetScheduleLocation_(schedule.row, String(event.location || ''));
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
    endTime: endTime,
    title: String(event.summary),
    location: String(event.location || ''),
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
      if (token && zIsExpiredSyncToken_(error) && attempt === 0) {
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

// Delete a Zephyrus calendar event safely even when Calendar API event.id and
// CalendarApp event IDs differ.  Never delete an event unless its schedule
// marker matches the requested schedule ID.
