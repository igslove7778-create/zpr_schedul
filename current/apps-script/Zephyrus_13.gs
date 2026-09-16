function zPurgeOrphanedMemberCalendarCopies_() {
  if (typeof Calendar === 'undefined' || !Calendar.Events) return 0;

  var schedules = {};
  zSchedules_().forEach(function(schedule) {
    if (schedule.id) schedules[schedule.id] = schedule;
  });

  var removed = 0;
  zMembers_().forEach(function(member) {
    if (!member.calendarId) return;
    var pageToken = null;
    try {
      do {
        var options = {
          q: 'ZEPHYRUS',
          singleEvents: true,
          showDeleted: false,
          maxResults: 250
        };
        if (pageToken) options.pageToken = pageToken;
        var response = Calendar.Events.list(member.calendarId, options);
        (response.items || []).forEach(function(event) {
          var scheduleId = zMarkerScheduleId_(event.description);
          if (!scheduleId) return;
          var schedule = schedules[scheduleId] || null;
          if (schedule && schedule.status !== ZEPHYRUS.status.deleted) return;
          try {
            Calendar.Events.remove(member.calendarId, event.id, { sendUpdates: 'none' });
            removed++;
          } catch (removeError) {
            zLog_('오류', scheduleId, '개인캘린더', '실패', '고아 복사본 삭제 ' + member.name + ': ' + removeError);
          }
        });
        pageToken = response.nextPageToken;
      } while (pageToken);
    } catch (error) {
      zLog_('오류', '', '개인캘린더', '실패', '고아 복사본 확인 ' + member.name + ': ' + error);
    }
  });
  return removed;
}

// One-time repair for rows that were removed before the complete private
// index existed.  It deletes only personal-calendar events with a Zephyrus
// schedule marker whose schedule row no longer exists (or is marked 삭제).
function zRemoveDeletedScheduleCopies() {
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Apps Script 왼쪽 “서비스”의 Calendar API가 필요합니다.');
  }
  var schedules = {};
  zSchedules_().forEach(function(schedule) { if (schedule.id) schedules[schedule.id] = schedule; });
  var from = new Date(Date.now() - 365 * 86400000).toISOString();
  var until = new Date(Date.now() + 5 * 365 * 86400000).toISOString();
  var deleted = 0;
  var failed = [];
  zMembers_().forEach(function(member) {
    if (!member.calendarId) return;
    var pageToken = null;
    try {
      do {
        var options = { timeMin: from, timeMax: until, singleEvents: true, showDeleted: false, maxResults: 250 };
        if (pageToken) options.pageToken = pageToken;
        var response = Calendar.Events.list(member.calendarId, options);
        (response.items || []).forEach(function(event) {
          var scheduleId = zMarkerScheduleId_(event.description);
          var schedule = scheduleId ? schedules[scheduleId] : null;
          if (!scheduleId || (schedule && schedule.status !== ZEPHYRUS.status.deleted)) return;
          try {
            Calendar.Events.remove(member.calendarId, event.id, { sendUpdates: 'none' });
            deleted++;
          } catch (removeError) {
            failed.push(member.name + ':' + scheduleId);
            zLog_('오류', scheduleId, '개인캘린더', '실패', '삭제 복구 ' + member.name + ': ' + removeError);
          }
        });
        pageToken = response.nextPageToken;
      } while (pageToken);
    } catch (error) {
      failed.push(member.name);
      zLog_('오류', '', '개인캘린더', '실패', '삭제 복구 읽기 ' + member.name + ': ' + error);
    }
  });
  if (failed.length) throw new Error('일부 개인캘린더 정리에 실패했습니다: ' + failed.join(', '));
  zNotice_(deleted + '건의 삭제된 일정 개인캘린더 복사본을 정리했습니다.');
}

function syncCalendar_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  var removedRows = 0;
  try {
    var calendar = zTeamCalendar_();
    // Deleted notices are detected from the hidden notice snapshot.
    zSyncDeletedNotices_();
    // A deleted sheet row must win over any calendar copy.  Do this before
    // reading team/personal calendars, otherwise a still-existing event can
    // be imported back into the sheet and appear to resurrect itself.
    removedRows += zSyncDeletedRows_(calendar);
    // The hidden index can be incomplete for older rows.  Sweep marked personal
    // copies before reverse sync so a deleted sheet row can never resurrect.
    if (zMemberCalendarSyncEnabled_()) {
      removedRows += zPurgeOrphanedMemberCalendarCopies_();
      zPullMemberCalendars_();
    }
    zSchedules_().forEach(function(schedule) {
      if (!schedule.id) return;
      if (schedule.status === ZEPHYRUS.status.deleted) {
        if (zTeamEventId_(schedule)) {
          zDeleteMarkedScheduleEvent_(zTeamCalendarId_(), schedule.id, zTeamEventId_(schedule), '공용캘린더');
          zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, ''));
        }
        if (zMemberCalendarSyncEnabled_()) zDeleteScheduleFromMemberCalendars_(schedule);
        return;
      }
      if (!schedule.date || !schedule.title) return;
      zSyncScheduleFiles_(schedule, null);
      var eventId = zUpsertCalendarEvent_(calendar, schedule);
      if (!eventId) return;
      if (eventId !== zTeamEventId_(schedule)) {
        zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, eventId));
        zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
        schedule = zFindScheduleById_(schedule.id) || schedule;
      }
      if (zMemberCalendarSyncEnabled_()) zSyncScheduleToMemberCalendars_(schedule);
    });
    // Refresh the private index after all new/upserted event IDs are known.
    removedRows += zSyncDeletedRows_(calendar);
    var detail = zMemberCalendarSyncEnabled_() ? '시트·개인 캘린더·팀 캘린더 반영 완료' : '시트에서 팀 캘린더 반영 완료';
    if (removedRows) detail += ' / 행 삭제 일정 ' + removedRows + '건 삭제';
    zLog_('동기화', '', '캘린더', '성공', detail);
  } catch (error) {
    zLog_('오류', '', '캘린더', '실패', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
  // Read external team-calendar edits only after deletion has completed.
  // This keeps a removed row from being recreated by its old calendar event.
  try {
    onCalendarChange_();
  } catch (reverseError) {
    zLog_('오류', '', '캘린더', '건너뜀', '팀 캘린더 확인 실패: ' + reverseError);
  }
}

// Visible manual-run entry point.  The real worker keeps its trailing
// underscore because it is also used by the one-minute trigger, but Apps
// Script can hide such names from the function picker.
function zRunCalendarSyncNow() {
  syncCalendar_();
  zNotice_('시트·공용캘린더·개인캘린더 동기화를 지금 실행했습니다.');
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

function zApiEndTime_(event) {
  if (!event.end || !event.end.dateTime) return '';
  return Utilities.formatDate(new Date(event.end.dateTime), ZEPHYRUS.kst, 'HH:mm');
}

function zMarkerScheduleId_(description) {
  var match = String(description || '').match(/\[(?:ZEPHYRUS|일정ID):([^\]]+)\]/);
  return match ? match[1] : '';
}

