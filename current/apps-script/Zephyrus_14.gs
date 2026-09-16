function zApplyCalendarApiEvent_(event) {
  // Public notices are delivery-only records.  Do not import them as regular
  // schedules, otherwise one notice would reappear in the 일정 sheet.
  if (zMarkerNoticeId_(event.description)) return;
  var scheduleId = zMarkerScheduleId_(event.description);
  var schedule = scheduleId ? zFindScheduleById_(scheduleId) : null;
  if (!schedule && !scheduleId) schedule = zFindScheduleByCalendarEvent_(String(event.id || ''));
  if (!schedule && !scheduleId && event.iCalUID) schedule = zFindScheduleByCalendarEvent_(String(event.iCalUID));
  if (event.status === 'cancelled') {
    if (schedule) {
      zHandleCancelledCalendarSchedule_(
        zTeamCalendarId_(),
        schedule,
        '팀',
        event,
        true,
        '공용캘린더'
      );
    }
    return;
  }
  // This team-calendar event was written from a sheet row that has since
  // been deleted.  Never recreate the row from its old calendar copy.
  if (scheduleId && !schedule) {
    zDeleteOrphanedMarkedEvent_(zTeamCalendarId_(), event, scheduleId, '공용캘린더');
    return;
  }
  var date = zApiDate_(event);
  if (!date || !event.summary) return;
  var ymd = Utilities.formatDate(date, ZEPHYRUS.kst, 'yyyy-MM-dd');
  var time = zApiTime_(event);
  var endTime = zApiEndTime_(event);
  var eventKey = String(event.id || event.iCalUID || '');
  var targetInfo = zCalendarTargetInfo_(event.description);
  if (schedule) {
    zSetSchedule_(schedule.row, ZEPHYRUS.col.date, date);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.time, time);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.endTime, endTime);
    zSetSchedule_(schedule.row, ZEPHYRUS.col.title, String(event.summary));
    zSetScheduleLocation_(schedule.row, String(event.location || ''));
    if (targetInfo.found) {
      if (zHasValidTargets_(targetInfo.targets)) {
        zSetSchedule_(schedule.row, ZEPHYRUS.col.targets, targetInfo.targets);
        zWriteScheduleMemberChecks_(schedule.row, targetInfo.targets);
      } else {
        zLog_('오류', schedule.id, '캘린더', '대상자 미반영', '구성원에 없는 대상자: ' + targetInfo.targets);
      }
    }
    zSetSchedule_(schedule.row, ZEPHYRUS.col.calendarEvent, zCalendarCellWithTeamId_(schedule, eventKey));
    zSetSchedule_(schedule.row, ZEPHYRUS.col.updated, zNowText_());
    schedule = zFindScheduleById_(schedule.id) || schedule;
    zSyncScheduleFiles_(schedule, event);
    return;
  }
  // A Calendar event has no "대상자" field.  When the creator is registered
  // in 구성원, treat that person as the registrant so they and administrators
  // receive the same Telegram record as a Telegram-created schedule.
  var creator = zMemberByEmail_(event.creator && event.creator.email || event.organizer && event.organizer.email);
  var targets = creator ? creator.name : '';
  if (targetInfo.found) {
    if (zHasValidTargets_(targetInfo.targets)) targets = targetInfo.targets;
    else zLog_('오류', '', '캘린더', '대상자 미반영', '구성원에 없는 대상자: ' + targetInfo.targets);
  }
  var imported = zAppendSchedule_({
    date: ymd, time: time, endTime: endTime, title: String(event.summary), location: String(event.location || ''),
    targets: targets, registrant: creator ? creator.name : '캘린더',
    channel: '캘린더', calendarEvent: '팀:' + eventKey, deferSort: true
  });
  zSyncScheduleFiles_(imported, event);
  if (zSettings_()['즉시알림'] !== 'N') {
    var notified = zNotifySchedule_(imported, '[캘린더 새 일정]');
    zSetSchedule_(imported.row, ZEPHYRUS.col.alarm, zAlarmResult_(notified, false));
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
    var hasCalendarChanges = false;
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
          var items = response.items || [];
          if (items.length) hasCalendarChanges = true;
          items.forEach(zApplyCalendarApiEvent_);
          pageToken = response.nextPageToken;
          if (response.nextSyncToken) nextToken = response.nextSyncToken;
        } while (pageToken);
        break;
      } catch (error) {
        if (token && zIsExpiredSyncToken_(error) && attempt === 0) {
          props.deleteProperty('CALENDAR_SYNC_TOKEN');
          token = '';
          continue;
        }
        throw error;
      }
    }
    if (nextToken) props.setProperty('CALENDAR_SYNC_TOKEN', nextToken);
    if (hasCalendarChanges) zNormalizeAndSortScheduleSheet_();
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

