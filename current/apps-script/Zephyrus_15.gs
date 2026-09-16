function handleEdit(event) {
  try {
    if (!event || !event.range) return;
    var sheet = event.range.getSheet();
    if (sheet.getName() === ZEPHYRUS.sheet.notice) {
      if (event.range.getRow() >= 2) zHandleNoticeEdit_(event);
      return;
    }
    if (sheet.getName() !== ZEPHYRUS.sheet.schedule || event.range.getRow() < 2) return;
    var map = zHeaders_(sheet);
    var members = zMembers_();
    var shouldSort = false;
    var singleRowEdit = event.range.getNumRows() === 1;
    for (var row = event.range.getRow(); row < event.range.getRow() + event.range.getNumRows(); row++) {
      var values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      // 대상자(F열)와 사람별 체크칸을 같은 값으로 유지한다.
      if (zSyncTargetInputsForEdit_(sheet, row, map, values, event.range, members)) {
        values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
      var schedule = zScheduleFromRow_(values, row, map, members);
      var rawTime = values[map[ZEPHYRUS.col.time] - 1];
      // Show the normalized result in the sheet as well as using it for the
      // calendar.  For example, "오후 2시" becomes the unambiguous "14:00".
      if (schedule.time && String(rawTime).trim() !== schedule.time) {
        sheet.getRange(row, map[ZEPHYRUS.col.time]).setValue(schedule.time).setNumberFormat('@');
        values[map[ZEPHYRUS.col.time] - 1] = schedule.time;
        schedule = zScheduleFromRow_(values, row, map, members);
      }
      var rawEndTime = map[ZEPHYRUS.col.endTime] ? values[map[ZEPHYRUS.col.endTime] - 1] : '';
      if (schedule.endTime && String(rawEndTime).trim() !== schedule.endTime) {
        sheet.getRange(row, map[ZEPHYRUS.col.endTime]).setValue(schedule.endTime).setNumberFormat('@');
        values[map[ZEPHYRUS.col.endTime] - 1] = schedule.endTime;
        schedule = zScheduleFromRow_(values, row, map, members);
      }
      if (!schedule.date || !schedule.title || schedule.status === ZEPHYRUS.status.deleted) continue;
      // A new row stays a draft until the writer has set a valid target.
      // This keeps the row in place and prevents an early Telegram message
      // while the memo and recipients are still being entered.
      if (!schedule.id && !zHasValidTargets_(schedule.targets)) continue;
      if (!schedule.id) {
        var createLock = LockService.getDocumentLock();
        createLock.waitLock(10000);
        try {
          // 다른 중복 트리거가 먼저 일정ID를 만들었을 수 있으므로 잠금 후 반드시 다시 읽는다.
          var lockedValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
          schedule = zScheduleFromRow_(lockedValues, row, map, members);
          if (!schedule.id) {
            zSetSchedule_(row, ZEPHYRUS.col.id, zId_());
            zSetSchedule_(row, ZEPHYRUS.col.registrant, zRegistrantFromSheetEdit_(event, schedule));
            zSetSchedule_(row, ZEPHYRUS.col.channel, '시트');
            zSetSchedule_(row, ZEPHYRUS.col.status, ZEPHYRUS.status.normal);
            zSetSchedule_(row, ZEPHYRUS.col.alarm, schedule.alarm || ZEPHYRUS.alarm.pending);
            zSetSchedule_(row, ZEPHYRUS.col.updated, zNowText_());
            schedule = zScheduleFromRow_(sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0], row, map, members);
            // 행을 바로 지워도 취소 알림이 빠지지 않도록 생성 즉시 백업한다.
            zSaveScheduleDeleteBackup_(schedule);
            if (zSettings_()['즉시알림'] !== 'N') {
              var notified = zNotifySchedule_(schedule, '[새 일정]');
              zSetSchedule_(schedule.row, ZEPHYRUS.col.alarm, zAlarmResult_(notified, false));
            }
          }
        } finally {
          createLock.releaseLock();
        }
      } else {
        zSetSchedule_(row, ZEPHYRUS.col.updated, zNowText_());
        schedule = zFindScheduleById_(schedule.id) || schedule;
        zSaveScheduleDeleteBackup_(schedule);
      }
      // 완성된 한 행 편집은 캘린더 API보다 먼저 즉시 정렬한다.
      // Google Calendar 처리 지연 때문에 시트 정렬까지 늦어지는 것을 막는다.
      if (singleRowEdit) {
        SpreadsheetApp.flush();
        zNormalizeAndSortScheduleSheet_();
        schedule = zFindScheduleById_(schedule.id) || schedule;
      } else {
        // 여러 행 붙여넣기/일괄편집은 행 번호가 중간에 움직이지 않도록
        // 전체 처리 후 한 번만 정렬한다.
        shouldSort = true;
      }

      // 시트 입력/수정도 텔레그램과 같은 가벼운 즉시반영 경로를 사용한다.
      // 정렬이 끝난 뒤 고유 일정ID로 새 행 위치를 다시 찾았으므로,
      // 캘린더 작업이 늦어도 사용자가 보는 시트는 먼저 정리된다.
      var immediateResult = zImmediateCalendarSyncForNewSchedule_(schedule);
      if (immediateResult && immediateResult.schedule) {
        schedule = immediateResult.schedule;
      }
    }
    if (shouldSort) zNormalizeAndSortScheduleSheet_();
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

