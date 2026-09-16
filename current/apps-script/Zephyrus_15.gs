function zActiveScheduleEditCount_() {
  var props = zProps_();
  var count = Number(props.getProperty('ACTIVE_SCHEDULE_EDITS') || 0) || 0;
  var lastAt = Number(props.getProperty('LAST_SCHEDULE_EDIT_AT') || 0) || 0;
  // 비정상 종료로 숫자가 남아도 1분이 지나면 자동 복구한다.
  if (count > 0 && lastAt && Date.now() - lastAt > 60000) {
    props.setProperty('ACTIVE_SCHEDULE_EDITS', '0');
    return 0;
  }
  return count;
}

function zBeginScheduleEdit_() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    var props = zProps_();
    var count = zActiveScheduleEditCount_();
    var token = Utilities.getUuid();
    props.setProperties({
      ACTIVE_SCHEDULE_EDITS: String(count + 1),
      LAST_SCHEDULE_EDIT_TOKEN: token,
      LAST_SCHEDULE_EDIT_AT: String(Date.now())
    });
    return token;
  } finally {
    lock.releaseLock();
  }
}

function zEndScheduleEdit_(token) {
  var lock = LockService.getDocumentLock();
  var latestToken = '';
  var remaining = 0;
  lock.waitLock(5000);
  try {
    var props = zProps_();
    remaining = Math.max(0, zActiveScheduleEditCount_() - 1);
    props.setProperty('ACTIVE_SCHEDULE_EDITS', String(remaining));
    props.setProperty('LAST_SCHEDULE_EDIT_AT', String(Date.now()));
    latestToken = String(props.getProperty('LAST_SCHEDULE_EDIT_TOKEN') || token || '');
  } finally {
    lock.releaseLock();
  }

  // 다른 편집 실행이 아직 작업 중이면 마지막 실행이 끝날 때 정렬한다.
  if (remaining > 0) return false;

  // 사람이 연속으로 3~4칸을 빠르게 고칠 때는 잠깐 기다렸다가
  // 추가 편집이 없을 때만 한 번 정렬한다. 이 동안 행 번호가 움직이지 않는다.
  Utilities.sleep(1200);

  lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    var props2 = zProps_();
    if (zActiveScheduleEditCount_() > 0) return false;
    if (String(props2.getProperty('LAST_SCHEDULE_EDIT_TOKEN') || '') !== latestToken) return false;
    SpreadsheetApp.flush();
    zNormalizeAndSortScheduleSheet_(true);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function handleEdit(event) {
  var editToken = '';
  try {
    if (!event || !event.range) return;
    var sheet = event.range.getSheet();
    if (sheet.getName() === ZEPHYRUS.sheet.notice) {
      if (event.range.getRow() >= 2) zHandleNoticeEdit_(event);
      return;
    }
    if (sheet.getName() !== ZEPHYRUS.sheet.schedule || event.range.getRow() < 2) return;
    editToken = zBeginScheduleEdit_();
    var map = zHeaders_(sheet);
    var members = zMembers_();
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
      // 여기서는 행을 움직이지 않는다. 연속 수정이 모두 끝난 뒤
      // zEndScheduleEdit_()가 한 번만 정렬하므로 다른 편집 트리거의 행 번호가 안전하다.
      SpreadsheetApp.flush();

      // 시트 입력/수정도 텔레그램과 같은 가벼운 즉시반영 경로를 사용한다.
      // 정렬이 끝난 뒤 고유 일정ID로 새 행 위치를 다시 찾았으므로,
      // 캘린더 작업이 늦어도 사용자가 보는 시트는 먼저 정리된다.
      var immediateResult = zImmediateCalendarSyncLocked_(schedule);
      if (immediateResult && immediateResult.schedule) {
        schedule = immediateResult.schedule;
      }
    }
  } catch (error) {
    zLog_('오류', '', '시트', '실패', 'handleEdit: ' + error);
  } finally {
    if (editToken) {
      try { zEndScheduleEdit_(editToken); }
      catch (finishError) { zLog_('오류', '', '시트', '건너뜀', '연속편집 정렬: ' + finishError); }
    }
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
