function zNoticeIndexSheet_() {
  var ss = zSpreadsheet_();
  var sheet = ss.getSheetByName(ZEPHYRUS.sheet.noticeIndex);
  if (!sheet) sheet = ss.insertSheet(ZEPHYRUS.sheet.noticeIndex);

  // 예전 3열 백업이면 날짜 열을 하나 끼워 넣어 안전하게 확장한다.
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['공지ID', '공지날짜', '공지제목', '공지내용']);
  } else {
    var header2 = String(sheet.getRange(1, 2).getValue() || '').trim();
    if (header2 === '공지제목') sheet.insertColumnAfter(1);
    sheet.getRange(1, 1, 1, 4).setValues([['공지ID', '공지날짜', '공지제목', '공지내용']]);
  }
  sheet.hideSheet();
  return sheet;
}

function zSaveNoticeSnapshot_(notice) {
  if (!notice || !notice.id) return;
  var sheet = zNoticeIndexSheet_();
  var rows = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var targetRow = 0;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '') === String(notice.id)) {
      targetRow = i + 2;
      break;
    }
  }
  if (!targetRow) targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1, 1, 4).setValues([[
    String(notice.id || ''),
    zNoticeDateText_(notice.writtenAt) || '',
    String(notice.title || ''),
    String(notice.content || '')
  ]]);
}

function zDeleteNoticeEventsInCalendar_(calendarId, notice, channel) {
  if (!calendarId || !notice || typeof Calendar === 'undefined' || !Calendar.Events) return 0;
  var noticeId = String(notice.id || '');
  var wantedDate = zNoticeDateText_(notice.writtenAt) || '';
  var wantedTitle = String(notice.title || '');
  var wantedContent = String(notice.content || '');
  var deleted = 0;
  var pageToken = null;
  var from = new Date(Date.now() - 2 * 365 * 86400000).toISOString();
  var until = new Date(Date.now() + 10 * 365 * 86400000).toISOString();
  try {
    do {
      var options = {
        timeMin: from,
        timeMax: until,
        singleEvents: true,
        showDeleted: false,
        maxResults: 250
      };
      if (pageToken) options.pageToken = pageToken;
      var response = Calendar.Events.list(calendarId, options);
      (response.items || []).forEach(function(event) {
        var description = String(event.description || '');
        var summary = String(event.summary || '');
        var markerMatch = noticeId && zMarkerNoticeId_(description) === noticeId;
        // 과거 중복 실행에서 서로 다른 공지ID로 만들어진 복제본도 같은 날짜/제목/내용이면 같이 지운다.
        var sameNotice = wantedDate && wantedTitle &&
          summary === '[공지] ' + wantedTitle &&
          description.indexOf('공지일: ' + wantedDate) !== -1 &&
          (!wantedContent || description.indexOf(wantedContent) !== -1);
        if (!markerMatch && !sameNotice) return;
        try {
          Calendar.Events.remove(calendarId, event.id, { sendUpdates: 'none' });
          deleted++;
        } catch (removeError) {
          zLog_('오류', noticeId, channel, '실패', '공지 삭제: ' + removeError);
        }
      });
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (error) {
    zLog_('오류', noticeId, channel, '실패', '공지 검색/삭제: ' + error);
  }
  return deleted;
}

function zNotifyDeletedNotice_(notice) {
  var cleanTitle = zNoticeTimeText_(notice.title || '공지');
  var dateText = zNoticeDateText_(notice.writtenAt);
  var text = '[공지 취소]' + (dateText ? '\n날짜: ' + dateText : '') + '\n' + cleanTitle + (notice.content ? '\n\n' + notice.content : '');
  var sentChats = {};
  zMembers_().forEach(function(member) {
    var chatId = String(member.telegramId || '').trim();
    if (!chatId || sentChats[chatId]) return;
    sentChats[chatId] = true;
    try { zSendTelegram_(chatId, text); }
    catch (error) { zLog_('오류', notice.id, '텔레그램', '실패', '공지 취소 ' + member.name + ': ' + error); }
  });
}

function zSyncDeletedNotices_() {
  var indexSheet = zNoticeIndexSheet_();
  var current = {};
  var noticeSheet = zSpreadsheet_().getSheetByName(ZEPHYRUS.sheet.notice);
  if (noticeSheet && noticeSheet.getLastRow() >= 2) {
    var map = zHeaders_(noticeSheet);
    noticeSheet.getRange(2, 1, noticeSheet.getLastRow() - 1, noticeSheet.getLastColumn()).getValues().forEach(function(row, index) {
      var notice = zNoticeFromRow_(row, index + 2, map);
      if (notice.id && notice.title && notice.content) current[notice.id] = notice;
    });
  }

  var previous = indexSheet.getLastRow() < 2 ? [] : indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 4).getValues();
  var deleted = 0;
  previous.forEach(function(row) {
    var id = String(row[0] || '');
    if (!id || current[id]) return;
    var notice = {
      id: id,
      writtenAt: String(row[1] || ''),
      title: String(row[2] || ''),
      content: String(row[3] || '')
    };
    zNotifyDeletedNotice_(notice);
    deleted += zDeleteNoticeEventsInCalendar_(zTeamCalendarId_(), notice, '공용캘린더');
    zMembers_().forEach(function(member) {
      if (member.calendarId) deleted += zDeleteNoticeEventsInCalendar_(member.calendarId, notice, '개인캘린더');
    });
  });

  var rows = Object.keys(current).map(function(id) {
    return [
      id,
      zNoticeDateText_(current[id].writtenAt) || '',
      current[id].title || '',
      current[id].content || ''
    ];
  });
  var clearRows = Math.max(indexSheet.getLastRow() - 1, rows.length, 1);
  indexSheet.getRange(2, 1, clearRows, 4).clearContent();
  if (rows.length) indexSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return deleted;
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
  var time = schedule.time ? ' ' + schedule.time + (schedule.endTime ? '–' + schedule.endTime : '') : ' 종일';
  var title = zDisplayTitleForSchedule_(schedule);
  return '[' + schedule.date + time + '] ' + title + (schedule.location ? '\n장소: ' + schedule.location : '');
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

function zOwnedMemberCalendar_(member) {
  var name = '제피로스 일정 - ' + member.name;
  return CalendarApp.getAllOwnedCalendars().filter(function(calendar) { return calendar.getName() === name; })[0] || null;
}

function zGrantMemberCalendarAccess_(calendar, email) {
  var wanted = String(email || '').trim().toLowerCase();
  if (!wanted) return;
  var calendarId = calendar.getId();
  var rules = Calendar.Acl.list(calendarId).items || [];
  var existing = rules.filter(function(rule) {
    return rule.scope && rule.scope.type === 'user' && String(rule.scope.value || '').toLowerCase() === wanted;
  })[0] || null;
  var access = { role: 'writer', scope: { type: 'user', value: wanted } };
  if (existing) {
    // Do not downgrade the owner role to writer.
    if (existing.role !== 'writer' && existing.role !== 'owner') Calendar.Acl.update(access, calendarId, existing.id, { sendNotifications: true });
    return;
  }
  Calendar.Acl.insert(access, calendarId, { sendNotifications: true });
}

