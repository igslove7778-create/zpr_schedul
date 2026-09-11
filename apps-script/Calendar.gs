/**
 * 구글 캘린더 연동 (신규, 팀 공유 캘린더).
 * - [일정] 시트에 '종료시간' 열이 없으면 자동으로 추가한다. 비워두면 기본일정길이분(설정)만큼 길이로 처리.
 * - 10분마다 자동 실행 + [일정관리] 메뉴에서 수동 실행(캘린더 지금 동기화) 가능.
 * - '캘린더ID' 열에 이벤트ID를 저장해 재실행해도 중복 생성하지 않고 갱신한다.
 * - [일정] 상태가 '삭제'인데 캘린더ID가 남아있으면 해당 이벤트를 지운다.
 */

var END_TIME_HEADER_ = '종료시간';

function getTeamCalendarId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TEAM_CAL_ID');
  if (id) {
    try { CalendarApp.getCalendarById(id).getName(); return id; } catch (e) { /* 재생성 */ }
  }
  var cal = CalendarApp.createCalendar('제피로스 팀 일정');
  props.setProperty('TEAM_CAL_ID', cal.getId());
  return cal.getId();
}

/** [일정] 시트에 '종료시간' 헤더가 없으면 맨 끝열 다음에 추가 */
function ensureEndTimeColumn_() {
  var sh = sheet_(SHEET.SCHEDULE);
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(END_TIME_HEADER_) >= 0) return;
  sh.getRange(1, lastCol + 1).setValue(END_TIME_HEADER_).setFontWeight('bold');
}

function syncCalendar_() {
  var sh = sheet_(SHEET.SCHEDULE);
  if (!sh) return;
  ensureEndTimeColumn_();
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  function idx(name) { return headers.indexOf(name); }

  var iDate = idx(COL.DATE), iTime = idx(COL.TIME), iEnd = idx(END_TIME_HEADER_), iTitle = idx(COL.TITLE),
      iMemo = idx(COL.MEMO), iTarget = idx(COL.TARGET), iStatus = idx(COL.STATUS), iCal = idx(COL.CAL_ID);
  if (iDate < 0 || iTitle < 0 || iCal < 0) return;

  var calId = getTeamCalendarId_();
  var cal = CalendarApp.getCalendarById(calId);
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var settings = getSettings_();
  var defaultMin = Number((settings && settings['기본일정길이분']) || 60) || 60;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var title = row[iTitle];
    var sheetRow = r + 2;
    var status = iStatus >= 0 ? row[iStatus] : STATUS.NORMAL;
    var calEventId = row[iCal];

    if (!title || status === STATUS.DELETED) {
      if (calEventId) {
        try { cal.getEventById(calEventId).deleteEvent(); } catch (e) { }
        sh.getRange(sheetRow, iCal + 1).setValue('');
      }
      continue;
    }

    var dateVal = row[iDate];
    var dateObj = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(dateObj.getTime())) continue;

    var timeStr = iTime >= 0 ? String(row[iTime] || '').trim() : '';
    var endStr = iEnd >= 0 ? String(row[iEnd] || '').trim() : '';
    var start, end, allDay = false;

    if (timeStr) {
      var tp = timeStr.split(':').map(Number);
      start = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), tp[0] || 0, tp[1] || 0);
      if (endStr) {
        var ep = endStr.split(':').map(Number);
        end = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), ep[0] || 0, ep[1] || 0);
        if (end.getTime() <= start.getTime()) end = new Date(start.getTime() + defaultMin * 60000);
      } else {
        end = new Date(start.getTime() + defaultMin * 60000);
      }
    } else {
      allDay = true;
      start = dateObj;
    }

    var desc = (iMemo >= 0 ? String(row[iMemo] || '') : '') + (iTarget >= 0 && row[iTarget] ? ('\n대상: ' + row[iTarget]) : '');

    var ev = null;
    if (calEventId) { try { ev = cal.getEventById(calEventId); } catch (e) { ev = null; } }

    try {
      if (ev) {
        ev.setTitle(String(title));
        ev.setDescription(desc);
        if (allDay) { ev.setAllDayDate(start); } else { ev.setTime(start, end); }
      } else {
        var newEv = allDay
          ? cal.createAllDayEvent(String(title), start, { description: desc })
          : cal.createEvent(String(title), start, end, { description: desc });
        sh.getRange(sheetRow, iCal + 1).setValue(newEv.getId());
      }
    } catch (e) {
      log_('오류', row[idx(COL.ID)] || '', CHANNEL.CALENDAR, '시스템', '실패', '캘린더 동기화: ' + e);
    }
  }
}

function syncCalendarNow() {
  syncCalendar_();
  SpreadsheetApp.getActive().toast('캘린더 동기화 완료', '일정관리');
}

function setupCalendarTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncCalendar_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncCalendar_').timeBased().everyMinutes(10).create();
}

function aaaSetupCalendar() {
  var calId = getTeamCalendarId_();
  ensureEndTimeColumn_();
  setupCalendarTrigger_();
  syncCalendar_();
  Logger.log('팀 캘린더 ID: ' + calId);
  Logger.log('캘린더 링크: https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(calId));
}
