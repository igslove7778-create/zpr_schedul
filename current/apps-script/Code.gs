/**
 * Zephyrus schedule recovery
 *
 * This is a safe replacement for the missing Apps Script project. It is meant
 * to be bound to the existing Zephyrus spreadsheet. It never creates a team
 * calendar automatically: TEAM_CAL_ID must be set in Script Properties first.
 */

var ZEPHYRUS = {
  // These are the confirmed Zephyrus recovery targets.  Script Properties can
  // override them later, but this script will never create another calendar.
  recoverySpreadsheetId: '1b9ACCLP6mr9V_xTLUhATsPbYSwJovUa97U44D4WjbYA',
  recoveryTeamCalendarId: 'b5166331d70ed13b91ae0ad2150e0240a30a8756515efe1de22a33f66cee8395@group.calendar.google.com',
  sheet: {
    schedule: '일정',
    members: '구성원',
    settings: '설정',
    log: '로그',
    notice: '공지사항',
    syncIndex: '제피로스_연동백업',
    noticeIndex: '제피로스_공지백업',
    fileIndex: '제피로스_파일연동'
  },
  col: {
    date: '날짜',
    time: '시작시간',
    endTime: '종료시간',
    title: '일정',
    location: '장소',
    targets: '대상자',
    alarm: '알람',
    file: '파일',
    id: '일정ID',
    registrant: '등록자',
    channel: '등록채널',
    // The old live sheet already uses this column.  Keeping it avoids making
    // a second set of calendar links while recovering the project.
    calendarEvent: '캘린더ID',
    updated: '수정일시',
    status: '상태',
    reminders: '리마인드'
  },
  noticeCol: {
    writtenAt: '작성일시',
    title: '공지제목',
    content: '공지내용',
    alarm: '알림',
    id: '공지ID',
    registrant: '등록자',
    updated: '수정일시',
    status: '상태',
    teamEvent: '공용캘린더ID'
  },
  memberCol: {
    name: '이름',
    email: '구글 계정',
    // The live sheet already has this header with a space.  Reuse it so
    // existing Telegram connections continue to work.
    telegramId: '텔레그램 ID',
    role: '권한',
    active: '사용여부',
    code: '인증코드',
    calendarId: '캘린더ID'
  },
  status: { normal: '정상', deleted: '삭제' },
  alarm: { pending: '대기', sent: '발송완료', failed: '실패' },
  kst: 'Asia/Seoul',
  markerPrefix: '[ZEPHYRUS:'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('제피로스 복구')
    .addItem('1. 시트 구조 점검', 'zPrepareSystem')
    .addItem('2. 연결 상태 확인', 'zShowStatus')
    .addSeparator()
    .addItem('시트 수정 감지 켜기', 'zInstallEditTrigger')
    .addItem('텔레그램 리마인드 켜기 (1분)', 'zInstallReminderTrigger')
    .addItem('시트에서 캘린더 동기화 켜기 (1분)', 'zInstallCalendarSyncTrigger')
    .addItem('캘린더에서 시트 반영 켜기', 'zInstallCalendarReverseTrigger')
    .addItem('자동 실행 4개 한꺼번에 켜기', 'zInstallCoreTriggers')
    .addSeparator()
    .addItem('지금 시트에서 캘린더로 동기화', 'zRunCalendarSyncNow')
    .addItem('지금 리마인드 확인', 'runReminders')
    .addItem('인증코드 만들기', 'zGenerateAuthCodes')
    .addItem('텔레그램 웹훅 등록', 'zRegisterWebhook')
    .addSeparator()
    .addItem('직원 전용 캘린더 만들기/공유', 'zProvisionMemberCalendars')
    .addItem('직원 캘린더 양방향 동기화 켜기', 'zEnableMemberCalendarSync')
    .addItem('대상자 체크칸 만들기', 'zPrepareMemberCheckColumns')
    .addItem('공지사항 탭 만들기 / 일반 로그 정리', 'zApplyNoticeAndLogFinalization')
    .addItem('기존 공지 공용캘린더에 반영', 'zSyncNoticeTeamCalendar')
    .addItem('삭제된 일정의 개인캘린더 복사본 정리', 'zRemoveDeletedScheduleCopies')
    .addToUi();
}

function zProps_() {
  return PropertiesService.getScriptProperties();
}

function zProp_(key) {
  return zProps_().getProperty(key) || '';
}

function zRequireProp_(key) {
  var value = zProp_(key).trim();
  if (!value) {
    throw new Error('스크립트 속성 ' + key + ' 값이 비어 있습니다.');
  }
  return value;
}

function zSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  return SpreadsheetApp.openById(zProp_('SPREADSHEET_ID') || ZEPHYRUS.recoverySpreadsheetId);
}

function zTeamCalendarId_() {
  return zProp_('TEAM_CAL_ID') || ZEPHYRUS.recoveryTeamCalendarId;
}

function zTeamCalendar_() {
  var id = zTeamCalendarId_();
  var calendar = CalendarApp.getCalendarById(id);
  // The handover contained one stale calendar id.  If that id is no longer
  // visible, adopt only one calendar with the exact confirmed team name.
  // This never creates a calendar and refuses to guess when there are two.
  if (!calendar) {
    var matches = CalendarApp.getAllCalendars().filter(function(item) {
      return item.getName() === '제피로스 팀 일정';
    });
    // In the captured account, several old self-owned copies have the same
    // name.  The real team calendar is the single shared calendar under
    // "Other calendars", so prefer one exact-name calendar not owned by this
    // account.  We still refuse to guess if more than one shared copy exists.
    var sharedMatches = matches.filter(function(item) { return !item.isOwnedByMe(); });
    var candidate = sharedMatches.length === 1 ? sharedMatches[0] : (matches.length === 1 ? matches[0] : null);
    if (candidate) {
      calendar = candidate;
      zProps_().setProperty('TEAM_CAL_ID', calendar.getId());
    }
  }
  if (!calendar) {
    throw new Error('제피로스 팀 일정 캘린더에 접근할 수 없습니다. 새 캘린더를 만들지 않았습니다.');
  }
  return calendar;
}

function zNowText_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy-MM-dd HH:mm:ss');
}

// This recovery project may be created as a standalone project.  In that
// case there is no spreadsheet UI to show a popup in, so use the execution
// log instead of treating a completed setup as an error.
function zNotice_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (ignore) { console.log(message); }
}

function zToday_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy-MM-dd');
}

function zId_() {
  return Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyyMMddHHmmss') + '-' + Math.random().toString(36).slice(2, 7);
}

function zNormalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, ZEPHYRUS.kst, 'yyyy-MM-dd');
  }
  var text = String(value || '').trim();
  var match = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (match) return zDateTextFromParts_(match[1], match[2], match[3]);

  // The original sheet contains date-like text such as "9월 16일" and
  // "9-20".  Treat these as dates in the current Korean calendar year so
  // old and new schedules can share one real date column and sort correctly.
  match = text.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/) ||
    text.match(/^(\d{1,2})[-./](\d{1,2})$/);
  if (!match) return '';
  return zDateTextFromParts_(Utilities.formatDate(new Date(), ZEPHYRUS.kst, 'yyyy'), match[1], match[2]);
}

function zDateTextFromParts_(year, month, day) {
  var y = Number(year);
  var m = Number(month);
  var d = Number(day);
  // 03:00 UTC is noon in Korea, which avoids a date shifting at midnight.
  var date = new Date(Date.UTC(y, m - 1, d, 3, 0, 0));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return '';
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}

function zDateObject_(ymd) {
  if (ymd instanceof Date && !isNaN(ymd.getTime())) return ymd;
  var normalized = zNormalizeDate_(ymd);
  var parts = normalized.split('-').map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) return null;
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 3, 0, 0));
}

function zNormalizeTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, ZEPHYRUS.kst, 'HH:mm');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  var match = text.match(/^(\d{1,2}):(\d{2})$/);
  var hour;
  var minute;
  if (match) {
    hour = Number(match[1]);
    minute = Number(match[2]);
  } else {
    // Sheet users may write natural Korean time text.  A bare "2시" is
    // treated as 02:00; use "오후 2시" when the afternoon is intended.
    match = text.match(/^(오전|오후)\s*(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분?)?$/);
    if (match) {
      hour = Number(match[2]);
      minute = Number(match[3] || 0);
      if (hour < 1 || hour > 12) return '';
      if (match[1] === '오전' && hour === 12) hour = 0;
      if (match[1] === '오후' && hour < 12) hour += 12;
    } else {
      match = text.match(/^(\d{1,2})\s*시(?:\s*(\d{1,2})\s*분?)?$/);
      if (!match) return '';
      hour = Number(match[1]);
      minute = Number(match[2] || 0);
    }
  }
  if (hour > 23 || minute > 59) return '';
  return ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2);
}

