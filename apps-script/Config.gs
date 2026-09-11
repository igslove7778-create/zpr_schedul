/**
 * 제피로스 통합 일정관리 - 공통 설정
 * 시트명·헤더명은 기능명세서 v0.3 3장을 따른다.
 */
var SHEET = { SCHEDULE: '일정', MEMBER: '구성원', LOG: '로그', SETTING: '설정' };

var COL = {
  DATE: '날짜', TIME: '시간', TITLE: '일정', MEMO: '메모', ALARM: '알람', ALL: '전부',
  ID: '일정ID', REGISTRANT: '등록자', CHANNEL: '등록채널', CAL_ID: '캘린더ID',
  UPDATED: '수정일시', STATUS: '상태', REMIND: '리마인드'
};
var SYSTEM_COLS = [COL.ID, COL.REGISTRANT, COL.CHANNEL, COL.CAL_ID, COL.UPDATED, COL.STATUS, COL.REMIND];

var MEMBER_COL = { NAME: '이름', EMAIL: '구글 계정', TG_ID: '텔레그램 ID', ROLE: '권한', ACTIVE: '사용여부', CODE: '인증코드' };
var MEMBER_HEADERS = [MEMBER_COL.NAME, MEMBER_COL.EMAIL, MEMBER_COL.TG_ID, MEMBER_COL.ROLE, MEMBER_COL.ACTIVE, MEMBER_COL.CODE];
var LOG_HEADERS = ['일시', '구분', '일정ID', '채널', '처리자', '결과', '상세'];

var ALARM = { PENDING: '대기', SENT: '발송완료', FAILED: '실패' };
var STATUS = { NORMAL: '정상', DELETED: '삭제' };
var CHANNEL = { SHEET: '시트', TELEGRAM: '텔레그램', CALENDAR: '캘린더', WEB: '웹' };
var CHECK_VALUES = ['o', 'O', 'ㅇ', '○', 'TRUE', true];

// [설정] 시트 기본값. 키 / 값 / 설명
var DEFAULT_SETTINGS = [
  ['담당자명', '', '시트 직접 입력 건의 등록자로 기록할 이름 ([구성원] 이름과 동일). 비우면 권한=관리자인 첫 구성원'],
  ['즉시알림', 'Y', '등록 즉시 대상자에게 텔레그램 알림 (Y/N)'],
  ['당일리마인드시각', '08:00', '당일 일정 리마인드 발송 시각 (HH:mm). 비우면 미발송'],
  ['사전리마인드분', '30', '시간이 있는 일정의 N분 전 리마인드. 0이면 미발송'],
  ['오후해석시작시', '0', '오전/오후 없이 입력된 시각이 이 값 이상 12 이하이면 오후로 해석 (D-01). 0이면 24시간제 그대로'],
  ['기본일정길이분', '60', '캘린더 이벤트 기본 길이 (2단계)'],
  ['재시도횟수', '3', 'API 실패 시 재시도 횟수'],
  ['봇사용법', '일정 {날짜} [{시간}] {일정명} [@대상자 ...]\n예) 일정 6/6 18:30 전체회의 @전부\n예) 일정 9월 15일 방산IR 제출 @김유선 @성도형', '사용법 안내 메시지']
];

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function getSettings_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('settings');
  if (cached) return JSON.parse(cached);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTING);
  var out = {};
  DEFAULT_SETTINGS.forEach(function (r) { out[r[0]] = r[1]; });
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '') out[String(r[0])] = r[1] === '' ? out[String(r[0])] : String(r[1]);
    });
  }
  cache.put('settings', JSON.stringify(out), 60);
  return out;
}

function isChecked_(v) {
  if (v === true) return true;
  var s = String(v).trim();
  return CHECK_VALUES.indexOf(s) >= 0;
}

function nowString_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function newId_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMddHHmmss') + '-' + Math.random().toString(36).slice(2, 6);
}
