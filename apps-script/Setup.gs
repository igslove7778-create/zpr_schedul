/**
 * 초기 설정·메뉴·트리거 (ADM-01, ADM-02, ADM-04, REG-01, REG-06).
 *
 * 설치 순서는 docs/설치가이드.md 참조.
 *  1. 스크립트 속성에 BOT_TOKEN 저장
 *  2. setup() 실행 -> 시트·트리거 생성
 *  3. 웹앱 배포 후 URL 을 스크립트 속성 WEB_APP_URL 에 저장 -> registerWebhook() 실행
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('일정관리')
    .addItem('초기 설정 (시트·트리거 생성)', 'setup')
    .addItem('구성원 열 동기화', 'syncMemberColumns')
    .addItem('인증코드 생성 (미연결 구성원)', 'generateAuthCodes')
    .addSeparator()
    .addItem('웹훅 등록', 'registerWebhook')
    .addItem('웹훅 상태 확인', 'showWebhookInfo')
    .addSeparator()
    .addItem('대기 알림 지금 발송', 'sendPendingNotifications')
    .addItem('선택 행 삭제 처리 (취소 알림)', 'deleteSelectedRows')
    .addItem('이체내역 지금 발송 (테스트)', 'sendTransfersNow')
    .addItem('캘린더 지금 동기화', 'syncCalendarNow')
    .addToUi();
}

function setup() {
  var ss = ss_();
  // [구성원]
  var mem = ss.getSheetByName(SHEET.MEMBER) || ss.insertSheet(SHEET.MEMBER);
  if (mem.getLastRow() === 0) {
    mem.appendRow(MEMBER_HEADERS);
    mem.getRange(1, 1, 1, MEMBER_HEADERS.length).setFontWeight('bold').setBackground('#DCE6F1');
    mem.getRange('D2:D200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['관리자', '일반'], true).build());
    mem.getRange('E2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).build());
    mem.getRange('C2:C200').setNumberFormat('@');
    mem.setFrozenRows(1);
  }
  // [설정]
  var set = ss.getSheetByName(SHEET.SETTING) || ss.insertSheet(SHEET.SETTING);
  if (set.getLastRow() === 0) {
    set.appendRow(['키', '값', '설명']);
    set.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#DCE6F1');
    DEFAULT_SETTINGS.forEach(function (r) { set.appendRow(r); });
    set.getRange('B2:B100').setNumberFormat('@');
    set.setColumnWidth(3, 520);
    set.setFrozenRows(1);
  }
  // [로그]
  var log = ss.getSheetByName(SHEET.LOG) || ss.insertSheet(SHEET.LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(LOG_HEADERS);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#DCE6F1');
    log.setFrozenRows(1);
  }
  // [일정]
  var sch = ss.getSheetByName(SHEET.SCHEDULE) || ss.insertSheet(SHEET.SCHEDULE, 0);
  if (sch.getLastRow() === 0) {
    var headers = [COL.DATE, COL.TIME, COL.TITLE, COL.MEMO, COL.TARGET, COL.ALARM, '', COL.ALL].concat(SYSTEM_COLS);
    sch.appendRow(headers);
    sch.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#DCE6F1');
    var hIdx = {}; headers.forEach(function (h, i) { if (h) hIdx[h] = i + 1; });
    var blankIdx = headers.indexOf('') + 1;
    sch.getRange(2, hIdx[COL.DATE], 999, 1).setNumberFormat('MM"월" dd"일"');
    sch.getRange(2, hIdx[COL.TIME], 999, 1).setNumberFormat('@');
    sch.getRange(2, hIdx[COL.ALARM], 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([ALARM.PENDING, ALARM.SENT, ALARM.FAILED], true).build());
    sch.getRange(1, hIdx[COL.TARGET]).setNote('이름을 공백이나 쉼표로 구분해서 입력하면 아래 구성원 열에 자동으로 o가 표시됩니다 (예: 김유선 성도형). "전체"라고 쓰면 전원 체크. 비워두면 체크 칸을 직접 눌러도 됩니다.');
    sch.setColumnWidth(hIdx[COL.TITLE], 260);
    sch.setColumnWidth(hIdx[COL.TARGET], 140);
    sch.setColumnWidth(blankIdx, 20);
    sch.setFrozenRows(1);
  }
  // [이체내역]
  var trf = ss.getSheetByName(SHEET.TRANSFER) || ss.insertSheet(SHEET.TRANSFER);
  if (trf.getLastRow() === 0) {
    trf.appendRow(TRANSFER_HEADERS);
    trf.getRange(1, 1, 1, TRANSFER_HEADERS.length).setFontWeight('bold').setBackground('#DCE6F1');
    trf.appendRow([28, '(예시) 유포리아지식산업센터', '지역농협', '3010277538481', 500000, '제주도 관리비 - 이 줄은 지우고 실제 내용을 넣으세요']);
    trf.getRange(2, 1, 1, TRANSFER_HEADERS.length).setFontColor('#999999').setFontStyle('italic');
    trf.getRange('A2:A500').setNumberFormat('0"일"');
    trf.getRange('D2:D500').setNumberFormat('@');
    trf.getRange('E2:E500').setNumberFormat('#,##0"원"');
    trf.setColumnWidth(2, 160);
    trf.setColumnWidth(6, 220);
    trf.setFrozenRows(1);
  }
  syncMemberColumns();
  installTriggers_();
  SpreadsheetApp.getActive().toast('초기 설정 완료. [구성원] 시트를 채운 뒤 인증코드를 생성하세요.', '일정관리', 8);
}

/** [구성원] 이름을 [일정] 시트 '전부' 오른쪽에 열로 생성 (ADM-01). 시스템 열은 항상 맨 오른쪽·숨김 */
function syncMemberColumns() {
  var sch = sheet_(SHEET.SCHEDULE);
  var names = getMembers_().map(function (m) { return m.name; });
  var hm = headerMap_(sch);
  var existing = hm.__headers;
  var sysStart = existing.indexOf(COL.ID) + 1; // 1-based
  if (sysStart === 0) throw new Error('[일정] 시트에 시스템 열(일정ID)이 없습니다. 초기 설정을 실행하세요.');
  names.forEach(function (n) {
    if (hm[n]) return;
    sch.insertColumnBefore(sysStart);
    sch.getRange(1, sysStart).setValue(n).setFontWeight('bold').setBackground('#DCE6F1');
    sch.setColumnWidth(sysStart, 70);
    sysStart++;
    hm = headerMap_(sch);
  });
  // 체크 열 데이터 확인 (AUTH-02): o 로 통일, 다른 체크 표기도 허용
  var allCol = hm[COL.ALL];
  var lastMemberCol = hm[COL.ID] - 1;
  if (lastMemberCol >= allCol) {
    sch.getRange(2, allCol, 999, lastMemberCol - allCol + 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['o', 'O', 'ㅇ', '○'], true).setAllowInvalid(true).build())
      .setHorizontalAlignment('center');
  }
  // 시스템 열 숨김
  hm = headerMap_(sch);
  sch.hideColumns(hm[COL.ID], SYSTEM_COLS.length);
  sch.getRange(1, hm[COL.ID], 1, SYSTEM_COLS.length).setBackground('#EEEEEE');
}

function installTriggers_() {
  var ss = ss_();
  var have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have.handleEdit) ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  if (!have.runReminders) ScriptApp.newTrigger('runReminders').timeBased().everyMinutes(10).create();
}

/** 텔레그램 ID 가 없는 구성원에게 6자리 인증코드 발급 (ADM-02) */
function generateAuthCodes() {
  var sh = sheet_(SHEET.MEMBER);
  var hm = headerMap_(sh);
  var n = 0;
  getMembers_().forEach(function (m) {
    if (m.tgId || m.code) return;
    var code = String(Math.floor(100000 + Math.random() * 900000));
    sh.getRange(m.row, hm[MEMBER_COL.CODE]).setValue(code);
    n++;
  });
  SpreadsheetApp.getActive().toast(n + '명에게 인증코드 발급. 각자에게 코드를 전달하고 봇에 /start 후 입력하도록 안내하세요.', '일정관리', 8);
}

function registerWebhook() {
  var url = getProp_('WEB_APP_URL');
  var key = getProp_('WEBHOOK_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    PropertiesService.getScriptProperties().setProperty('WEBHOOK_KEY', key);
  }
  if (!url) throw new Error('스크립트 속성 WEB_APP_URL 이 없습니다. 웹앱 배포 URL 을 먼저 저장하세요.');
  var r = tgCall_('setWebhook', { url: url + '?key=' + key, allowed_updates: ['message'], drop_pending_updates: true });
  SpreadsheetApp.getActive().toast('웹훅 등록 결과: ' + JSON.stringify(r), '일정관리', 8);
}

function showWebhookInfo() {
  var r = tgCall_('getWebhookInfo', {});
  SpreadsheetApp.getUi().alert(JSON.stringify(r, null, 2));
}

// ---------- 시트 직접 입력 (REG-01, REG-06) - 설치형 onEdit 트리거 ----------

function handleEdit(e) {
  try {
    var range = e.range;
    var sh = range.getSheet();
    if (sh.getName() !== SHEET.SCHEDULE || range.getRow() < 2) return;
    var hm = headerMap_(sh);
    var editedCol = range.getColumn();
    var editedHeader = hm.__headers[editedCol - 1] || '';
    if (SYSTEM_COLS.indexOf(editedHeader) >= 0 && editedHeader !== COL.STATUS) return; // 시스템 열 수동 편집은 무시

    for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) processSheetRow_(sh, hm, r, editedHeader, e);
  } catch (err) {
    log_('오류', '', CHANNEL.SHEET, '', '실패', 'handleEdit: ' + err + ' ' + (err.stack || ''));
  }
}

function processSheetRow_(sh, hm, row, editedHeader, e) {
  var s = getSchedule_(row);
  var dateCell = sh.getRange(row, hm[COL.DATE]);
  var titleCell = sh.getRange(row, hm[COL.TITLE]);
  var rowHasData = s.date || s.title || s.time || s.memo || s.target;
  if (!rowHasData) return;

  // 삭제 처리 (상태=삭제)
  if (editedHeader === COL.STATUS && s.status === STATUS.DELETED && s.id) {
    setCell_(sh, hm, row, COL.UPDATED, nowString_());
    notifyDeleted_(s);
    return;
  }

  // 필수값 검사: 누락 시 셀 강조, 동기화 제외
  var invalid = !s.date || !s.title || !/^\d{4}-\d{2}-\d{2}$/.test(s.date) || (s.time && !/^\d{1,2}:\d{2}$/.test(s.time));
  dateCell.setBackground(!s.date || !/^\d{4}-\d{2}-\d{2}$/.test(s.date) ? '#F8CBAD' : null);
  titleCell.setBackground(!s.title ? '#F8CBAD' : null);
  if (s.time) sh.getRange(row, hm[COL.TIME]).setBackground(!/^\d{1,2}:\d{2}$/.test(s.time) ? '#F8CBAD' : null);
  if (invalid) return;

  // 대상자 열 반영 (편의 기능) - 새 행이거나 '대상자' 칸 자체가 수정됐을 때만 체크 칸을 새로 계산한다
  if (hm[COL.TARGET] && (editedHeader === COL.TARGET || !s.id)) {
    applyTargetColumn_(sh, hm, row, s);
    s = getSchedule_(row);
  }

  var settings = getSettings_();
  if (!s.id) {
    // 신규 등록
    var id = newId_();
    setCell_(sh, hm, row, COL.ID, id);
    setCell_(sh, hm, row, COL.REGISTRANT, s.registrant || getOperatorName_());
    setCell_(sh, hm, row, COL.CHANNEL, CHANNEL.SHEET);
    setCell_(sh, hm, row, COL.STATUS, STATUS.NORMAL);
    setCell_(sh, hm, row, COL.UPDATED, nowString_());
    if (!s.alarm) setCell_(sh, hm, row, COL.ALARM, ALARM.PENDING);
    sh.getRange(row, hm[COL.TIME]).setNumberFormat('@');
    log_('등록', id, CHANNEL.SHEET, getOperatorName_(), '성공', s.date + ' ' + s.time + ' ' + s.title);
    s = getSchedule_(row);
    if (settings['즉시알림'] !== 'N' && s.alarm === ALARM.PENDING) notifyImmediate_(s, null);
    return;
  }

  // 기존 행 수정
  setCell_(sh, hm, row, COL.UPDATED, nowString_());
  var watched = [COL.DATE, COL.TIME, COL.TITLE, COL.MEMO, COL.TARGET, COL.ALL].concat(getMembers_().map(function (m) { return m.name; }));
  if (s.status !== STATUS.DELETED && s.alarm === ALARM.SENT && watched.indexOf(editedHeader) >= 0) {
    var before = '';
    if (e && e.oldValue !== undefined && (editedHeader === COL.DATE || editedHeader === COL.TIME)) before = String(e.oldValue);
    log_('수정', s.id, CHANNEL.SHEET, getOperatorName_(), '성공', editedHeader + ' 변경');
    notifyChanged_(s, before);
  } else if (s.alarm === ALARM.PENDING && settings['즉시알림'] !== 'N') {
    notifyImmediate_(s, null);
  }
}

/** 선택한 행을 상태=삭제 로 처리 (행 삭제 금지 원칙) */
function deleteSelectedRows() {
  var sh = ss_().getActiveSheet();
  if (sh.getName() !== SHEET.SCHEDULE) { SpreadsheetApp.getUi().alert('[일정] 시트에서 행을 선택하세요.'); return; }
  var hm = headerMap_(sh);
  var range = sh.getActiveRange();
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r < 2) continue;
    var s = getSchedule_(r);
    if (!s.id || s.status === STATUS.DELETED) continue;
    setCell_(sh, hm, r, COL.STATUS, STATUS.DELETED);
    setCell_(sh, hm, r, COL.UPDATED, nowString_());
    sh.getRange(r, 1, 1, hm[COL.ID] - 1).setFontLine('line-through').setFontColor('#999999');
    notifyDeleted_(s);
  }
}
