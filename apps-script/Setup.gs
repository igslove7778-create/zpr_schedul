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
    .addItem('편집 감지 트리거를 내 계정으로 설치', 'installEditTrigger')
    .addItem('이름 열 추가 (쉼표로 여러 명)', 'addNameColumns')
    .addItem('대상자 체크 다시 맞추기 (기존 행 전체)', 'resyncTargetChecks')
    .addItem('이름 열 인원을 구성원에 추가 + 인증코드 생성', 'addCheckColumnMembers')
    .addItem('인증코드 생성 (미연결 구성원)', 'generateAuthCodes')
    .addSeparator()
    .addItem('웹훅 등록', 'registerWebhook')
    .addItem('웹훅 상태 확인', 'showWebhookInfo')
    .addSeparator()
    .addItem('대기 알림 지금 발송', 'sendPendingNotifications')
    .addItem('선택 행 삭제 처리 (취소 알림)', 'deleteSelectedRows')
    .addItem('이체내역 지금 발송 (테스트)', 'sendTransfersNow')
    .addItem('캘린더 지금 동기화', 'syncCalendarNow')
    .addItem('구성원 캘린더 만들기·공유 (구글 계정 있는 구성원)', 'setupMemberCalendars')
    .addItem('캘린더 자동 동기화 켜기 (10분마다, 내 계정)', 'enableCalendarSync')
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

/**
 * [일정] 시트 사람 체크 열(리마인드 오른쪽 헤더)에 있는 이름 중 [구성원] 시트에 없는 사람을 자동 추가한다.
 * 권한=일반, 사용여부=Y 로 넣고, 이어서 인증코드를 발급한다. 각자 봇에 /start 후 코드를 입력하면 알림을 받기 시작한다.
 */
function addCheckColumnMembers() {
  var sch = sheet_(SHEET.SCHEDULE);
  var names = checkColumnNames_(headerMap_(sch));
  var mem = sheet_(SHEET.MEMBER);
  var hm = headerMap_(mem);
  var existing = getMembers_(true).map(function (m) { return m.name; });
  var added = [];
  names.forEach(function (n) {
    if (existing.indexOf(n) >= 0) return;
    var row = new Array(Math.max(mem.getLastColumn(), MEMBER_HEADERS.length)).fill('');
    row[hm[MEMBER_COL.NAME] - 1] = n;
    row[hm[MEMBER_COL.ROLE] - 1] = '일반';
    row[hm[MEMBER_COL.ACTIVE] - 1] = 'Y';
    mem.appendRow(row);
    added.push(n);
  });
  if (hm[MEMBER_COL.TG_ID]) mem.getRange(2, hm[MEMBER_COL.TG_ID], Math.max(mem.getLastRow() - 1, 1), 1).setNumberFormat('@');
  generateAuthCodes();
  SpreadsheetApp.getActive().toast((added.length ? '구성원 추가: ' + added.join(', ') : '새로 추가할 이름 없음') + ' / [구성원] 시트에서 각자 인증코드를 확인해 전달하세요.', '일정관리', 10);
}

/**
 * 편집 감지(handleEdit) 트리거를 현재 계정 소유로 설치한다 (메뉴).
 * 개인 Gmail 은 트리거 소유자 본인의 편집만 편집자 정보(e.user)가 넘어오므로,
 * 시트를 주로 편집하는 사람이 이 메뉴를 실행해 트리거를 본인 소유로 두는 것이 좋다.
 * 다른 계정에 같은 트리거가 남아 있으면 편집이 두 번 처리되므로, 그 계정에서는 handleEdit 트리거를 삭제해야 한다.
 */
function installEditTrigger() {
  var ss = ss_();
  var mine = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'handleEdit'; });
  if (mine.length) { SpreadsheetApp.getActive().toast('이미 내 계정 소유의 편집 감지 트리거가 있습니다.', '일정관리', 6); return; }
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  var me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) { /* 무시 */ }
  log_('설정', '', CHANNEL.SHEET, me || '시스템', '성공', '편집 감지 트리거 설치');
  SpreadsheetApp.getActive().toast('편집 감지 트리거를 ' + (me || '내 계정') + ' 소유로 설치했습니다. 다른 계정에 남은 handleEdit 트리거는 삭제하세요.', '일정관리', 10);
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
  // '전부' 칸을 직접 체크하면 전원 표시, 해제하면 전원 해제 (그 뒤 '대상자' 텍스트대로 재표시)
  if (editedHeader === COL.ALL) {
    applyAllColumn_(sh, hm, row, s);
    if (s.all && hm[COL.TARGET]) sh.getRange(row, hm[COL.TARGET]).setValue('전부').setBackground(null).clearNote();
    s = getSchedule_(row);
  }
  // 사람 체크 열을 직접 체크/해제하면 '대상자' 칸에 체크된 이름을 써 준다 (반대 방향 동기화)
  if (checkColumnNames_(hm).indexOf(editedHeader) >= 0) {
    syncTargetFromChecks_(sh, hm, row, s);
    s = getSchedule_(row);
  }

  var settings = getSettings_();
  var editor = getEditorName_(e); // 편집한 구성원 (이메일 대조). 못 찾으면 설정 '담당자명' 또는 첫 관리자
  if (!s.id) {
    // 신규 등록
    var id = newId_();
    setCell_(sh, hm, row, COL.ID, id);
    setCell_(sh, hm, row, COL.REGISTRANT, s.registrant || editor);
    setCell_(sh, hm, row, COL.CHANNEL, CHANNEL.SHEET);
    setCell_(sh, hm, row, COL.STATUS, STATUS.NORMAL);
    setCell_(sh, hm, row, COL.UPDATED, nowString_());
    if (!s.alarm) setCell_(sh, hm, row, COL.ALARM, ALARM.PENDING);
    sh.getRange(row, hm[COL.TIME]).setNumberFormat('@');
    log_('등록', id, CHANNEL.SHEET, editor, '성공', s.date + ' ' + s.time + ' ' + s.title);
    s = getSchedule_(row);
    if (settings['즉시알림'] !== 'N' && s.alarm === ALARM.PENDING) notifySheetRegistered_(s);
    return;
  }

  // 기존 행 수정
  setCell_(sh, hm, row, COL.UPDATED, nowString_());
  var watched = [COL.DATE, COL.TIME, COL.TITLE, COL.MEMO, COL.TARGET, COL.ALL].concat(checkColumnNames_(hm));
  if (s.status !== STATUS.DELETED && s.alarm === ALARM.SENT && watched.indexOf(editedHeader) >= 0) {
    var before = '';
    if (e && e.oldValue !== undefined && (editedHeader === COL.DATE || editedHeader === COL.TIME)) before = String(e.oldValue);
    log_('수정', s.id, CHANNEL.SHEET, editor, '성공', editedHeader + ' 변경');
    notifyChanged_(s, before);
  } else if (s.alarm === ALARM.PENDING && settings['즉시알림'] !== 'N') {
    notifySheetRegistered_(s);
  }
}

/**
 * [일정] 시트 맨 오른쪽에 사람 체크 열을 추가한다 (메뉴). 이름을 쉼표·공백으로 여러 명 입력 가능.
 * 이미 같은 이름의 헤더가 있으면 건너뛴다. 헤더 서식과 o/O/ㅇ/○ 드롭다운을 함께 넣는다.
 */
function addNameColumns() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('이름 열 추가', '추가할 이름을 쉼표나 공백으로 구분해 입력하세요.\n예) 김유선, 성도형, 조인희', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var names = String(res.getResponseText() || '').split(/[,\s]+/).map(function (t) { return t.trim(); }).filter(function (t) { return t; });
  if (!names.length) return;
  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  var added = [], skipped = [];
  names.forEach(function (n) {
    if (hm[n]) { skipped.push(n); return; }
    var col = sh.getLastColumn() + 1;
    if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
    sh.getRange(1, col).setValue(n).setFontWeight('bold').setBackground('#DCE6F1');
    sh.setColumnWidth(col, 70);
    sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['o', 'O', 'ㅇ', '○'], true).setAllowInvalid(true).build())
      .setHorizontalAlignment('center');
    hm = headerMap_(sh);
    added.push(n);
  });
  SpreadsheetApp.getActive().toast('추가: ' + (added.join(', ') || '없음') + (skipped.length ? ' / 이미 있음: ' + skipped.join(', ') : ''), '일정관리', 8);
}

/**
 * 기존 행 전체의 '대상자' 텍스트와 사람 체크 열을 다시 맞춘다 (메뉴).
 * - '대상자' 텍스트가 있으면 텍스트 기준으로 체크 열을 표시
 * - 텍스트가 없고 체크만 있으면 체크 기준으로 '대상자' 텍스트를 채움
 * 알림은 보내지 않는다 (표시만 맞춤).
 */
function resyncTargetChecks() {
  var sh = sheet_(SHEET.SCHEDULE);
  var hm = headerMap_(sh);
  if (!hm[COL.TARGET]) { SpreadsheetApp.getUi().alert("[일정] 시트에 '대상자' 열이 없습니다."); return; }
  var n = 0;
  getAllSchedules_().forEach(function (s) {
    if (s.status === STATUS.DELETED) return;
    if (s.all && !s.target) { applyAllColumn_(sh, hm, s.row, s); sh.getRange(s.row, hm[COL.TARGET]).setValue('전부'); n++; return; }
    if (s.target) { applyTargetColumn_(sh, hm, s.row, s); n++; return; }
    if (s.checked.length) { syncTargetFromChecks_(sh, hm, s.row, s); n++; }
  });
  SpreadsheetApp.getActive().toast(n + '행의 대상자 체크를 다시 맞췄습니다.', '일정관리', 6);
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
