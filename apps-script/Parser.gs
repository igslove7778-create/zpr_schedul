/**
 * 텔레그램 입력 인식 (PAR-01 ~ PAR-06). 기능명세서 v0.3 5장.
 * Apps Script API를 사용하지 않는 순수 함수 - Node 테스트(tests/parser.test.js)로 검증한다.
 *
 * parseScheduleCommand(text, opts)
 *   opts.now        : Date (기준 시각, 기본 현재)
 *   opts.members    : ['김유선','성도형',...] (@대상자 검증용)
 *   opts.pmFrom     : 0 이면 24시간제 그대로, N(1~11)이면 오전/오후 없는 N~11시를 오후로 해석 (D-01)
 * 반환 { ok, date:'yyyy-MM-dd', time:'HH:mm'|'', allDay, title, targets:['전부']|[이름...], error, message }
 */
function parseScheduleCommand(text, opts) {
  opts = opts || {};
  var raw = String(text || '').replace(/\s+/g, ' ').trim();
  if (!/^일정(\s|$)/.test(raw)) return fail_('NOT_COMMAND', '"일정"으로 시작해야 합니다.');
  var tokens = raw.split(' ').slice(1);
  var r = parseDateTimeTitleTargets_(tokens, opts);
  if (!r.ok) return r;
  return { ok: true, date: r.date, time: r.time, allDay: r.allDay, title: r.title, targets: r.targets, error: null, message: '' };
}

/**
 * 날짜·시간·일정명·@대상자 토큰을 해석하는 공통 로직 (등록·수정 명령이 공유).
 * opts.requireDate / opts.requireTitle 기본 true (수정 명령에서는 false 로 넘겨 부분 변경을 허용한다).
 * 반환 { ok, date, dateFound, time, timeFound, allDay, title, targets, error, message }
 */
function parseDateTimeTitleTargets_(tokens, opts) {
  opts = opts || {};
  var now = opts.now || new Date();
  var members = opts.members || [];
  var pmFrom = Number(opts.pmFrom || 0);
  var requireDate = opts.requireDate !== false;
  var requireTitle = opts.requireTitle !== false;

  // 1) @대상자 (PAR-06)
  var targets = [];
  var rest = [];
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (t.charAt(0) === '@' && t.length > 1) {
      var name = t.slice(1);
      if (name === '전부' || name === '전체' || name === 'all') { targets.push('전부'); continue; }
      if (members.length && members.indexOf(name) < 0) return fail_('UNKNOWN_TARGET', '알 수 없는 대상자: @' + name);
      if (targets.indexOf(name) < 0) targets.push(name);
      continue;
    }
    rest.push(t);
  }
  if (targets.indexOf('전부') >= 0) targets = ['전부'];

  // 2) 날짜 (PAR-02, PAR-05)
  var d = extractDate_(rest, now);
  if (d.error) return fail_(d.error, d.message);
  if (requireDate && !d.found) return fail_('NO_DATE', '날짜를 인식하지 못했습니다.');
  if (d.found) rest = d.rest;

  // 3) 시간 (PAR-03, PAR-05)
  var tm = extractTime_(rest, pmFrom);
  if (tm.error) return fail_(tm.error, tm.message);
  rest = tm.rest;

  // 4) 일정명 (PAR-04)
  var title = rest.join(' ').trim();
  if (requireTitle && !title) return fail_('NO_TITLE', '일정명이 없습니다.');

  return {
    ok: true, date: d.found ? d.date : '', dateFound: d.found,
    time: tm.time || '', timeFound: !!tm.time, allDay: !tm.time,
    title: title, targets: targets, error: null, message: ''
  };
}

/**
 * 텔레그램 수정 명령: "변경 {번호} {바꿀 날짜·시간·일정명·@대상자}" (전부 생략 가능, 최소 하나는 있어야 함)
 * 번호는 "오늘 일정"·"이번주 일정" 목록에서 보여준 순번을 가리킨다 (Telegram.gs 에서 매핑).
 * 반환 { ok, index, date, dateFound, time, timeFound, title, targets, error, message }
 */
function parseEditCommand(text, opts) {
  opts = opts || {};
  var raw = String(text || '').replace(/\s+/g, ' ').trim();
  var m = /^(변경|수정)\s+(\d+)\s*(.*)$/.exec(raw);
  if (!m) return fail_('NOT_EDIT_COMMAND', '"변경 {번호} {바꿀 내용}" 형식이어야 합니다.');
  var idx = Number(m[2]);
  var restText = (m[3] || '').trim();
  var tokens = restText ? restText.split(' ') : [];
  var r = parseDateTimeTitleTargets_(tokens, {
    now: opts.now, members: opts.members, pmFrom: opts.pmFrom, requireDate: false, requireTitle: false
  });
  if (!r.ok) return r;
  if (!r.dateFound && !r.timeFound && !r.title && !r.targets.length) {
    return fail_('NO_CHANGE', '변경할 내용이 없습니다. 예) 변경 2 15:00');
  }
  return {
    ok: true, index: idx, date: r.date, dateFound: r.dateFound,
    time: r.time, timeFound: r.timeFound, title: r.title, targets: r.targets,
    error: null, message: ''
  };
}

/** 텔레그램 삭제 명령: "취소 {번호}" 또는 "삭제 {번호}" */
function parseCancelCommand(text) {
  var raw = String(text || '').replace(/\s+/g, ' ').trim();
  var m = /^(취소|삭제)\s+(\d+)$/.exec(raw);
  if (!m) return fail_('NOT_CANCEL_COMMAND', '"취소 {번호}" 형식이어야 합니다.');
  return { ok: true, index: Number(m[2]), error: null, message: '' };
}

/**
 * [일정] 시트 '대상자' 열 텍스트 -> 구성원 이름 목록. 공백·쉼표로 구분한다.
 * '전부'/'전체'/'all' 은 전체 대상으로 취급. members 에 없는 토큰은 unknown 에 담아 반환한다.
 * 반환 { targets:['전부']|[이름...], unknown:[인식 못한 토큰...] }
 */
function parseTargetNames_(text, members) {
  members = members || [];
  var tokens = String(text || '').split(/[,，\s]+/).map(function (t) { return t.trim(); }).filter(function (t) { return t; });
  var targets = [];
  var unknown = [];
  tokens.forEach(function (t) {
    if (t === '전부' || t === '전체' || t.toLowerCase() === 'all') {
      if (targets.indexOf('전부') < 0) targets.push('전부');
      return;
    }
    if (members.indexOf(t) >= 0) { if (targets.indexOf(t) < 0) targets.push(t); return; }
    if (unknown.indexOf(t) < 0) unknown.push(t);
  });
  if (targets.indexOf('전부') >= 0) targets = ['전부'];
  return { targets: targets, unknown: unknown };
}

function fail_(code, message) {
  return { ok: false, error: code, message: message };
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function ymd_(y, m, d) { return y + '-' + pad2_(m) + '-' + pad2_(d); }

function validDay_(y, m, d) {
  if (m < 1 || m > 12 || d < 1) return false;
  return d <= new Date(y, m, 0).getDate();
}

// 연도 생략 시: 올해, 단 오늘보다 이전이면 내년 (D-02 안2)
function resolveYear_(m, d, now) {
  var y = now.getFullYear();
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var cand = new Date(y, m - 1, d);
  if (cand < today) y += 1;
  return y;
}

function extractDate_(tokens, now) {
  var re = [
    { rx: /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/, y: 1, m: 2, d: 3 },
    { rx: /^(\d{1,2})[-\/.](\d{1,2})$/, m: 1, d: 2 },
    { rx: /^(\d{1,2})월(\d{1,2})일$/, m: 1, d: 2 }
  ];
  var rel = { '오늘': 0, '내일': 1, '모레': 2, '글피': 3 };
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (rel.hasOwnProperty(t)) {
      var dd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + rel[t]);
      return { found: true, date: ymd_(dd.getFullYear(), dd.getMonth() + 1, dd.getDate()), rest: without_(tokens, i, 1) };
    }
    // "6월" "6일" 두 토큰
    var mm = /^(\d{1,2})월$/.exec(t);
    if (mm && i + 1 < tokens.length && /^(\d{1,2})일$/.test(tokens[i + 1])) {
      var m1 = Number(mm[1]), d1 = Number(tokens[i + 1].replace('일', ''));
      var y1 = resolveYear_(m1, d1, now);
      if (!validDay_(y1, m1, d1)) return { error: 'INVALID_DATE', message: '유효하지 않은 날짜입니다: ' + t + ' ' + tokens[i + 1] };
      return { found: true, date: ymd_(y1, m1, d1), rest: without_(tokens, i, 2) };
    }
    for (var k = 0; k < re.length; k++) {
      var g = re[k].rx.exec(t);
      if (!g) continue;
      var m = Number(g[re[k].m]), d = Number(g[re[k].d]);
      var y = re[k].y ? Number(g[re[k].y]) : resolveYear_(m, d, now);
      if (!validDay_(y, m, d)) return { error: 'INVALID_DATE', message: '유효하지 않은 날짜입니다: ' + t };
      return { found: true, date: ymd_(y, m, d), rest: without_(tokens, i, 1) };
    }
  }
  return { found: false, rest: tokens };
}

function extractTime_(tokens, pmFrom) {
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    var ampm = null, consume = 1, h = null, mi = 0, body = t;

    if (t === '오전' || t === '오후') {
      if (i + 1 >= tokens.length) continue;
      ampm = t; body = tokens[i + 1]; consume = 2;
    } else if (/^(오전|오후)\d/.test(t)) {
      ampm = t.slice(0, 2); body = t.slice(2);
    }

    var g = /^(\d{1,2}):(\d{2})$/.exec(body);
    if (g) { h = Number(g[1]); mi = Number(g[2]); }
    else {
      g = /^(\d{1,2})시(?:(\d{1,2})분|반)?$/.exec(body);
      if (g) {
        h = Number(g[1]);
        if (g[2] !== undefined) mi = Number(g[2]);
        else if (/반$/.test(body)) mi = 30;
        else {
          var nx = tokens[i + consume];
          var g2 = nx && /^(\d{1,2})분$/.exec(nx);
          if (g2) { mi = Number(g2[1]); consume += 1; }
          else if (nx === '반') { mi = 30; consume += 1; }
        }
      }
    }
    if (h === null) continue;
    if (ampm === '오후' && h < 12) h += 12;
    if (ampm === '오전' && h === 12) h = 0;
    if (!ampm && pmFrom > 0 && h >= pmFrom && h < 12) h += 12;
    if (h < 0 || h > 23 || mi < 0 || mi > 59) return { error: 'INVALID_TIME', message: '유효하지 않은 시간입니다: ' + tokens.slice(i, i + consume).join(' ') };
    return { time: pad2_(h) + ':' + pad2_(mi), rest: without_(tokens, i, consume) };
  }
  return { time: '', rest: tokens };
}

function without_(arr, idx, n) {
  return arr.slice(0, idx).concat(arr.slice(idx + n));
}

/** 표시용: 2026-09-15(화) */
function formatDateKo(ymd) {
  var p = ymd.split('-').map(Number);
  var d = new Date(p[0], p[1] - 1, p[2]);
  return ymd + '(' + '일월화수목금토'.charAt(d.getDay()) + ')';
}

/** 표시용: 09월 15일 */
function formatMonthDay(ymd) {
  var p = ymd.split('-');
  return p[1] + '월 ' + p[2] + '일';
}

if (typeof module !== 'undefined') module.exports = {
  parseScheduleCommand: parseScheduleCommand,
  parseEditCommand: parseEditCommand,
  parseCancelCommand: parseCancelCommand,
  parseTargetNames_: parseTargetNames_,
  formatDateKo: formatDateKo,
  formatMonthDay: formatMonthDay
};
