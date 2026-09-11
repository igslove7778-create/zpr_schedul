// 실행: node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const { parseScheduleCommand, formatDateKo } = require('../apps-script/Parser.gs');

const members = ['김유선', '성도형', '김명재', '조인희'];
const now = new Date(2026, 8, 11, 10, 0); // 2026-09-11 (금)
const P = (t, o) => parseScheduleCommand(t, Object.assign({ now, members }, o));

// 명세 5.4 인식 예시
test('1 일정 6월6일 6:30 전체회의 (D-01 기본: 24시간제)', () => {
  const r = P('일정 6월6일 6:30 전체회의');
  assert.deepStrictEqual([r.ok, r.date, r.time, r.title, r.targets], [true, '2027-06-06', '06:30', '전체회의', []]);
});
test('2 @전부', () => {
  const r = P('일정 6/6 18:30 전체회의 @전부');
  assert.deepStrictEqual([r.date, r.time, r.title, r.targets], ['2027-06-06', '18:30', '전체회의', ['전부']]);
});
test('3 오후 6시30분', () => {
  const r = P('일정 6-6 오후 6시30분 전체회의');
  assert.deepStrictEqual([r.time, r.title], ['18:30', '전체회의']);
});
test('4 종일 + 특정 대상자', () => {
  const r = P('일정 9월 15일 방산IR 제출 @김유선 @성도형');
  assert.deepStrictEqual([r.date, r.time, r.allDay, r.title, r.targets], ['2026-09-15', '', true, '방산IR 제출', ['김유선', '성도형']]);
});
test('5 일정명 숫자 유지', () => {
  const r = P('일정 6/6 14:00 2차 기술검토');
  assert.deepStrictEqual([r.time, r.title], ['14:00', '2차 기술검토']);
});
test('6 날짜 없음', () => assert.strictEqual(P('일정 18:30 전체회의').error, 'NO_DATE'));
test('7 일정명 없음', () => assert.strictEqual(P('일정 6/6 18:30').error, 'NO_TITLE'));
test('8 유효하지 않은 날짜', () => assert.strictEqual(P('일정 13/40 회의').error, 'INVALID_DATE'));
test('9 유효하지 않은 시간', () => assert.strictEqual(P('일정 6/6 25:10 회의').error, 'INVALID_TIME'));

// 추가 규칙
test('명령어 아님', () => assert.strictEqual(P('내일 회의').error, 'NOT_COMMAND'));
test('알 수 없는 대상자', () => assert.strictEqual(P('일정 6/6 회의 @홍길동').error, 'UNKNOWN_TARGET'));
test('D-02: 지난 날짜는 내년', () => assert.strictEqual(P('일정 1/5 회의').date, '2027-01-05'));
test('D-02: 오늘은 올해', () => assert.strictEqual(P('일정 9/11 회의').date, '2026-09-11'));
test('D-01 안2: pmFrom=1 이면 6:30 -> 18:30, 12:00은 그대로', () => {
  assert.strictEqual(P('일정 6/6 6:30 회의', { pmFrom: 1 }).time, '18:30');
  assert.strictEqual(P('일정 6/6 12:00 회의', { pmFrom: 1 }).time, '12:00');
});
test('오전 9시 / 오전 12시 / 오후 12시', () => {
  assert.strictEqual(P('일정 6/6 오전 9시 회의').time, '09:00');
  assert.strictEqual(P('일정 6/6 오전 12시 회의').time, '00:00');
  assert.strictEqual(P('일정 6/6 오후 12시 회의').time, '12:00');
});
test('6시 반 / 6시반 / 오후 6시 30분 / 오후6시', () => {
  assert.strictEqual(P('일정 6/6 6시 반 회의').time, '06:30');
  assert.strictEqual(P('일정 6/6 6시반 회의').time, '06:30');
  assert.strictEqual(P('일정 6/6 오후 6시 30분 회의').time, '18:30');
  assert.strictEqual(P('일정 6/6 오후6시 회의').time, '18:00');
});
test('오늘/내일/모레, 6.6, 2026-06-06', () => {
  assert.strictEqual(P('일정 내일 회의').date, '2026-09-12');
  assert.strictEqual(P('일정 6.6 회의').date, '2027-06-06');
  assert.strictEqual(P('일정 2026-06-06 회의').date, '2026-06-06');
});
test('연속 공백 정리, 전체회의라는 단어는 대상자 아님', () => {
  const r = P('  일정   6/6   전체회의  ');
  assert.deepStrictEqual([r.title, r.targets], ['전체회의', []]);
});
test('첫 번째 날짜·시간 토큰만 사용', () => {
  const r = P('일정 6/6 10:00 7/7 11:00 검토');
  assert.deepStrictEqual([r.date, r.time, r.title], ['2027-06-06', '10:00', '7/7 11:00 검토']);
});
test('"오후 회의"처럼 시간이 아닌 오후는 일정명 유지', () => {
  const r = P('일정 6/6 오후 회의');
  assert.deepStrictEqual([r.time, r.title], ['', '오후 회의']);
});
test('formatDateKo', () => assert.strictEqual(formatDateKo('2026-09-15'), '2026-09-15(화)'));
