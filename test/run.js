/**
 * test/run.js
 * 실행: npm test  (TZ=Asia/Seoul node test/run.js)
 * 목업 시트로 setup → 할 일 정렬 → 오늘↔위클리 → 못 한 일 넘기기 → 일정 추가 → 마이그레이션을 검증
 */
const assert = require('assert');
const { createEnv, Sheet } = require('./mock');

let passed = 0;
/** vm 안에서 만든 배열·객체도 내용으로 비교 */
const same = (actual, expected) => assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected);
function test(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); } catch (err) { console.log('  ✗ ' + name + '\n    ' + err.stack.split('\n').slice(0, 3).join('\n    ')); process.exitCode = 1; }
}

/** "오늘"을 고정한 환경 만들기 */
function envAt(y, m, d) {
  const env = createEnv();
  env.ctx.today_ = () => new Date(y, m - 1, d);
  env.run('setup()');
  return env;
}
const ws = (env) => env.ss.getSheetByName('위클리');
const ts = (env) => env.ss.getSheetByName('오늘');
const W = (env) => env.run('WEEKLY');
const T = (env) => env.run('TODAY');
/** 위클리 요일 d의 할 일 10줄 [체크, 시간, 내용] */
const weeklyTodos = (env, d) => ws(env).getRange(W(env).TODO_START, W(env).checkCol(d), 10, 3).getValues();
const setWeeklyTodos = (env, d, rows) => {
  const full = rows.concat(Array(10 - rows.length).fill([false, '', '']));
  ws(env).getRange(W(env).TODO_START, W(env).checkCol(d), 10, 3).setValues(full);
};
const edit = (env, sheet, r, c, nr, nc) => env.ctx.handleEdit({ range: sheet.getRange(r, c, nr || 1, nc || 1) });
const archiveRows = (env) => env.run('readArchive_()');

console.log('setup');
test('setup이 시트·트리거·기본값을 만든다', () => {
  const env = envAt(2026, 11, 11); // 수요일
  ['위클리', '오늘', '먼슬리', '지난 주 보기', '기록', '설정', '_일정'].forEach((n) => assert.ok(env.ss.getSheetByName(n), n));
  assert.strictEqual(env.triggers.length, 4);
  env.run('setup()');
  assert.strictEqual(env.triggers.length, 4, '다시 실행해도 트리거 중복 없음');
  assert.strictEqual(env.props.get('LAYOUT_VERSION'), '2');
  assert.strictEqual(env.ss.getSheetByName('설정').getRange(4, 2).getValue(), true, '자동 넘기기 기본 켜짐');
  assert.strictEqual(env.props.get('CARRY_DONE'), '2026-11-10', '첫날 어제 넘기기 안 함');
});

console.log('할 일 정렬');
test('미완료(시간순) → 시간 없음(입력순) → 체크된 일 → 빈 줄', () => {
  const env = envAt(2026, 11, 11);
  setWeeklyTodos(env, 0, [
    [false, '', '나'], [true, '', '끝난 일'], [false, '15:00', '다'], [false, '9:00', '가'], [false, '', '라'],
  ]);
  edit(env, ws(env), W(env).TODO_START, W(env).checkCol(0));
  const got = weeklyTodos(env, 0).slice(0, 6);
  assert.deepStrictEqual(got, [
    [false, '09:00', '가'], [false, '15:00', '다'], [false, '', '나'], [false, '', '라'], [true, '', '끝난 일'], [false, '', ''],
  ]);
});
test('체크하면 아래로 이동', () => {
  const env = envAt(2026, 11, 11);
  setWeeklyTodos(env, 1, [[false, '', 'A'], [false, '', 'B'], [false, '', 'C']]);
  ws(env).getRange(W(env).TODO_START, W(env).checkCol(1)).setValue(true); // A 체크
  edit(env, ws(env), W(env).TODO_START, W(env).checkCol(1));
  assert.deepStrictEqual(weeklyTodos(env, 1).slice(0, 3).map((r) => r[2]), ['B', 'C', 'A']);
});
test('오늘 시트에서 정렬 후 위클리 오늘 칸에 반영', () => {
  const env = envAt(2026, 11, 11); // 수요일 = d 2
  const t = T(env);
  ts(env).getRange(t.TODO_START, 1, 3, 3).setValues([[false, '', '책 읽기'], [false, '1400', '회의'], [true, '', '청소']]);
  edit(env, ts(env), t.TODO_START + 1, 2);
  assert.deepStrictEqual(ts(env).getRange(t.TODO_START, 1, 3, 3).getValues(),
    [[false, '14:00', '회의'], [false, '', '책 읽기'], [true, '', '청소']]);
  assert.deepStrictEqual(weeklyTodos(env, 2).slice(0, 3), [[false, '14:00', '회의'], [false, '', '책 읽기'], [true, '', '청소']]);
  const todo = archiveRows(env).filter((r) => r.cat === '할일');
  assert.strictEqual(todo.length, 3, '기록에 즉시 저장');
});
test('위클리 오늘 칸 수정이 오늘 시트에 반영', () => {
  const env = envAt(2026, 11, 11);
  setWeeklyTodos(env, 2, [[false, '', '운동']]);
  edit(env, ws(env), W(env).TODO_START, W(env).textCol(2));
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, 3).getValue(), '운동');
});

console.log('못 한 일 넘기기');
test('전날 미완료 → 오늘 빈 줄에 ↪, 전날은 미완 →, 기록은 FALSE+이월', () => {
  const env = envAt(2026, 11, 11);
  setWeeklyTodos(env, 2, [[false, '10:00', '보고서'], [true, '', '장보기'], [false, '', '↪ 운동']]); // 수
  setWeeklyTodos(env, 3, [[false, '', '목요일 일']]); // 목
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  const wed = weeklyTodos(env, 2);
  assert.deepStrictEqual(wed.slice(0, 3).map((r) => r[2]), ['미완 → 보고서', '장보기', '미완 → 운동']);
  const thu = weeklyTodos(env, 3).slice(0, 3);
  assert.deepStrictEqual(thu, [[false, '10:00', '↪ 보고서'], [false, '', '목요일 일'], [false, '', '↪ 운동']]);
  const rows = archiveRows(env).filter((r) => r.key === '2026-11-11' && r.cat === '할일');
  const report = rows.find((r) => r.item === '보고서');
  assert.strictEqual(report.done, false);
  assert.strictEqual(report.value, '이월');
  assert.ok(rows.find((r) => r.item === '운동' && r.value === '이월'), '↪ 표시는 떼고 저장');
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, 3).getValue(), '↪ 보고서', '오늘 시트도 갱신');
  // 같은 날 다시 실행해도 두 번 넘기지 않음
  env.run('dailyRollover_()');
  assert.strictEqual(weeklyTodos(env, 3).filter((r) => r[2] === '↪ 보고서').length, 1);
});
test('다음 날 10줄이 꽉 차면 오늘 시트 "넘어온 일"로', () => {
  const env = envAt(2026, 11, 11);
  setWeeklyTodos(env, 2, [[false, '', 'X'], [false, '', 'Y']]);
  setWeeklyTodos(env, 3, Array.from({ length: 10 }, (_, i) => [false, '', '할 일 ' + i]));
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  const t = T(env);
  assert.deepStrictEqual(ts(env).getRange(t.OVERFLOW_START, 1, 3, 1).getValues(), [['↪ X'], ['↪ Y'], ['']]);
});
test('일요일 미완료는 새 주 월요일로 (초기화 전에 처리)', () => {
  const env = envAt(2026, 11, 15); // 일요일
  setWeeklyTodos(env, 6, [[false, '20:00', '주간 정리'], [true, '', '빨래']]);
  env.ctx.today_ = () => new Date(2026, 10, 16); // 월요일
  env.run('midnightJob()');
  const mon = env.run('dateKey_(weeklyMonday_())');
  assert.strictEqual(mon, '2026-11-16', '새 주로 바뀜');
  assert.deepStrictEqual(weeklyTodos(env, 0)[0], [false, '20:00', '↪ 주간 정리']);
  assert.strictEqual(weeklyTodos(env, 6)[0][2], '', '지난주 칸은 비워짐');
  const sun = archiveRows(env).filter((r) => r.key === '2026-11-15' && r.cat === '할일');
  assert.ok(sun.find((r) => r.item === '주간 정리' && r.value === '이월' && r.done === false));
  assert.ok(sun.find((r) => r.item === '빨래' && r.done === true));
});
test('설정에서 끄면 넘기지 않음', () => {
  const env = envAt(2026, 11, 11);
  env.ss.getSheetByName('설정').getRange(4, 2).setValue(false);
  setWeeklyTodos(env, 2, [[false, '', '보고서']]);
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  assert.strictEqual(weeklyTodos(env, 2)[0][2], '보고서');
  assert.strictEqual(weeklyTodos(env, 3)[0][2], '');
});

console.log('새 일정 추가');
function addEvent(env, start, end, title) {
  const sh = ts(env), A = T(env).ADD;
  sh.getRange(A.date, 2).setValue(new Date(2026, 10, 20));
  sh.getRange(A.start, 2).setValue(start);
  sh.getRange(A.end, 2).setValue(end);
  sh.getRange(A.title, 2).setValue(title);
  sh.getRange(A.check, 1).setValue(true);
  edit(env, sh, A.check, 1);
  return sh.getRange(A.result, 2).getValue();
}
test('시작+끝 → 그 시간대 일정', () => {
  const env = envAt(2026, 11, 11);
  const res = addEvent(env, '14:00', '15:30', '회의');
  const ev = env.events[0];
  assert.ok(/추가됨/.test(res), res);
  assert.strictEqual(ev.allDay, false);
  assert.strictEqual(ev.title, '회의');
  assert.strictEqual(ev.start.getHours(), 14);
  assert.strictEqual(ev.end.getMinutes(), 30);
  assert.strictEqual(ts(env).getRange(T(env).ADD.title, 2).getValue(), '', '입력칸 비움');
});
test('시작만 → 종일 + 제목 앞에 시간', () => {
  const env = envAt(2026, 11, 11);
  addEvent(env, '15:00', '', '스터디');
  assert.strictEqual(env.events[0].allDay, true);
  assert.strictEqual(env.events[0].title, '15:00 스터디');
});
test('둘 다 없음 → 종일, 제목만', () => {
  const env = envAt(2026, 11, 11);
  addEvent(env, '', '', '생일');
  assert.strictEqual(env.events[0].allDay, true);
  assert.strictEqual(env.events[0].title, '생일');
});
test('끝이 시작보다 빠르면 "끝시간 확인" + 추가 안 함', () => {
  const env = envAt(2026, 11, 11);
  const res = addEvent(env, '16:00', '15:00', '회의');
  assert.strictEqual(res, '끝시간 확인');
  assert.strictEqual(env.events.length, 0);
  assert.strictEqual(ts(env).getRange(T(env).ADD.title, 2).getValue(), '회의', '입력값 유지');
});
test('같은 일정은 두 번 만들지 않음', () => {
  const env = envAt(2026, 11, 11);
  addEvent(env, '15:00', '', '스터디');
  env.cache.clear();
  const res = addEvent(env, '15:00', '', '스터디');
  assert.strictEqual(env.events.length, 1);
  assert.ok(/중복/.test(res), res);
});

console.log('마이그레이션');
test('예전 2열 위클리 입력값이 새 3열 위치로 옮겨짐', () => {
  const env = createEnv();
  env.ctx.today_ = () => new Date(2026, 11 - 1, 11);
  const old = env.ss.insertSheet('위클리');
  old.getRange(1, 4).setValue(new Date(2026, 10, 9));
  old.getRange(1, 8).setValue('운동 습관 만들기');
  old.getRange(18, 4).setValue(true); old.getRange(18, 5).setValue('화요일 일'); // 예전 화요일 체크·내용
  old.getRange(29, 2).setValue('4 🙂');
  old.getRange(14, 18 + 2).setValue(true); // 습관1 수요일
  env.run('setup()');
  const w = W(env), sh = ws(env);
  assert.strictEqual(env.run('dateKey_(weeklyMonday_())'), '2026-11-09');
  assert.strictEqual(sh.getRange(1, w.INFO_VALUE_COL).getValue(), '운동 습관 만들기');
  assert.deepStrictEqual(weeklyTodos(env, 1)[0], [true, '', '화요일 일']);
  assert.strictEqual(sh.getRange(w.MOOD, w.checkCol(0)).getValue(), '4 🙂');
  assert.strictEqual(sh.getRange(w.HABIT_START, w.HABIT_DAY_COL + 2).getValue(), true);
});


console.log('웹앱 API');
test('webBoot: 오늘 위치·7일·카테고리', () => {
  const env = envAt(2026, 11, 11);
  const b = env.ctx.webBoot();
  assert.strictEqual(b.todayKey, '2026-11-11');
  assert.strictEqual(b.todayIndex, 2);
  assert.strictEqual(b.week.days.length, 7);
  assert.strictEqual(b.week.monday, '2026-11-09');
  assert.strictEqual(b.categories[0].name, '공부');
  assert.strictEqual(b.week.days[2].todos.length, 10);
});
test('webSetTodos: 정렬 결과 반환 + 위클리·오늘 시트·기록 반영', () => {
  const env = envAt(2026, 11, 11);
  const rows = env.ctx.webSetTodos('2026-11-11', [[true, '', '끝'], [false, '1530', '회의'], [false, '', '독서']]);
  same(rows.slice(0, 4).map((r) => r.slice(0, 3)), [[false, '15:30', '회의'], [false, '', '독서'], [true, '', '끝'], [false, '', '']]);
  same(weeklyTodos(env, 2)[0], [false, '15:30', '회의']);
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, 3).getValue(), '회의', '오늘 시트 반영');
  assert.strictEqual(archiveRows(env).filter((r) => r.cat === '할일').length, 3);
  assert.throws(() => env.ctx.webSetTodos('2026-11-20', []), /이번 주/);
});
test('webSetHabit·webSetMood·webSetMust → 시트 반영', () => {
  const env = envAt(2026, 11, 11);
  const w = W(env), t = T(env);
  env.ctx.webSetHabit('2026-11-11', 1, true);
  env.ctx.webSetMood('2026-11-11', 4, '좋은 하루');
  env.ctx.webSetMust([[true, '보고서 제출'], [false, '운동 3번']]);
  assert.strictEqual(ws(env).getRange(w.HABIT_START + 1, w.HABIT_DAY_COL + 2).getValue(), true);
  assert.strictEqual(ws(env).getRange(w.MOOD, w.checkCol(2)).getValue(), '4 🙂');
  assert.strictEqual(ts(env).getRange(t.HABIT_START + 1, 1).getValue(), true);
  assert.strictEqual(ts(env).getRange(t.MOOD_MEMO, 2).getValue(), '좋은 하루');
  assert.strictEqual(ts(env).getRange(t.MUST_START, 2).getValue(), '보고서 제출');
  const mood = archiveRows(env).find((r) => r.cat === '기분');
  assert.strictEqual(mood.value, 4);
  assert.ok(archiveRows(env).find((r) => r.cat === '주간목표' && r.item === '보고서 제출' && r.done === true));
});
test('webGetMonth: 기분·습관 달성률·월간 목표', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-10', 0, true);
  env.ctx.webSetHabit('2026-11-11', 0, true);
  env.ctx.webSetMood('2026-11-10', 5, '');
  env.ctx.webSetMonthGoals(2026, 11, [[false, '책 2권'], [true, '저축']]);
  const M = env.ctx.webGetMonth(2026, 11);
  assert.strictEqual(M.days, 30);
  assert.strictEqual(M.offset, 6, '11/1은 일요일');
  assert.strictEqual(M.dates[9].mood, 5);
  assert.strictEqual(M.elapsed, 11);
  assert.strictEqual(M.habits[0].done, 2);
  same(M.goals.slice(0, 2), [[false, '책 2권'], [true, '저축']]);
});
test('타임라인: 저장·불러오기, 없는 카테고리는 버림, 다시 저장하면 덮어씀', () => {
  const env = envAt(2026, 11, 11);
  const slots = new Array(144).fill('');
  slots[54] = '공부'; slots[55] = '공부'; slots[60] = '없는것';
  env.ctx.webSetTimeline('2026-11-11', slots);
  let got = env.ctx.webGetTimeline('2026-11-11').slots;
  assert.strictEqual(got[54], '공부');
  assert.strictEqual(got[60], '');
  const rows = env.ss.getSheetByName('타임라인').getRange(2, 1, 2, 3).getDisplayValues();
  same(rows.map((r) => r.slice(1)), [['09:00', '공부'], ['09:10', '공부']]);
  slots[54] = ''; env.ctx.webSetTimeline('2026-11-11', slots);
  got = env.ctx.webGetTimeline('2026-11-11').slots;
  assert.strictEqual(got.filter(Boolean).length, 1);
});
test('webAddEvent: 같은 규칙 + 그날 일정에 바로 보임', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webAddEvent({ date: '2026-11-12', start: '15:00', end: '', title: '스터디' });
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events[0].title, '15:00 스터디');
  same(res.week.days[3].schedule.map((x) => x.text), ['15:00 스터디']);
  const bad = env.ctx.webAddEvent({ date: '2026-11-12', start: '16:00', end: '15:00', title: '회의' });
  assert.strictEqual(bad.message, '끝시간 확인');
});
test('예전 문서 속성 값은 스크립트 속성으로 옮겨짐 (넘기기 중복 방지 유지)', () => {
  const env = createEnv();
  env.docProps.set('CARRY_DONE', '2026-11-10');
  env.docProps.set('LAYOUT_VERSION', '2');
  assert.strictEqual(env.run("props_().getProperty('CARRY_DONE')"), '2026-11-10');
  assert.strictEqual(env.props.get('LAYOUT_VERSION'), '2');
});


console.log('먼슬리 일정·시간표·습관 편집');
const evRows = (env) => env.run('readEvents_()').map((r) => [env.run('dateKey_')(r[0]), r[1], r[2], r[4] ? 'id' : '', r[5]]);
const NEW = null;
const itemRef = (items, text) => (items.find((it) => it.text === text) || {}).ref;
test('새 일정: 항상 구글 캘린더에 추가 + 먼슬리 메모(연동), 일정 칸에 한 번만, 할 일에는 안 들어감', () => {
  const env = envAt(2026, 11, 11);
  let res = env.ctx.webSaveItem('2026-11-12', NEW, { title: '스터디', start: '15:00', end: '16:30' }, 2026, 11);
  assert.ok(res.ok, res.message);
  assert.ok(/구글 캘린더/.test(res.message));
  assert.strictEqual(env.events.length, 1, '토글 없이 항상 캘린더에 생성');
  assert.strictEqual(env.events[0].allDay, false);
  res = env.ctx.webSaveItem('2026-11-12', NEW, { title: '택배 받기', start: '', end: '' }, 2026, 11);
  same(res.week.days[3].schedule.map((x) => [x.kind, x.text]), [['cal', '택배 받기'], ['cal', '15:00-16:30 스터디']]);
  same(res.month.dates[11].items.map((x) => x.text), ['택배 받기', '15:00-16:30 스터디']);
  assert.ok(res.week.days[3].todos.every((r) => !r[2]), '할 일 10줄은 그대로');
  const M = env.run('MONTHLY'), ms = env.ss.getSheetByName('먼슬리');
  assert.strictEqual(ms.getRange(M.memoRow(2), M.dayCol(3)).getValue(), '택배 받기\n15:00-16:30 스터디', '먼슬리 메모 칸');
  res = env.ctx.webSaveItem('2026-11-12', NEW, { title: '회의', start: '16:00', end: '15:00' }, 2026, 11);
  assert.strictEqual(res.message, '끝시간 확인');
  assert.strictEqual(env.events.length, 2);
});
test('캘린더 일정 수정: ID로 찾아 수정 + 연동 메모도 수정 (오늘·먼슬리 어디서든)', () => {
  const env = envAt(2026, 11, 11);
  let res = env.ctx.webSaveItem('2026-11-13', NEW, { title: '발표', start: '14:00', end: '' }, 2026, 11);
  const ev = env.events[0];
  let ref = itemRef(res.week.days[4].schedule, '14:00 발표'); // 위클리(요일 상세)에서 누른 것
  res = env.ctx.webSaveItem('2026-11-13', ref, { title: '발표 리허설', start: '15:00', end: '16:00' }, 2026, 11);
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events.length, 1, '새로 만들지 않고 그 일정을 수정');
  assert.strictEqual(ev.title, '발표 리허설');
  assert.strictEqual(ev.start.getHours(), 15);
  same(res.month.dates[12].items.map((x) => x.text), ['15:00-16:00 발표 리허설']);
  const memo = env.run('readEvents_()').find((r) => r[1] === '메모');
  assert.strictEqual(memo[2], '15:00-16:00 발표 리허설', '연동 메모도 같이 수정');
  // 먼슬리에서 다시 수정 → 종일
  ref = itemRef(res.month.dates[12].items, '15:00-16:00 발표 리허설');
  res = env.ctx.webSaveItem('2026-11-13', ref, { title: '발표', start: '', end: '' }, 2026, 11);
  assert.strictEqual(ev.allDay, true);
  same(res.week.days[4].schedule.map((x) => x.text), ['발표']);
});
test('캘린더 일정 삭제: 캘린더·연동 메모 모두 삭제, 목록·달력 바로 갱신', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSaveItem('2026-11-13', NEW, { title: '발표', start: '14:00', end: '15:00' }, 2026, 11);
  let res = env.ctx.webSaveItem('2026-11-13', NEW, { title: '친구 생일', start: '', end: '' }, 2026, 11);
  res = env.ctx.webDeleteItem('2026-11-13', itemRef(res.month.dates[12].items, '14:00-15:00 발표'), 2026, 11);
  assert.ok(res.ok, res.message);
  same(env.events.map((e) => e.title), ['친구 생일']);
  same(res.month.dates[12].items.map((x) => x.text), ['친구 생일']);
  same(res.week.days[4].schedule.map((x) => x.text), ['친구 생일']);
  assert.strictEqual(env.run('readEvents_()').filter((r) => r[1] === '메모').length, 1, '연동 메모도 삭제');
});
test('예전 "새 일정 추가"로 만든 일정·구글에서 불러온 일정도 수정·삭제 (ID 없으면 날짜+제목)', () => {
  const env = envAt(2026, 11, 11);
  // 시트의 예전 입력칸으로 만든 일정
  env.ctx.createCalendarEvent_(new Date(2026, 10, 12), '10:00', '11:00', '치과');
  // 구글 캘린더에서 직접 만든 일정 (앱이 모름)
  env.ctx.CalendarApp.getDefaultCalendar().createAllDayEvent('가족 모임', new Date(2026, 10, 12));
  env.ctx.syncCalendarHourly();
  let b = env.ctx.webBoot();
  same(b.week.days[3].schedule.map((x) => x.text), ['가족 모임', '10:00-11:00 치과']);
  let res = env.ctx.webSaveItem('2026-11-12', itemRef(b.week.days[3].schedule, '10:00-11:00 치과'), { title: '치과 검진', start: '10:30', end: '11:30' }, null, null);
  assert.ok(res.ok, res.message);
  assert.ok(env.events.find((e) => e.title === '치과 검진' && e.start.getMinutes() === 30));
  // ID 없이 저장된 예전 줄 흉내 → 날짜+제목으로 찾아 삭제
  const sh = env.ss.getSheetByName('_일정');
  const rows = env.run('readEvents_()');
  const i = rows.findIndex((r) => r[2] === '가족 모임');
  sh.getRange(i + 2, 5).setValue('');
  env.ctx.bumpCache_();
  b = env.ctx.webBoot();
  res = env.ctx.webDeleteItem('2026-11-12', itemRef(b.week.days[3].schedule, '가족 모임'), null, null);
  assert.ok(res.ok, res.message);
  assert.ok(!env.events.find((e) => e.title === '가족 모임'), '구글 캘린더에서도 삭제');
});
test('삭제 실패 이유 표시: 캘린더에서 이미 지워진 일정', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.CalendarApp.getDefaultCalendar().createAllDayEvent('없어질 일정', new Date(2026, 10, 12));
  env.ctx.syncCalendarHourly();
  const b = env.ctx.webBoot();
  env.events.length = 0; // 구글 캘린더에서 먼저 삭제
  const res = env.ctx.webDeleteItem('2026-11-12', itemRef(b.week.days[3].schedule, '없어질 일정'), null, null);
  assert.strictEqual(res.ok, false);
  assert.ok(/찾지 못했어요/.test(res.message), res.message);
  same(res.week.days[3].schedule, [], '목록은 최신으로 갱신');
});
test('다른 기기에서 바뀐 일정은 막음 + 시간표는 수정 불가', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webDeleteItem('2026-11-12', { kind: 'cal', index: 0, text: '없는 일정' }, null, null);
  assert.ok(/다른 곳에서/.test(res.message), res.message);
  env.ctx.webSetTimetable({ start: '', end: '', rows: [{ day: 3, start: '09:00', end: '10:00', name: '통계학' }] });
  const r2 = env.ctx.webDeleteItem('2026-11-12', { kind: 'tt', index: 0, text: '09:00-10:00 통계학' }, null, null);
  assert.ok(/바꿀 수 없|지울 수 없/.test(r2.message), r2.message);
});
test('시트 메모 칸을 직접 고쳐도 줄마다 일정 하나, 연동 정보 유지', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSaveItem('2026-11-13', NEW, { title: '발표', start: '14:00', end: '' }, 2026, 11);
  const M = env.run('MONTHLY'), ms = env.ss.getSheetByName('먼슬리');
  ms.getRange(M.memoRow(2), M.dayCol(4)).setValue('14:00 발표\n저녁 약속');
  edit(env, ms, M.memoRow(2), M.dayCol(4));
  const rows = evRows(env).filter((r) => r[0] === '2026-11-13' && r[1] === '메모');
  same(rows, [['2026-11-13', '메모', '저녁 약속', '', ''], ['2026-11-13', '메모', '14:00 발표', 'id', '연동']]);
});
test('고정 시간표: 학기 기간에만 매주 일정 칸에 표시', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webSetTimetable({ start: '2026-09-01', end: '2026-11-11', rows: [
    { day: 1, start: '10:00', end: '11:30', name: '데이터마이닝' }, { day: 3, start: '9:00', end: '', name: '통계학' },
  ] });
  same(res.timetable.rows.map((r) => r.name), ['데이터마이닝', '통계학']);
  same(res.week.days[1].schedule.map((x) => [x.kind, x.text]), [['tt', '10:00-11:30 데이터마이닝']]);
  same(res.week.days[3].schedule, [], '학기 종료(11/11) 뒤인 목요일은 표시 안 함');
  assert.ok(res.week.days[1].todos.every((r) => !r[2]), '할 일 줄 차지 안 함');
  const set = env.ss.getSheetByName('설정'), S = env.run('SETTINGS');
  same(set.getRange(S.TT_START, 1, 1, 4).getValues()[0], ['화', '10:00', '11:30', '데이터마이닝']);
  assert.throws(() => env.ctx.webSetTimetable({ rows: [{ day: 0, start: '10:00', end: '09:00', name: 'x' }] }), /끝시간/);
});
test('습관 편집: 순서·이름·삭제 → 설정 B6~, 위클리 체크도 따라 이동, 지난 기록 유지', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-10', 0, true); // 어제 '물 8잔 마시기'
  env.ctx.webSetHabit('2026-11-11', 2, true); // 오늘 '독서 20분'
  const b = env.ctx.webSetHabits([
    { name: '독서 30분', from: 2 }, { name: '물 8잔 마시기', from: 0 }, { name: '명상', from: null },
  ]);
  same(b.habits.slice(0, 4), ['독서 30분', '물 8잔 마시기', '명상', '']);
  const S = env.run('SETTINGS');
  same(env.ss.getSheetByName('설정').getRange(S.HABIT_START, 2, 3, 1).getValues(), [['독서 30분'], ['물 8잔 마시기'], ['명상']]);
  assert.strictEqual(b.week.days[2].habits[0], true, '독서 체크가 새 위치로 이동');
  assert.strictEqual(b.week.days[1].habits[1], true, '물 체크도 이동');
  const arc = archiveRows(env).filter((r) => r.cat === '습관');
  assert.ok(arc.find((r) => r.key === '2026-11-10' && r.item === '운동 30분'), '지운 습관의 어제 기록 유지');
  assert.ok(arc.find((r) => r.key === '2026-11-11' && r.item === '독서 30분' && r.done === true), '오늘은 새 이름으로');
  assert.throws(() => env.ctx.webSetHabits(Array.from({ length: 11 }, (_, i) => ({ name: 'h' + i, from: null }))), /최대/);
});



console.log('시간 범위·공휴일·캐시');
test('할 일 시간 범위: "11-13", "11~13", "1100-1300" 자동 변환 + 시작 시간순 정렬', () => {
  const env = envAt(2026, 11, 11);
  const rows = env.ctx.webSetTodos('2026-11-11', [
    [false, '', '점심 약속 11~13'], [false, '', '1400-1530 회의'], [false, '', '9-10 운동'], [false, '15:00', '전화'],
    [false, '', '3-4쪽 읽기'], [false, '10:00-09:00', '이상한 시간'],
  ]);
  same(rows.slice(0, 6).map((r) => r.slice(0, 3)), [
    [false, '09:00-10:00', '운동'], [false, '11:00-13:00', '점심 약속'], [false, '14:00-15:30', '회의'],
    [false, '15:00', '전화'], [false, '', '3-4쪽 읽기'], [false, '10:00-09:00', '이상한 시간'],
  ]);
  assert.strictEqual(env.run("normRange_('11-13')"), '11:00-13:00');
  assert.strictEqual(env.run("normRange_('13-11')"), '');
});
test('공휴일: 구독 안 돼 있으면 구독, 일정 목록에 안 섞임, 날짜에 이름 표시', () => {
  const env = envAt(2026, 11, 11);
  assert.strictEqual(env.holidays.subscribed, true, 'setup 동기화 때 구독');
  const b = env.ctx.webBoot();
  assert.strictEqual(b.week.days[2].holiday, '테스트 기념일');
  same(b.week.days[2].schedule, [], '일정 목록에는 없음');
  const M = env.ctx.webGetMonth(2026, 12);
  assert.strictEqual(M.dates[24].holiday, '기독탄신일');
  assert.ok(M.dates[24].items.every((it) => it.kind !== 'holiday'));
  assert.ok(env.run('scheduleLineFormula_("B$13", 1, "")').includes('<>"공휴일"'), '시트 일정 칸에서도 제외');
  const r = env.ctx.webDeleteItem('2026-11-11', { kind: 'cal', index: 0, text: '테스트 기념일' }, null, null);
  assert.strictEqual(r.ok, false, '공휴일은 수정·삭제 불가');
});
test('서버 캐시: 같은 데이터는 캐시에서, 쓰기·시트 편집 뒤에는 새로 읽음', () => {
  const env = envAt(2026, 11, 11);
  const a = env.ctx.webBoot();
  ts(env).getRange(T(env).TODO_START, 3).setValue('캐시 확인'); // 트리거 없이 몰래 바꿈
  same(env.ctx.webBoot().week, JSON.parse(JSON.stringify(a.week)), '캐시된 값');
  edit(env, ts(env), T(env).TODO_START, 3); // 시트 편집(설치형 onEdit) → 캐시 무효화
  assert.strictEqual(env.ctx.webBoot().week.days[2].todos[0][2], '캐시 확인');
});

console.log('숨김 메모');
test('할 일 메모: 같은 줄 숨김 칸에 저장, 정렬 때 같이 이동, 오늘 시트에도 반영, 기록에는 없음', () => {
  const env = envAt(2026, 11, 11);
  const rows = env.ctx.webSetTodos('2026-11-11', [[true, '', '끝난 일', '끝난 일 메모'], [false, '', '독서', ''], [false, '9-10', '운동', '스쿼트 3세트\n런지']]);
  same(rows.slice(0, 3), [[false, '09:00-10:00', '운동', '스쿼트 3세트\n런지'], [false, '', '독서', ''], [true, '', '끝난 일', '끝난 일 메모']]);
  const W2 = W(env);
  same(ws(env).getRange(W2.TODO_START, W2.memoCol(2), 3, 1).getValues(), [['스쿼트 3세트\n런지'], [''], ['끝난 일 메모']]);
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, T(env).MEMO_COL).getValue(), '스쿼트 3세트\n런지', '오늘 시트 D열');
  // 시트에서 체크해 순서가 바뀌어도 메모가 따라감
  ws(env).getRange(W2.TODO_START, W2.checkCol(2)).setValue(true);
  edit(env, ws(env), W2.TODO_START, W2.checkCol(2));
  same(ws(env).getRange(W2.TODO_START, W2.textCol(2), 3, 1).getValues(), [['독서'], ['운동'], ['끝난 일']]);
  same(ws(env).getRange(W2.TODO_START, W2.memoCol(2), 3, 1).getValues(), [[''], ['스쿼트 3세트\n런지'], ['끝난 일 메모']]);
  assert.ok(archiveRows(env).every((r) => !String(r.item).includes('스쿼트') && !String(r.value).includes('스쿼트')), '기록 시트엔 메모 없음');
  assert.strictEqual(env.ctx.webBoot().week.days[2].todos[1][3], '스쿼트 3세트\n런지');
});
test('할 일이 다음 날로 넘어갈 때 메모도 같이', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTodos('2026-11-11', [[false, '', '보고서', '3쪽까지 썼음'], [true, '', '장보기', '우유']]);
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  const thu = env.ctx.webBoot().week.days[3].todos;
  same(thu[0], [false, '', '↪ 보고서', '3쪽까지 썼음']);
});
test('일정 메모: 먼슬리 메모와 함께 저장 + 구글 캘린더 설명, 캘린더 설명은 메모로 보임', () => {
  const env = envAt(2026, 11, 11);
  let res = env.ctx.webSaveItem('2026-11-12', null, { title: '스터디', start: '15:00', end: '16:00', memo: '3장 발표 준비\n노트북 챙기기' }, 2026, 11);
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events[0].description, '3장 발표 준비\n노트북 챙기기', '캘린더 설명');
  const item = res.week.days[3].schedule[0];
  assert.strictEqual(item.memo, '3장 발표 준비\n노트북 챙기기');
  assert.strictEqual(item.text, '15:00-16:00 스터디', '목록 글자에는 메모 없음');
  const memoRow = env.run('readEvents_()').find((r) => r[1] === '메모');
  assert.strictEqual(memoRow[6], '3장 발표 준비\n노트북 챙기기', '먼슬리 메모 줄에 함께 저장');
  // 수정하면 캘린더 설명도 수정
  res = env.ctx.webSaveItem('2026-11-12', item.ref, { title: '스터디', start: '15:00', end: '16:00', memo: '장소: 도서관' }, 2026, 11);
  assert.strictEqual(env.events[0].description, '장소: 도서관');
  assert.strictEqual(res.month.dates[11].items[0].memo, '장소: 도서관');
  // 구글 캘린더에서 직접 만든 일정의 설명 → 메모
  env.ctx.CalendarApp.getDefaultCalendar().createAllDayEvent('동아리', new Date(2026, 10, 13), { description: '회비 내기' });
  env.ctx.syncCalendarHourly();
  assert.strictEqual(env.ctx.webBoot().week.days[4].schedule[0].memo, '회비 내기');
});

console.log(`\n${passed}개 통과` + (process.exitCode ? ' (실패 있음)' : ''));
