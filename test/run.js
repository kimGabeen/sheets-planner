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
  same(rows.slice(0, 4), [[false, '15:30', '회의'], [false, '', '독서'], [true, '', '끝'], [false, '', '']]);
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

console.log(`\n${passed}개 통과` + (process.exitCode ? ' (실패 있음)' : ''));
