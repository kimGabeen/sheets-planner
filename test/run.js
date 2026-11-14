/**
 * test/run.js
 * 실행: npm test  (TZ=Asia/Seoul node test/run.js)
 * 목업 시트·캘린더로 날짜별 저장소, 시트 ↔ 웹앱 양방향, 날짜 이동, 넘기기, 일정·중복, 마이그레이션을 검증
 */
const assert = require('assert');
const { createEnv } = require('./mock');

let passed = 0;
/** vm 안에서 만든 배열·객체도 내용으로 비교 */
const same = (actual, expected, msg) => {
  const a = JSON.parse(JSON.stringify(actual)), b = JSON.parse(JSON.stringify(expected));
  if (msg) assert.deepStrictEqual(a, b, msg); else assert.deepStrictEqual(a, b);
};
function test(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); } catch (err) { console.log('  ✗ ' + name + '\n    ' + String(err.stack).split('\n').slice(0, 40).join('\n    ')); process.exitCode = 1; }
}

/** "오늘"을 고정한 환경 (2026-11-11 수요일 기본) */
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
const edit = (env, sheet, r, c, nr, nc) => env.ctx.handleEdit({ range: sheet.getRange(r, c, nr || 1, nc || 1) });
// 기록 시트는 1시간마다(·밤·새 주) 한 번에 저장되므로, 확인 전에 1시간 트리거를 돌림
const archiveRows = (env) => { env.ctx.syncCalendarHourly(); return env.run('readArchive_()'); };
const boot = (env) => env.ctx.webBoot();
const dayOf = (week, key) => week.days.find((d) => d.key === key);
const todoTexts = (day) => day.todos.map((r) => r[2]);
/** 위클리 시트 요일 d의 보이는 할 일 [체크, 시간, 내용] */
const sheetTodos = (env, d) => ws(env).getRange(W(env).TODO_START, W(env).checkCol(d), 10, 3).getValues();
const todos = (list) => list.map((t) => (Array.isArray(t) ? t : [false, '', t, '', '']));

console.log('setup·저장소');
test('setup이 시트·저장소·트리거를 만든다 (다시 실행해도 트리거 중복 없음)', () => {
  const env = envAt(2026, 11, 11);
  ['위클리', '오늘', '먼슬리', '지난 주 보기', '기록', '할일저장소', '하루저장소', '주간저장소', '설정', '_일정'].forEach((n) => assert.ok(env.ss.getSheetByName(n), n));
  assert.strictEqual(env.triggers.length, 4);
  env.run('setup()');
  assert.strictEqual(env.triggers.length, 4);
  assert.strictEqual(env.props.get('STORE_VERSION'), '1', '처음 설치는 옮길 것 없이 저장소 사용');
  assert.strictEqual(env.copies.length, 0, '처음 설치는 백업 안 만듦');
  assert.strictEqual(env.ss.getSheetByName('설정').getRange(4, 2).getValue(), true, '자동 넘기기 기본 켜짐');
});

console.log('시트 ↔ 웹앱 양방향');
test('시트에서 할 일 추가·체크·메모 수정 → 저장소·웹앱·오늘 시트에 반영', () => {
  const env = envAt(2026, 11, 11); // 수요일 = d 2
  const w = W(env), sh = ws(env);
  sh.getRange(w.TODO_START, w.checkCol(2), 2, 3).setValues([[false, '', '보고서 쓰기'], [false, '1400', '회의']]);
  edit(env, sh, w.TODO_START, w.checkCol(2), 2, 3);
  let day = dayOf(boot(env).week, '2026-11-11');
  same(day.todos.map((r) => r.slice(0, 3)), [[false, '14:00', '회의'], [false, '', '보고서 쓰기']], '시간 있는 일이 위로');
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, 3).getValue(), '회의', '오늘 시트도 같은 순서');
  sh.getRange(w.TODO_START, w.checkCol(2)).setValue(true); // 회의 체크
  edit(env, sh, w.TODO_START, w.checkCol(2));
  sh.getRange(w.TODO_START + 1, w.memoCol(2)).setValue('회의록 정리하기'); // 정렬 후 2번째 줄 = 회의
  edit(env, sh, w.TODO_START + 1, w.memoCol(2));
  day = dayOf(boot(env).week, '2026-11-11');
  same(day.todos, [[false, '', '보고서 쓰기', '', ''], [true, '14:00', '회의', '회의록 정리하기', '']]);
  assert.strictEqual(ts(env).getRange(T(env).TODO_START + 1, 1).getValue(), true);
  assert.strictEqual(ts(env).getRange(T(env).TODO_START + 1, T(env).MEMO_COL).getValue(), '회의록 정리하기');
  // 오늘 시트에서 고쳐도 위클리·웹앱에
  ts(env).getRange(T(env).TODO_START, 1).setValue(true);
  edit(env, ts(env), T(env).TODO_START, 1);
  assert.ok(dayOf(boot(env).week, '2026-11-11').todos.every((r) => r[0]), '둘 다 체크됨');
  assert.strictEqual(sh.getRange(w.RATE, w.checkCol(2)).getValue(), '●  100%');
});
test('웹앱에서 할 일 추가·체크·메모 → 위클리·오늘 시트·기록에 바로 반영', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webSetTodos('2026-11-11', [[false, '9-10', '운동', '스쿼트 3세트', ''], [true, '', '장보기', '우유', ''], [false, '', '독서', '', '']]);
  same(todoTexts(res.day), ['운동', '독서', '장보기']);
  same(sheetTodos(env, 2).slice(0, 3), [[false, '09:00-10:00', '운동'], [false, '', '독서'], [true, '', '장보기']]);
  assert.strictEqual(ws(env).getRange(W(env).TODO_START, W(env).memoCol(2)).getValue(), '스쿼트 3세트');
  assert.strictEqual(ts(env).getRange(T(env).TODO_START + 2, T(env).MEMO_COL).getValue(), '우유');
  assert.strictEqual(archiveRows(env).filter((r) => r.cat === '할일').length, 3, '기록 시트에 저장');
  assert.ok(archiveRows(env).every((r) => !String(r.item).includes('스쿼트')), '기록에 메모 없음');
});
test('시트에서 할 일 내용을 지우면 그 메모도 삭제, 메모는 정렬·이동 때 항상 따라감', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTodos('2026-11-11', [[false, '', 'A', 'A 메모', ''], [false, '', 'B', 'B 메모', '']]);
  const w = W(env), sh = ws(env);
  sh.getRange(w.TODO_START, w.textCol(2)).setValue(''); // A 지움
  edit(env, sh, w.TODO_START, w.textCol(2));
  same(dayOf(boot(env).week, '2026-11-11').todos, [[false, '', 'B', 'B 메모', '']]);
  same(sh.getRange(w.TODO_START, w.memoCol(2), 2, 1).getValues(), [['B 메모'], ['']], '남은 메모 없음');
});
test('할 일이 늘면 시트 줄이 자동으로 늘고(빈 줄 2개 유지) 아래 칸은 같이 밀림, "+N개 더" 없음', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-11', '운동 30분', true);
  env.ctx.webSetMood('2026-11-11', 4, '좋음');
  const rows = Array.from({ length: 12 }, (_, i) => [i < 3, '', '할 일 ' + (i + 1), i === 11 ? '마지막 메모' : '', '']);
  env.ctx.webSetTodos('2026-11-11', rows);
  const w = W(env), sh = ws(env);
  assert.strictEqual(w.TODO_N, 14, '12개 + 빈 줄 2개');
  const block = sh.getRange(w.TODO_START, w.checkCol(2), w.TODO_N, 3).getValues();
  assert.strictEqual(block.filter((r) => r[2]).length, 12, '12개 모두 보임');
  same(block.slice(12), [[false, '', ''], [false, '', '']], '맨 아래 빈 줄 2개');
  assert.strictEqual(sh.getRange(w.TODO_START + 11, w.memoCol(2)).getValue(), '', '메모도 같은 줄 (정렬 후 체크된 일은 아래)');
  assert.strictEqual(sh.getRange(w.COUNT, w.checkCol(2)).getValue(), '3 완료 · 9 남음', '"+N개 더" 없음');
  assert.strictEqual(sh.getRange(w.MOOD, w.checkCol(2)).getValue(), '4 🙂', '기분 칸이 아래로 밀린 자리에 그대로');
  assert.strictEqual(sh.getRange(w.MOOD_MEMO, w.checkCol(2)).getValue(), '좋음');
  assert.strictEqual(sh.getRange(w.HABIT_START + 1, w.HABIT_DAY_COL + 2).getValue(), true, '습관은 위쪽이라 그대로');
  // 오늘 시트도
  const t = T(env), today = ts(env);
  assert.strictEqual(t.TODO_N, 14);
  assert.strictEqual(today.getRange(t.TODO_START, 1, t.TODO_N, 3).getValues().filter((r) => r[2]).length, 12);
  assert.strictEqual(today.getRange(t.TODO_HEADER, 1).getValue(), '오늘 할 일   3 / 12');
  assert.strictEqual(today.getRange(t.HABIT_START + 1, 1).getValue(), true, '습관 칸이 밀린 자리에');
  assert.strictEqual(today.getRange(t.MOOD, 2).getValue(), '4 🙂', '기분 칸이 밀린 자리에');
  assert.strictEqual(today.getRange(t.ADD.check, 1).getValue() === true, false, '일정 추가 칸도 같이 밀림');
});
test('시트에서 새로 생긴 빈 줄에 적은 할 일 → 저장소·웹앱 반영, 줄이 또 늘어남', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTodos('2026-11-11', Array.from({ length: 10 }, (_, i) => [false, '', 'A' + i, '', '']));
  const w = W(env), sh = ws(env);
  assert.strictEqual(w.TODO_N, 12);
  sh.getRange(w.TODO_START + 10, w.checkCol(2), 1, 3).setValues([[false, '9-10', '11번째 할 일']]); // 11번째 줄 (새 빈 줄)
  sh.getRange(w.TODO_START + 10, w.memoCol(2)).setValue('새 줄 메모');
  edit(env, sh, w.TODO_START + 10, w.checkCol(2), 1, 3);
  const day = dayOf(boot(env).week, '2026-11-11');
  assert.strictEqual(day.todos.length, 11);
  same(day.todos[0], [false, '09:00-10:00', '11번째 할 일', '새 줄 메모', ''], '시간 있는 일이 맨 위로, 메모 함께');
  assert.strictEqual(W(env).TODO_N, 13, '빈 줄 2개 유지');
  // 오늘 시트 새 줄에 적어도
  const t = T(env);
  ts(env).getRange(t.TODO_START + 11, 3).setValue('오늘 시트에서 12번째');
  edit(env, ts(env), t.TODO_START + 11, 3);
  assert.strictEqual(dayOf(boot(env).week, '2026-11-11').todos.length, 12);
  assert.strictEqual(T(env).TODO_N, 14);
  assert.strictEqual(W(env).TODO_N, 14);
});
test('할 일이 줄면 다시 기본 10줄, 꼭 끝낼 일도 늘었다가 기본 5줄로 (아래 칸은 제자리로)', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetMood('2026-11-11', 5, '');
  env.ctx.webSetTodos('2026-11-11', Array.from({ length: 15 }, (_, i) => [false, '', 'T' + i, '', '']));
  let res = env.ctx.webSetMust('2026-11-09', Array.from({ length: 7 }, (_, i) => [i === 0, '목표 ' + (i + 1)]));
  assert.strictEqual(res.weekPart.must.length, 7);
  let w = W(env);
  same([w.TODO_N, w.MUST_N, w.DAY_HEADER], [17, 9, 17], '꼭 끝낼 일이 늘면 요일 칸 전체가 4줄 밀림');
  assert.strictEqual(ws(env).getRange(w.MUST_START + 6, 2).getValue(), '목표 7');
  assert.strictEqual(ws(env).getRange(w.MUST_PROGRESS, w.MUST_WIDTH - 1).getValue(), '14%', '"+N개 더" 없음');
  assert.strictEqual(ws(env).getRange(w.TODO_START + 14, w.textCol(2)).getValue(), 'T14', '할 일도 밀린 자리에');
  assert.strictEqual(ws(env).getRange(w.MOOD, w.checkCol(2)).getValue(), '5 😄');
  assert.strictEqual(ts(env).getRange(T(env).MUST_START + 6, 2).getValue(), '목표 7');
  assert.strictEqual(T(env).MUST_N, 9);
  // 줄이기
  env.ctx.webSetTodos('2026-11-11', [[false, '', '하나만', '', '']]);
  env.ctx.webSetMust('2026-11-09', [[false, '목표 하나']]);
  w = W(env);
  same([w.TODO_N, w.MUST_N, w.DAY_HEADER, w.MOOD], [10, 5, 13, 29], '기본 줄 수·원래 위치');
  same(sheetTodos(env, 2).slice(0, 2), [[false, '', '하나만'], [false, '', '']]);
  assert.strictEqual(ws(env).getRange(w.MOOD, w.checkCol(2)).getValue(), '5 😄', '기분 칸도 원래 자리에');
  same([T(env).TODO_N, T(env).MUST_N], [10, 5]);
  assert.strictEqual(ts(env).getRange(T(env).MOOD, 2).getValue(), '5 😄');
});
test('습관·기분: 시트 ↔ 웹앱', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-11', '운동 30분', true);
  env.ctx.webSetMood('2026-11-11', 4, '좋은 하루');
  const w = W(env);
  assert.strictEqual(ws(env).getRange(w.HABIT_START + 1, w.HABIT_DAY_COL + 2).getValue(), true);
  assert.strictEqual(ws(env).getRange(w.MOOD, w.checkCol(2)).getValue(), '4 🙂');
  assert.strictEqual(ts(env).getRange(T(env).HABIT_START + 1, 1).getValue(), true);
  // 시트에서 습관 체크·기분 변경 → 웹앱
  ws(env).getRange(w.HABIT_START, w.HABIT_DAY_COL + 2).setValue(true);
  edit(env, ws(env), w.HABIT_START, w.HABIT_DAY_COL + 2);
  ts(env).getRange(T(env).MOOD, 2).setValue('2 🙁');
  edit(env, ts(env), T(env).MOOD, 2);
  const day = dayOf(boot(env).week, '2026-11-11');
  same(day.habits.slice(0, 2), [true, true]);
  same(day.mood, { score: 2, memo: '좋은 하루' });
});
test('거의 동시에 고치면 나중 저장으로 맞추고 conflict 표시', () => {
  const env = envAt(2026, 11, 11);
  const res1 = env.ctx.webSetTodos('2026-11-11', [[false, '', '웹 할 일', '', '']]);
  const rev = res1.day.rev;
  ws(env).getRange(W(env).TODO_START + 1, W(env).textCol(2)).setValue('시트 할 일'); // 그사이 시트에서 추가
  edit(env, ws(env), W(env).TODO_START + 1, W(env).textCol(2));
  const res2 = env.ctx.webSetTodos('2026-11-11', [[true, '', '웹 할 일', '', '']], rev);
  assert.strictEqual(res2.conflict, true);
  same(todoTexts(res2.day), ['웹 할 일'], '나중에 저장한 웹앱 내용으로');
  const res3 = env.ctx.webSetTodos('2026-11-11', [[false, '', '웹 할 일', '', '']], res2.day.rev);
  assert.strictEqual(res3.conflict, false);
});
test('서버 캐시: 시트 편집(onEdit) 뒤에는 새로 읽음', () => {
  const env = envAt(2026, 11, 11);
  const a = boot(env);
  ts(env).getRange(T(env).TODO_START, 3).setValue('캐시 확인'); // 트리거 없이 몰래 바꿈
  same(boot(env).week, a.week, '캐시된 값');
  edit(env, ts(env), T(env).TODO_START, 3);
  assert.strictEqual(todoTexts(dayOf(boot(env).week, '2026-11-11'))[0], '캐시 확인');
});

console.log('날짜 이동');
test('3주 전·3주 뒤로 이동해 할 일 추가·체크, 습관·기분·꼭 끝낼 일도 저장', () => {
  const env = envAt(2026, 11, 11);
  let wk = env.ctx.webGetWeek('2026-10-19'); // 3주 전 월요일
  assert.strictEqual(wk.monday, '2026-10-19');
  let res = env.ctx.webSetTodos('2026-10-21', [[false, '', '3주 전 할 일', '메모', '']]);
  res = env.ctx.webSetTodos('2026-10-21', [[true, '', '3주 전 할 일', '메모', '']], res.day.rev);
  same(res.day.todos, [[true, '', '3주 전 할 일', '메모', '']]);
  env.ctx.webSetHabit('2026-10-21', '물 8잔 마시기', true);
  env.ctx.webSetMood('2026-10-21', 5, '');
  env.ctx.webSetMust('2026-10-19', [[true, '지난 목표']]);
  wk = env.ctx.webGetWeek('2026-10-19');
  assert.strictEqual(dayOf(wk, '2026-10-21').habits[0], true);
  assert.strictEqual(dayOf(wk, '2026-10-21').mood.score, 5);
  same(wk.must, [[true, '지난 목표']]);
  assert.ok(archiveRows(env).find((r) => r.key === '2026-10-21' && r.item === '3주 전 할 일' && r.done === true), '지난 날짜는 기록에도');
  res = env.ctx.webSetTodos('2026-12-02', [[false, '10-11', '3주 뒤 할 일', '', '']]);
  res = env.ctx.webSetTodos('2026-12-02', [[true, '10:00-11:00', '3주 뒤 할 일', '', '']]);
  same(dayOf(env.ctx.webGetWeek('2026-11-30'), '2026-12-02').todos, [[true, '10:00-11:00', '3주 뒤 할 일', '', '']]);
  assert.ok(!archiveRows(env).find((r) => r.key === '2026-12-02'), '미래 날짜는 아직 기록 안 함');
  same(sheetTodos(env, 2)[0], [false, '', ''], '이번 주 시트는 그대로');
});
test('위클리 시트 주 시작을 바꾸면 그 주를 저장소에서 보여 줌', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTodos('2026-10-21', [[false, '', '지난 일', '', '']]);
  ws(env).getRange(W(env).WEEK_START.row, W(env).WEEK_START.col).setValue(new Date(2026, 9, 19));
  edit(env, ws(env), W(env).WEEK_START.row, W(env).WEEK_START.col);
  assert.strictEqual(sheetTodos(env, 2)[0][2], '지난 일');
});

console.log('자동 넘기기·새 주');
test('전날 미완료 → 오늘 할 일로 전부(개수 제한 없음), 메모도 같이, 전날은 미완 → + 기록 이월', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTodos('2026-11-11', [[false, '10:00', '보고서', '3쪽까지', ''], [true, '', '장보기', '', '']]);
  env.ctx.webSetTodos('2026-11-12', todos(Array.from({ length: 10 }, (_, i) => '목요일 ' + i)));
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  const wk = boot(env).week;
  const thu = dayOf(wk, '2026-11-12');
  assert.strictEqual(thu.todos.length, 11, '10개가 있어도 넘어옴 (넘어온 일 칸 없음)');
  same(thu.todos[0], [false, '10:00', '보고서', '3쪽까지', '넘어옴']);
  same(dayOf(wk, '2026-11-11').todos.map((r) => [r[2], r[4]]), [['보고서', '미완'], ['장보기', '']]);
  assert.strictEqual(sheetTodos(env, 2)[0][2], '미완 → 보고서', '시트에는 미완 → 표시');
  assert.strictEqual(ts(env).getRange(T(env).TODO_START, 3).getValue(), '↪ 보고서', '오늘 시트에는 ↪ 표시');
  const r = archiveRows(env).find((x) => x.key === '2026-11-11' && x.item === '보고서');
  assert.strictEqual(r.done, false);
  assert.strictEqual(r.value, '이월');
  env.run('dailyRollover_()');
  assert.strictEqual(dayOf(boot(env).week, '2026-11-12').todos.length, 11, '하루 한 번만');
});
test('일요일 미완료는 새 주 월요일로, 지난주 데이터는 저장소에 남아 다시 볼 수 있음', () => {
  const env = envAt(2026, 11, 15); // 일요일
  env.ctx.webSetTodos('2026-11-15', [[false, '20:00', '주간 정리', '', ''], [true, '', '빨래', '', '']]);
  env.ctx.today_ = () => new Date(2026, 10, 16);
  env.run('midnightJob()');
  assert.strictEqual(env.run('dateKey_(weeklyMonday_())'), '2026-11-16', '위클리 시트가 새 주로');
  same(sheetTodos(env, 0)[0], [false, '20:00', '↪ 주간 정리']);
  const last = env.ctx.webGetWeek('2026-11-09');
  same(dayOf(last, '2026-11-15').todos.map((r) => [r[2], r[4]]), [['주간 정리', '미완'], ['빨래', '']], '지난주 일요일 데이터 그대로');
  assert.ok(archiveRows(env).find((r) => r.key === '2026-11-15' && r.item === '주간 정리' && r.value === '이월'));
  assert.ok(archiveRows(env).find((r) => r.key === '2026-11-15' && r.item === '빨래' && r.done === true));
});
test('설정에서 끄면 넘기지 않음', () => {
  const env = envAt(2026, 11, 11);
  env.ss.getSheetByName('설정').getRange(4, 2).setValue(false);
  env.ctx.webSetTodos('2026-11-11', [[false, '', '보고서', '', '']]);
  env.ctx.today_ = () => new Date(2026, 10, 12);
  env.run('midnightJob()');
  assert.strictEqual(dayOf(boot(env).week, '2026-11-12').todos.length, 0);
});

console.log('할 일 시간·정렬');
test('시간 범위 "11-13"·"11~13"·"1100-1300" 변환 + 시작 시간순, 메모가 따라감', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webSetTodos('2026-11-11', [
    [false, '', '점심 약속 11~13', '2층', ''], [false, '', '1400-1530 회의', '', ''], [false, '', '9-10 운동', '하체', ''],
    [false, '15:00', '전화', '', ''], [false, '', '3-4쪽 읽기', '', ''],
  ]);
  same(res.day.todos, [
    [false, '09:00-10:00', '운동', '하체', ''], [false, '11:00-13:00', '점심 약속', '2층', ''], [false, '14:00-15:30', '회의', '', ''],
    [false, '15:00', '전화', '', ''], [false, '', '3-4쪽 읽기', '', ''],
  ]);
});

console.log('일정');
const NEW = null;
const itemRef = (items, text) => (items.find((it) => it.text === text) || {}).ref;
const sched = (week, key) => dayOf(week, key).schedule;
test('새 일정 → 구글 캘린더 1개 + 연동 메모 1줄, 시트·웹앱 어디서든 한 번만 보임', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webSaveItem('2026-11-11', NEW, { title: '할머니집 기차', start: '13:42', end: '' }, 2026, 11, 'req-1');
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events.length, 1);
  const rows = env.run('readEvents_()');
  assert.strictEqual(rows.filter((r) => r[1] === '캘린더').length, 1);
  assert.strictEqual(rows.filter((r) => r[1] === '메모').length, 1);
  same(res.day.schedule.map((x) => x.text), ['13:42 할머니집 기차'], '웹앱 한 번');
  same(res.month.dates[10].items.map((x) => x.text), ['13:42 할머니집 기차'], '먼슬리 한 번');
  const M = env.run('MONTHLY'), ms = env.ss.getSheetByName('먼슬리');
  assert.strictEqual(ms.getRange(M.memoRow(2), M.dayCol(2)).getValue(), '', '먼슬리 시트 메모 칸에 또 쓰지 않음 (캘린더 줄에만)');
  assert.ok(env.run('scheduleLineFormula_("B$13", 1, "")').startsWith('=IFERROR(INDEX(UNIQUE(FILTER('), '시트 일정 칸도 중복 제거');
  // 연타·재시도: 같은 요청 번호면 한 번만
  env.ctx.webSaveItem('2026-11-11', NEW, { title: '할머니집 기차', start: '13:42', end: '' }, 2026, 11, 'req-1');
  assert.strictEqual(env.events.length, 1);
  // 요청 번호가 달라도 같은 일정이면 서버에서 막음
  const again = env.ctx.webSaveItem('2026-11-11', NEW, { title: '할머니집 기차', start: '13:42', end: '' }, 2026, 11, 'req-2');
  assert.strictEqual(again.ok, false);
  assert.strictEqual(env.events.length, 1);
});
test('연동 안 된 같은 메모(예전 데이터)도 한 번만 보이고, 이미 생긴 중복은 목록 확인 후 하나만 남김', () => {
  const env = envAt(2026, 11, 11);
  const cal = env.ctx.CalendarApp.getDefaultCalendar();
  cal.createAllDayEvent('동아리', new Date(2026, 10, 12));
  cal.createAllDayEvent('동아리', new Date(2026, 10, 12)); // 실수로 두 번 만든 캘린더 일정
  const M = env.run('MONTHLY'), ms = env.ss.getSheetByName('먼슬리');
  ms.getRange(M.memoRow(2), M.dayCol(3)).setValue('동아리'); // 같은 글자의 연동 안 된 메모
  edit(env, ms, M.memoRow(2), M.dayCol(3));
  env.ctx.syncCalendarHourly();
  same(sched(boot(env).week, '2026-11-12').map((x) => x.text), ['동아리'], '화면엔 한 번');
  const dups = env.ctx.webFindDuplicates();
  same(dups.map((d) => [d.date, d.text, d.kind, d.count, d.remove.length]), [['2026-11-12', '동아리', 'calendar', 2, 1]], '지우기 전 목록');
  const r = env.ctx.webCleanDuplicates(dups.length);
  assert.strictEqual(r.left, 0);
  assert.strictEqual(env.events.length, 1, '캘린더에 하나만 남음');
  assert.throws(() => env.ctx.webCleanDuplicates(5), /바뀌었어요/);
});
test('일정 완료 체크: 체크된 일정은 아래로, 캘린더 일정은 그대로, 기록에 일정/수업', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetTimetable({ start: '', end: '', rows: [{ day: 2, start: '09:00', end: '10:30', name: '통계학' }] }, '2026-11-09');
  env.ctx.webSaveItem('2026-11-11', NEW, { title: '치과', start: '14:00', end: '15:00' }, null, null, 'a');
  let items = sched(boot(env).week, '2026-11-11');
  same(items.map((x) => [x.kind, x.text, x.done]), [['tt', '09:00-10:30 통계학', false], ['cal', '14:00-15:00 치과', false]]);
  let res = env.ctx.webCheckItem('2026-11-11', items[0].key, true);
  items = res.day.schedule;
  same(items.map((x) => [x.text, x.done]), [['14:00-15:00 치과', false], ['09:00-10:30 통계학', true]], '체크된 수업은 아래로');
  res = env.ctx.webCheckItem('2026-11-11', items[0].key, true);
  assert.strictEqual(env.events[0].title, '치과', '구글 캘린더 일정은 안 바뀜');
  const arc = archiveRows(env);
  assert.ok(arc.find((r) => r.cat === '수업' && r.item === '09:00-10:30 통계학' && r.done === true));
  assert.ok(arc.find((r) => r.cat === '일정' && r.item === '14:00-15:00 치과' && r.done === true));
  assert.strictEqual(env.ctx.webGetWeek('2026-11-16').days[2].schedule[0].done, false, '수업 체크는 그날만');
});
test('캘린더 일정 수정·삭제: ID로 찾아 반영 + 연동 메모도, 예전 일정은 날짜+제목으로', () => {
  const env = envAt(2026, 11, 11);
  let res = env.ctx.webSaveItem('2026-11-13', NEW, { title: '발표', start: '14:00', end: '', memo: '자료 챙기기' }, 2026, 11, 'x');
  const ev = env.events[0];
  assert.strictEqual(ev.description, '자료 챙기기');
  res = env.ctx.webSaveItem('2026-11-13', itemRef(res.day.schedule, '14:00 발표'), { title: '발표 리허설', start: '15:00', end: '16:00', memo: '' }, 2026, 11);
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events.length, 1);
  assert.strictEqual(ev.title, '발표 리허설');
  assert.strictEqual(env.run('readEvents_()').find((r) => r[1] === '메모')[2], '15:00-16:00 발표 리허설');
  res = env.ctx.webDeleteItem('2026-11-13', itemRef(res.day.schedule, '15:00-16:00 발표 리허설'), 2026, 11);
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events.length, 0);
  assert.strictEqual(env.run('readEvents_()').filter((r) => r[1] === '메모').length, 0);
  // 예전 일정 (ID 없음) → 날짜+제목
  env.ctx.CalendarApp.getDefaultCalendar().createAllDayEvent('가족 모임', new Date(2026, 10, 12));
  env.ctx.syncCalendarHourly();
  const sh = env.ss.getSheetByName('_일정');
  const i = env.run('readEvents_()').findIndex((r) => r[2] === '가족 모임');
  sh.getRange(i + 2, 5).setValue('');
  env.ctx.bumpCache_();
  res = env.ctx.webDeleteItem('2026-11-12', itemRef(sched(boot(env).week, '2026-11-12'), '가족 모임'), null, null);
  assert.ok(res.ok, res.message);
  assert.strictEqual(env.events.length, 0);
});
test('삭제 실패 이유 표시 + 다른 기기에서 바뀐 일정은 막음', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.CalendarApp.getDefaultCalendar().createAllDayEvent('없어질 일정', new Date(2026, 10, 12));
  env.ctx.syncCalendarHourly();
  const b = boot(env);
  env.events.length = 0;
  const res = env.ctx.webDeleteItem('2026-11-12', itemRef(sched(b.week, '2026-11-12'), '없어질 일정'), null, null);
  assert.strictEqual(res.ok, false);
  assert.ok(/찾지 못했어요/.test(res.message), res.message);
  const r2 = env.ctx.webDeleteItem('2026-11-12', { kind: 'cal', index: 0, text: '없는 일정' }, null, null);
  assert.ok(/다른 곳에서/.test(r2.message), r2.message);
});

console.log('시트 일정 추가 입력칸 (기존 규칙)');
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
test('시작+끝 / 시작만 / 없음 / 끝 오류 / 중복', () => {
  const env = envAt(2026, 11, 11);
  assert.ok(/추가됨/.test(addEvent(env, '14:00', '15:30', '회의')));
  assert.strictEqual(env.events[0].allDay, false);
  addEvent(env, '15:00', '', '스터디');
  assert.strictEqual(env.events[1].title, '15:00 스터디');
  addEvent(env, '', '', '생일');
  assert.strictEqual(env.events[2].allDay, true);
  assert.strictEqual(addEvent(env, '16:00', '15:00', '회의2'), '끝시간 확인');
  env.cache.clear();
  assert.ok(/중복/.test(addEvent(env, '', '', '생일')));
  assert.strictEqual(env.events.length, 3);
});

console.log('공휴일·시간표·습관·타임라인·먼슬리');
test('공휴일: 구독, 일정·할 일과 안 섞임, 수정 불가', () => {
  const env = envAt(2026, 11, 11);
  assert.strictEqual(env.holidays.subscribed, true);
  const day = dayOf(boot(env).week, '2026-11-11');
  assert.strictEqual(day.holiday, '테스트 기념일');
  same(day.schedule, []);
  assert.strictEqual(env.ctx.webGetMonth(2026, 12).dates[24].holiday, '기독탄신일');
  assert.strictEqual(env.ctx.webDeleteItem('2026-11-11', { kind: 'cal', index: 0, text: '테스트 기념일' }, null, null).ok, false);
});
test('고정 시간표: 학기 기간에만, 할 일 줄 차지 안 함', () => {
  const env = envAt(2026, 11, 11);
  const res = env.ctx.webSetTimetable({ start: '2026-09-01', end: '2026-11-11', rows: [
    { day: 1, start: '10:00', end: '11:30', name: '데이터마이닝' }, { day: 3, start: '9:00', end: '', name: '통계학' }] }, '2026-11-09');
  same(dayOf(res.week, '2026-11-10').schedule.map((x) => [x.kind, x.text]), [['tt', '10:00-11:30 데이터마이닝']]);
  same(dayOf(res.week, '2026-11-12').schedule, [], '학기 끝난 뒤');
  assert.strictEqual(dayOf(res.week, '2026-11-10').todos.length, 0);
});
test('습관 편집: 이름 바꾸면 이번 주 체크 유지, 지난 기록은 예전 이름 그대로', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-02', '독서 20분', true); // 지난주
  env.ctx.webSetHabit('2026-11-10', '독서 20분', true); // 이번 주
  const b = env.ctx.webSetHabits([{ name: '독서 30분', from: 2 }, { name: '물 8잔 마시기', from: 0 }, { name: '명상', from: null }]);
  same(b.habits.slice(0, 4), ['독서 30분', '물 8잔 마시기', '명상', '']);
  assert.strictEqual(dayOf(b.week, '2026-11-10').habits[0], true, '이번 주 체크는 새 이름으로');
  assert.strictEqual(env.ctx.webGetWeek('2026-11-02').days[0].habits[0], false, '지난주는 예전 이름 그대로 (새 이름으로는 안 보임)');
  assert.ok(archiveRows(env).find((r) => r.key === '2026-11-02' && r.item === '독서 20분' && r.done === true), '지난 기록 유지');
});
test('타임라인 저장·불러오기', () => {
  const env = envAt(2026, 11, 11);
  const slots = new Array(144).fill('');
  slots[54] = '공부'; slots[60] = '없는것';
  env.ctx.webSetTimeline('2026-11-11', slots);
  const got = env.ctx.webGetTimeline('2026-11-11').slots;
  assert.strictEqual(got[54], '공부');
  assert.strictEqual(got[60], '');
});
test('먼슬리: 기분·습관 달성률·월간 목표 (기록 시트 기준)', () => {
  const env = envAt(2026, 11, 11);
  env.ctx.webSetHabit('2026-11-10', '물 8잔 마시기', true);
  env.ctx.webSetMood('2026-11-10', 5, '');
  env.ctx.webSetMonthGoals(2026, 11, [[false, '책 2권'], [true, '저축']]);
  const M = env.ctx.webGetMonth(2026, 11);
  assert.strictEqual(M.dates[9].mood, 5);
  assert.strictEqual(M.habits[0].done, 1);
  same(M.goals.slice(0, 2), [[false, '책 2권'], [true, '저축']]);
});

console.log('예전 데이터 옮기기');
test('위클리 시트 이번 주 + 기록 시트 지난 기록 → 저장소 (손실 없음, 백업, 줄 수 보고)', () => {
  const env = envAt(2026, 11, 11);
  // 예전 구조 흉내: 저장소 비우고 버전 지움
  ['할일저장소', '하루저장소', '주간저장소'].forEach((n) => { const s = env.ss.getSheetByName(n); s.getRange(2, 1, 50, 8).clearContent(); });
  env.props.delete('STORE_VERSION');
  env.run('STORE_MEMO_ = {}');
  const w = W(env), sh = ws(env);
  sh.getRange(w.TODO_START, w.checkCol(2), 3, 3).setValues([[false, '10:00-11:00', '미완 → 보고서'], [true, '', '장보기'], [false, '', '↪ 운동']]);
  sh.getRange(w.TODO_START, w.memoCol(2)).setValue('보고서 메모');
  sh.getRange(w.HABIT_START, w.HABIT_DAY_COL + 1).setValue(true);
  sh.getRange(w.MOOD, w.checkCol(2)).setValue('4 🙂');
  sh.getRange(w.MUST_START, 1, 1, 2).setValues([[true, '포트폴리오']]);
  sh.getRange(w.GOAL_ROW, w.INFO_VALUE_COL).setValue('플래너 완성');
  const arc = env.ss.getSheetByName('기록');
  arc.getRange(2, 1, 4, 6).setValues([
    [new Date(2026, 10, 3), '할일', '지난주 일', true, '', '2026-11-02'],
    [new Date(2026, 10, 3), '습관', '운동 30분', true, '', '2026-11-02'],
    [new Date(2026, 10, 3), '기분', '좋음', '', 5, '2026-11-02'],
    [new Date(2026, 10, 2), '주간목표', '지난 목표', false, '', '2026-11-02'],
  ]);
  env.props.set('CARRY_OVERFLOW', JSON.stringify({ date: '2026-11-11', items: [{ time: '', text: '넘어온 일', memo: '' }] }));
  boot(env); // 아무 진입점이나 → 자동으로 옮김
  assert.strictEqual(env.props.get('STORE_VERSION'), '1');
  assert.strictEqual(env.copies.length, 1, '옮기기 전 시트 사본 백업');
  const report = JSON.parse(env.props.get('STORE_REPORT'));
  same([report.todos, report.day, report.week], [5, 4, 3], '할 일 5(이번 주 3 + 지난주 1 + 넘어온 일 1) · 습관·기분 4 · 주간 3');
  const day = dayOf(boot(env).week, '2026-11-11');
  same(day.todos, [[false, '', '운동', '', '넘어옴'], [false, '', '넘어온 일', '', '넘어옴'], [false, '10:00-11:00', '보고서', '보고서 메모', '미완'], [true, '', '장보기', '', '']]);
  same(day.mood.score, 4);
  const last = env.ctx.webGetWeek('2026-11-02');
  same(dayOf(last, '2026-11-03').todos, [[true, '', '지난주 일', '', '']]);
  assert.strictEqual(dayOf(last, '2026-11-03').habits[1], true);
  same(last.must, [[false, '지난 목표']]);
  assert.strictEqual(archiveRows(env).filter((r) => r.key === '2026-11-03').length >= 3, true, '기록 시트는 그대로');
  assert.ok(/옮김/.test(env.ss.getSheetByName('설정').getRange(26, 2).getValue()), '설정 시트에 줄 수 보고');
});
test('예전 문서 속성 값은 스크립트 속성으로 옮겨짐', () => {
  const env = createEnv();
  env.docProps.set('CARRY_DONE', '2026-11-10');
  assert.strictEqual(env.run("props_().getProperty('CARRY_DONE')"), '2026-11-10');
});

console.log(`\n${passed}개 통과` + (process.exitCode ? ' (실패 있음)' : ''));
