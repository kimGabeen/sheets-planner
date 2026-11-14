/**
 * test/bench.js — 웹앱 작업별 서버 처리 비용 측정
 * 실행: npm run bench  (결과를 JSON으로 저장하려면: node test/bench.js 결과.json)
 *
 * Apps Script 서버를 직접 돌릴 수 없어서, 목업 시트에서 작업마다 서비스 호출 수를 정확히 세고
 * 일반적인 Apps Script 호출 비용(아래 COST, ms)을 곱해 처리 시간을 추정한다. (실측은 웹앱 "속도 기록"에서)
 */
process.env.TZ = 'Asia/Seoul';
const fs = require('fs');
const mock = require('./mock');
const { createEnv, calls } = mock;
let fresh = () => {};
const resetCalls = () => { fresh(); mock.resetCalls(); };

// 호출 한 번당 대략적인 비용 (ms): 시트 읽기·쓰기, 행 삭제·정렬, 수식 쓰기, 캘린더 읽기·쓰기, 캐시, 속성,
// flushRead = 쓰기 뒤 첫 읽기가 시트 재계산을 기다리는 추가 시간
const COST = { read: 35, write: 70, struct: 120, formula: 70, calRead: 450, calWrite: 600, cache: 15, props: 20, flushRead: 150 };
const estimate = (c) => Object.keys(COST).reduce((sum, k) => sum + (c[k] || 0) * COST[k], 0);

// 3개월치 실사용 규모 데이터
const env = createEnv();
env.ctx.today_ = () => new Date(2026, 10, 11);
env.run('setup()');
const c = env.ctx;
const habits = ['물 8잔 마시기', '운동 30분', '독서 20분'];
for (let i = 90; i >= 0; i--) {
  const d = new Date(2026, 10, 11 - i);
  c.setTodos_(d, Array.from({ length: 8 }, (_, k) => ({ done: k < 5, start: k < 2 ? '0' + (9 + k) + ':00' : '', end: '', text: '할 일 ' + k, memo: k === 0 ? '메모' : '', flag: '' })));
  c.setHabitsDone_(d, { [habits[0]]: true, [habits[1]]: i % 2 === 0, [habits[2]]: true });
  c.setMoodData_(d, 1 + (i % 5), i % 3 ? '' : '한 줄');
  if (i % 4 === 0) c.CalendarApp.getDefaultCalendar().createEvent('약속 ' + i, new Date(2026, 10, 11 - i, 14), new Date(2026, 10, 11 - i, 15));
}
c.saveRange_(new Date(2026, 7, 13), new Date(2026, 10, 11));
c.syncCalendar_();

const ops = [
  ['앱 열기 (캐시 없음)', () => { c.bumpCache_(); resetCalls(); c.webBoot(); }],
  ['앱 열기 (캐시 있음)', () => { c.webBoot(); resetCalls(); c.webBoot(); }],
  ['할 일 체크 (웹앱)', () => {
    const day = c.webBoot().week.days[2];
    const rows = JSON.parse(JSON.stringify(day.todos)); rows[5][0] = !rows[5][0];
    resetCalls(); c.webSetTodos(day.key, rows, day.rev);
  }],
  ['일정 저장 (웹앱)', () => { resetCalls(); c.webSaveItem('2026-11-12', null, { title: '측정용 ' + Math.random(), start: '10:00', end: '11:00', memo: '' }, null, null, 'r' + Math.random()); }],
  ['날짜 이동: 지난주 (웹앱)', () => { c.bumpCache_(); resetCalls(); c.webGetWeek('2026-11-02'); }],
  ['날짜 이동: 3주 전 (웹앱)', () => { c.bumpCache_(); resetCalls(); c.webGetWeek('2026-10-19'); }],
  ['할 일 체크 (시트 편집)', () => {
    const sh = env.ss.getSheetByName('위클리'), W = env.run('WEEKLY');
    const cell = sh.getRange(W.TODO_START + 1, W.checkCol(2));
    cell.setValue(!cell.getValue());
    resetCalls(); c.handleEdit({ range: cell });
  }],
];

// 실제 Apps Script는 실행마다 전역 변수가 새로 시작하므로, 작업마다 메모리 캐시를 비우고 잰다
fresh = () => env.run('STORE_MEMO_ = {}; if (typeof RUN_MEMO_ !== "undefined") RUN_MEMO_ = {};');
const results = {};
console.log('| 작업 | 시트 읽기 | 시트 쓰기 | 행 삭제·정렬 | 재계산 대기 | 캘린더 | 기타(캐시·속성) | 추정 시간 |');
console.log('|---|---:|---:|---:|---:|---:|---:|---:|');
ops.forEach(([name, run]) => {
  run();
  const snap = Object.assign({}, calls);
  results[name] = { calls: snap, ms: estimate(snap) };
  console.log('| ' + name + ' | ' + snap.read + ' | ' + (snap.write + snap.formula) + ' | ' + snap.struct + ' | ' + snap.flushRead +
    ' | ' + (snap.calRead + snap.calWrite) + ' | ' + (snap.cache + snap.props) + ' | ~' + (estimate(snap) / 1000).toFixed(1) + '초 |');
});
if (process.argv[2]) fs.writeFileSync(process.argv[2], JSON.stringify(results, null, 2));
