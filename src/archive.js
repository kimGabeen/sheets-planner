/**
 * archive.js
 * 기록 시트: 분석용 "한 줄에 하나" 형태로 쌓는 저장소
 * 열: 날짜 | 구분 | 항목 | 완료여부 | 값 | 주차
 * - 같은 날짜+구분 묶음은 지우고 다시 써서 중복 없이 덮어씀
 * - 행 삭제·추가만 하고 시트를 통째로 지우는 일은 절대 없음
 */

/** 기록 시트 만들기 (이미 있으면 머리글·서식만 확인, 데이터는 그대로) */
function buildArchive_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.ARCHIVE);
  const H = CONFIG.ARCHIVE_HEADERS;
  ensureSize_(sh, 50, H.length);
  if (!sh.getRange(1, 1).getValue()) sh.getRange(1, 1, 1, H.length).setValues([H]);
  styleHeader_(sh.getRange(1, 1, 1, H.length), theme);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('F:F').setNumberFormat('@'); // 주차는 글자로 (CSV에서 그대로)
  sh.setColumnWidth(1, 100).setColumnWidth(2, 80).setColumnWidth(3, 240)
    .setColumnWidth(4, 80).setColumnWidth(5, 60).setColumnWidth(6, 100);
  return sh;
}

/** 기록 전체 읽기 → [{row, key, cat, item, done, value, week}] */
function readArchive_() {
  const sh = sheet_(CONFIG.SHEETS.ARCHIVE);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues().map(function (r, i) {
    const d = dateOnly_(r[0]);
    return { row: i + 2, key: d ? dateKey_(d) : String(r[0]), cat: r[1], item: r[2], done: r[3], value: r[4], week: String(r[5]) };
  });
}

/**
 * 기록 덮어쓰기
 * keys: [{date, cat}] 에 해당하는 기존 행을 지우고 rows를 추가
 */
function replaceArchive_(keys, rows) {
  const sh = sheet_(CONFIG.SHEETS.ARCHIVE);
  const want = {};
  keys.forEach(function (k) { want[dateKey_(k.date) + '|' + k.cat] = true; });
  const del = readArchive_().filter(function (r) { return want[r.key + '|' + r.cat]; })
    .map(function (r) { return r.row; });

  // 빈 행 여유 확보 (고정 행만 남기고 전부 지우는 오류 방지)
  if (sh.getMaxRows() - sh.getLastRow() < 5) sh.insertRowsAfter(sh.getMaxRows(), 50);

  // 연속 구간끼리 묶어 아래쪽부터 삭제
  for (let i = del.length - 1; i >= 0;) {
    const end = del[i];
    let start = end;
    while (i > 0 && del[i - 1] === start - 1) { i--; start--; }
    sh.deleteRows(start, end - start + 1);
    i--;
  }
  if (rows.length) {
    const at = Math.max(sh.getLastRow(), 1) + 1;
    ensureSize_(sh, at + rows.length + 5, 6);
    sh.getRange(at, 1, rows.length, 6).setValues(rows);
  }
  // 날짜·구분 순으로 정리 (데이터는 그대로, 순서만)
  const last = sh.getLastRow();
  if (last > 2) sh.getRange(2, 1, last - 1, 6).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

/** 위클리의 하루치(할 일·습관·기분)를 기록 행으로 만들기 */
function buildDayRows_(date) {
  const L = WEEKLY, C = CONFIG.CATEGORY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const d = weeklyIndexOf_(date);
  if (d < 0) return [];
  const wk = weekKey_(date);
  const rows = [];

  // 할 일: "미완 →"(넘겨진 원래 일)은 FALSE + 값 "이월", 표시 기호는 떼고 저장
  sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 3).getValues().forEach(function (r) {
    const raw = String(r[2]).trim();
    if (!raw) return;
    const missed = isMissed_(raw);
    rows.push([date, C.TODO, stripCarry_(raw), missed ? false : r[0] === true, missed ? CONFIG.CARRY.VALUE : '', wk]);
  });

  const checks = sh.getRange(L.HABIT_START, L.HABIT_DAY_COL + d, CONFIG.HABIT_COUNT, 1).getValues();
  getHabits_().forEach(function (name, i) {
    if (name) rows.push([date, C.HABIT, name, checks[i][0] === true, '', wk]);
  });

  const score = moodScore_(sh.getRange(L.MOOD, L.checkCol(d)).getValue());
  const memo = String(sh.getRange(L.MOOD_MEMO, L.checkCol(d)).getValue()).trim();
  if (score || memo) rows.push([date, C.MOOD, memo, '', score, wk]); // 기분: 항목=메모, 값=점수
  return rows;
}

/** 하루치 덮어쓰기 키 */
function dayKeys_(date) {
  const C = CONFIG.CATEGORY;
  return [{ date: date, cat: C.TODO }, { date: date, cat: C.HABIT }, { date: date, cat: C.MOOD }];
}

/** 편집 즉시 그날 기록 갱신 (미래 날짜는 밤에 저장) */
function archiveDayLive_(date) {
  if (!date || date > today_()) return;
  replaceArchive_(dayKeys_(date), buildDayRows_(date));
}

/** 이번 주 꼭 끝낼 일 → 주간목표 (날짜 = 그 주 월요일) */
function weekGoalRows_() {
  const sh = sheet_(CONFIG.SHEETS.WEEKLY), L = WEEKLY;
  const mon = weeklyMonday_();
  if (!mon) return { keys: [], rows: [] };
  const rows = [];
  sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 2).getValues().forEach(function (r) {
    const text = String(r[1]).trim();
    if (text) rows.push([mon, CONFIG.CATEGORY.WEEK_GOAL, text, r[0] === true, '', dateKey_(mon)]);
  });
  return { keys: [{ date: mon, cat: CONFIG.CATEGORY.WEEK_GOAL }], rows: rows };
}

/** 이번 달 목표 → 월간목표 (날짜 = 그 달 1일) */
function monthGoalRows_() {
  const sh = sheet_(CONFIG.SHEETS.MONTHLY), L = MONTHLY;
  const ym = getMonthlyYM_();
  if (!ym) return { keys: [], rows: [] };
  const first = new Date(ym.y, ym.m - 1, 1);
  const rows = [];
  sh.getRange(L.GOAL_START, L.PANEL_COL, CONFIG.MONTH_GOALS, 2).getValues().forEach(function (r) {
    const text = String(r[1]).trim();
    if (text) rows.push([first, CONFIG.CATEGORY.MONTH_GOAL, text, r[0] === true, '', weekKey_(first)]);
  });
  return { keys: [{ date: first, cat: CONFIG.CATEGORY.MONTH_GOAL }], rows: rows };
}

/** 주간목표만 저장 */
function archiveWeekGoals_() {
  const g = weekGoalRows_();
  replaceArchive_(g.keys, g.rows);
}

/** 월간목표만 저장 */
function archiveMonthGoals_() {
  const g = monthGoalRows_();
  replaceArchive_(g.keys, g.rows);
}

/** 위클리 주의 월요일 ~ until까지 하루치 + 주간·월간 목표 저장 → 저장한 행 수 */
function saveWeek_(until) {
  const mon = weeklyMonday_();
  if (!mon) return 0;
  let keys = [], rows = [];
  for (let d = 0; d < 7; d++) {
    const date = addDays_(mon, d);
    if (date > until) break;
    keys = keys.concat(dayKeys_(date));
    rows = rows.concat(buildDayRows_(date));
  }
  [weekGoalRows_(), monthGoalRows_()].forEach(function (g) {
    keys = keys.concat(g.keys);
    rows = rows.concat(g.rows);
  });
  replaceArchive_(keys, rows);
  return rows.length;
}

/** 지금까지(오늘 포함) 전부 저장 */
function saveAll_() {
  const n = saveWeek_(today_());
  setStatus_('save', '✓ ' + n + '행');
  return n;
}

/** 특정 주차의 기록 행 수 (새 주 시작 전 확인용) */
function countArchiveWeek_(week) {
  return readArchive_().filter(function (r) { return r.week === week; }).length;
}
