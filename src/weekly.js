/**
 * weekly.js
 * 위클리 시트: 맥북·아이패드 가로 화면용 한 주 보기
 * - 월~일 7칸 (완료율 / 일정 / 할 일 10줄 / 합계 / 기분)
 * - 오른쪽 습관 트래커, 상단 목표·꼭 끝낼 일·차트
 */

/** 위클리 시트 레이아웃·수식·서식 그리기 (입력 값은 유지) */
function buildWeekly_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.WEEKLY);
  const L = WEEKLY, C = CONFIG.COLORS, Z = CONFIG.SIZE;
  const light = tint_(theme, CONFIG.TINT.LIGHT), pale = tint_(theme, CONFIG.TINT.PALE);
  ensureSize_(sh, L.MOOD_MEMO + 5, L.LAST_COL);
  // 예전 레이아웃이면 입력값을 챙겨 두고 시트를 비운 뒤 새로 그림
  const old = needsMigration_() && sh.getLastRow() > 0 ? collectWeeklyV1_(sh) : null;
  if (old) wipeSheet_(sh);
  resetLayout_(sh);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(10).setVerticalAlignment('middle');

  // 열 너비·행 높이
  sh.setColumnWidth(1, Z.LABEL_COL);
  for (let d = 0; d < 7; d++) {
    sh.setColumnWidth(L.checkCol(d), Z.CHECK_COL);
    sh.setColumnWidth(L.timeCol(d), Z.TIME_COL);
    sh.setColumnWidth(L.textCol(d), Z.TEXT_COL);
  }
  sh.setColumnWidth(L.HABIT_NAME_COL - 1, Z.GAP_COL);
  sh.setColumnWidth(L.HABIT_NAME_COL, Z.HABIT_NAME_COL);
  sh.setColumnWidths(L.HABIT_DAY_COL, 7, Z.HABIT_DAY_COL);
  sh.setColumnWidth(L.HABIT_BAR_COL, 80).setColumnWidth(L.HABIT_PCT_COL, 45);
  sh.setRowHeights(1, L.MOOD_MEMO + 2, Z.ROW);
  sh.setRowHeights(L.SCHED_START, CONFIG.SCHEDULE_LINES, 22);

  buildWeeklyTop_(sh, theme, pale);
  buildWeeklyDays_(sh, theme, pale);
  buildWeeklyHabits_(sh, theme);
  buildWeeklyCharts_(sh, theme, light);
  applyWeeklyRules_(sh, theme, pale);
  // 할 일 메모 (숨김, 할 일과 같은 줄)
  sh.getRange(L.TODO_START - 1, L.MEMO_COL0, 1, 7).setValues([WEEKDAY_NAMES_.map(function (n) { return n + ' 메모'; })]);
  sh.hideColumns(L.MEMO_COL0, 7);
  if (old) restoreWeekly_(sh, old);
  if (props_().getProperty('STORE_VERSION') === CONFIG.STORE_VERSION) renderWeekView_();
}

/** 상단: 제목, 주 시작, 수동 실행 체크박스, 목표·보상·확언, 꼭 끝낼 일 */
function buildWeeklyTop_(sh, theme, pale) {
  const L = WEEKLY, C = CONFIG.COLORS;
  sh.getRange(1, 1, 1, 2).merge().setValue('위클리').setFontSize(18).setFontWeight('bold')
    .setFontColor(theme);
  sh.getRange(1, 3, 1, 2).merge().setValue('주 시작 (월)').setFontColor(C.SUB).setHorizontalAlignment('right');
  dateInput_(sh.getRange(L.WEEK_START.row, L.WEEK_START.col, 1, 3).merge(), 'yyyy-mm-dd (ddd)');
  sh.getRange(L.WEEK_START.row, L.WEEK_START.col).setFontWeight('bold').setHorizontalAlignment('center');

  // 모바일용 수동 실행 체크박스 (메뉴 대신)
  [[L.ACTION_SAVE, '지금 기록 저장'], [L.ACTION_SYNC, '캘린더 새로고침']].forEach(function (a) {
    checkbox_(sh.getRange(a[0].row, a[0].col));
    sh.getRange(a[0].row, 2, 1, 3).merge().setValue(a[1]);
    sh.getRange(a[0].row, L.ACTION_RESULT_COL, 1, 3).merge().setFontColor(C.SUB).setFontSize(9);
  });

  // 목표·보상·확언 (각 1줄)
  [[L.GOAL_ROW, '이번 주 목표'], [L.REWARD_ROW, '보상'], [L.AFFIRM_ROW, '확언']].forEach(function (a) {
    sh.getRange(a[0], L.INFO_LABEL_COL, 1, L.INFO_VALUE_COL - L.INFO_LABEL_COL).merge().setValue(a[1])
      .setFontWeight('bold').setFontColor(theme);
    underline_(sh.getRange(a[0], L.INFO_VALUE_COL, 1, L.INFO_VALUE_END - L.INFO_VALUE_COL + 1).merge());
  });

  // 이번 주 꼭 끝낼 일
  const MW = L.MUST_WIDTH;
  styleHeader_(sh.getRange(L.MUST_HEADER, 1, 1, MW).merge().setValue('이번 주 꼭 끝낼 일'), theme);
  for (let i = 0; i < CONFIG.MUST_COUNT; i++) {
    checkbox_(sh.getRange(L.MUST_START + i, 1));
    underline_(sh.getRange(L.MUST_START + i, 2, 1, MW - 1).merge());
  }
  sh.getRange(L.MUST_PROGRESS, 1).setValue('진행률').setFontColor(C.SUB);
  // 진행률은 저장소의 전체 개수 기준으로 renderWeekView_가 채움
  sh.getRange(L.MUST_PROGRESS, 2, 1, MW - 3).merge();
  sh.getRange(L.MUST_PROGRESS, MW - 1, 1, 2).merge().setHorizontalAlignment('right').setFontWeight('bold').setFontSize(9);
}

/** 월~일 7칸 */
function buildWeeklyDays_(sh, theme, pale) {
  const L = WEEKLY, C = CONFIG.COLORS;
  const ws = a1_(L.WEEK_START.row, L.WEEK_START.col, true);
  const ev = CONFIG.SHEETS.EVENTS;

  // 왼쪽 라벨
  const labels = [[L.RATE, '완료율'], [L.SCHED_START, '일정'], [L.TODO_START, '할 일'],
    [L.COUNT, '합계'], [L.MOOD, '기분'], [L.MOOD_MEMO, '메모']];
  labels.forEach(function (a) {
    sh.getRange(a[0], 1).setValue(a[1]).setFontColor(C.SUB).setFontSize(9);
  });
  sh.getRange(L.TODO_START + 1, 1).setValue('(시간 선택)').setFontColor(C.SUB).setFontSize(8);

  for (let d = 0; d < 7; d++) {
    const c = L.checkCol(d), t = L.textCol(d);
    const date = colLetter_(c) + '$' + L.DAY_HEADER; // 예: B$13
    const S = L.DAY_SPAN;
    const chk = a1_(L.TODO_START, c) + ':' + a1_(L.COUNT - 1, c);
    const txt = a1_(L.TODO_START, t) + ':' + a1_(L.COUNT - 1, t);
    const done = 'COUNTIFS(' + chk + ',TRUE,' + txt + ',"?*")';
    const tot = 'COUNTIF(' + txt + ',"?*")';

    // 1) 요일 머리글 (날짜 자동)
    styleHeader_(sh.getRange(L.DAY_HEADER, c, 1, S).merge()
      .setFormula('=IF(' + ws + '="","",' + ws + '-WEEKDAY(' + ws + ',3)+' + d + ')')
      .setNumberFormat('m/d (ddd)').setHorizontalAlignment('center'), theme);

    // 2) 그날 완료율 (○◔◑◕● 원형 + %, 전체 할 일 기준 — renderWeekView_가 채움)
    sh.getRange(L.RATE, c, 1, S).merge()
      .setHorizontalAlignment('center').setFontSize(13).setFontColor(theme).setFontWeight('bold');

    // 3) 일정 (먼슬리와 같은 데이터, 읽기 전용 회색)
    for (let k = 1; k <= CONFIG.SCHEDULE_LINES; k++) {
      sh.getRange(L.SCHED_START + k - 1, c, 1, S).merge().setFormula(scheduleLineFormula_(date, k, ''))
        .setFontColor(C.SUB).setFontSize(9).setWrap(false);
    }

    // 4) 할 일 10줄 (체크박스 + 시간(선택) + 내용) — 시간은 글자로 저장해 정렬
    checkbox_(sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, 1));
    sh.getRange(L.TODO_START, L.timeCol(d), CONFIG.TODO_ROWS, 1).setNumberFormat('@')
      .setHorizontalAlignment('center').setFontSize(9).setFontColor(theme);
    underline_(sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, S));

    // 5) 완료 / 남은 개수 (+10개 넘으면 "+N개 더")
    sh.getRange(L.COUNT, c, 1, S).merge()
      .setHorizontalAlignment('center').setFontColor(C.SUB).setFontSize(9);

    // 6) 기분 점수 + 한 줄 메모
    moodInput_(sh.getRange(L.MOOD, c, 1, S).merge());
    sh.getRange(L.MOOD_MEMO, c, 1, S).merge().setFontSize(9);
  }
}

/** 오른쪽 습관 트래커 (습관 10개 × 월~일) */
function buildWeeklyHabits_(sh, theme) {
  const L = WEEKLY, C = CONFIG.COLORS;
  styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_NAME_COL).setValue('습관'), theme);
  for (let d = 0; d < 7; d++) {
    styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_DAY_COL + d)
      .setFormula('=' + a1_(L.DAY_HEADER, L.checkCol(d))).setNumberFormat('ddd')
      .setHorizontalAlignment('center'), theme);
  }
  styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_BAR_COL, 1, 2).merge().setValue('주간')
    .setHorizontalAlignment('center'), theme);

  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    const r = L.HABIT_START + i;
    const row = a1_(r, L.HABIT_DAY_COL) + ':' + a1_(r, L.HABIT_DAY_COL + 6);
    const name = a1_(r, L.HABIT_NAME_COL);
    const p = 'COUNTIF(' + row + ',TRUE)/7';
    sh.getRange(r, L.HABIT_NAME_COL).setFormula('=' + habitRef_(i) + '&""').setFontSize(9);
    checkbox_(sh.getRange(r, L.HABIT_DAY_COL, 1, 7));
    sh.getRange(r, L.HABIT_BAR_COL).setFormula(
      '=IF(' + name + '="","",SPARKLINE(' + p + ',{"charttype","bar";"max",1;"color1","' + theme + '"}))');
    sh.getRange(r, L.HABIT_PCT_COL).setFormula('=IF(' + name + '="","",' + p + ')')
      .setNumberFormat('0%').setFontColor(C.SUB).setFontSize(9).setHorizontalAlignment('right');
    underline_(sh.getRange(r, L.HABIT_NAME_COL, 1, 9));
  }
}

/** 이번 주 진행률 차트 (요일별 막대 + 전체 도넛) */
function buildWeeklyCharts_(sh, theme, light) {
  const L = WEEKLY;
  const dc = L.CHART_DATA_COL, top = L.DAY_HEADER;

  // 숨김 데이터: 요일 | 완료 | 남음
  sh.getRange(top, dc, 1, 3).setValues([['요일', '완료', '남음']]);
  for (let d = 0; d < 7; d++) {
    const c = L.checkCol(d), t = L.textCol(d);
    const chk = a1_(L.TODO_START, c) + ':' + a1_(L.COUNT - 1, c);
    const txt = a1_(L.TODO_START, t) + ':' + a1_(L.COUNT - 1, t);
    sh.getRange(top + 1 + d, dc).setFormula('=TEXT(' + a1_(L.DAY_HEADER, c) + ',"ddd")'); // 완료·남음은 renderWeekView_
  }
  const sumDone = 'SUM(' + a1_(top + 1, dc + 1) + ':' + a1_(top + 7, dc + 1) + ')';
  const sumLeft = 'SUM(' + a1_(top + 1, dc + 2) + ':' + a1_(top + 7, dc + 2) + ')';
  sh.getRange(top + 9, dc, 2, 2).setFormulas([['="완료"', '=' + sumDone], ['="남음"', '=' + sumLeft]]);
  sh.hideColumns(dc, 3);

  // 전체 진행률 문구
  sh.getRange(L.OVERALL.row, L.OVERALL.col, 1, 10).merge().setFormula(
    '=LET(d,' + sumDone + ',t,d+' + sumLeft + ',"이번 주 할 일 "&TEXT(IF(t=0,0,d/t),"0%")&"  ("&d&"/"&t&")")')
    .setFontColor(theme).setFontWeight('bold').setHorizontalAlignment('center');

  const bar = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(top, dc, 8, 3))
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setPosition(L.MUST_HEADER, L.INFO_LABEL_COL, 0, 0)
    .setOption('title', '요일별 완료 개수')
    .setOption('isStacked', true)
    .setOption('colors', [theme, light])
    .setOption('legend', { position: 'bottom' })
    .setOption('vAxis', { minValue: 0, gridlines: { color: '#F2F2F2' } })
    .setOption('width', 1000).setOption('height', 205)
    .build();
  sh.insertChart(bar);

  const donut = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange(top + 9, dc, 2, 2))
    .setNumHeaders(0)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setPosition(1, L.HABIT_NAME_COL, 0, 0)
    .setOption('title', '이번 주 전체')
    .setOption('pieHole', 0.55)
    .setOption('colors', [theme, '#EEEEEE'])
    .setOption('legend', { position: 'none' })
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 470).setOption('height', 320)
    .build();
  sh.insertChart(donut);
}

/** 조건부 서식: 할 일 체크 시 취소선, 미완 회색, 기분 배경, 습관 체크, 오늘 표시 */
function applyWeeklyRules_(sh, theme, pale) {
  const L = WEEKLY, C = CONFIG.COLORS;
  const first = L.checkCol(0), last = L.LAST_DAY_COL, width = last - first + 1;
  const todo = sh.getRange(L.TODO_START, first, CONFIG.TODO_ROWS, width);
  // 요일 블록 안에서 체크박스 열의 위치 (블록 = 체크|시간|내용)
  const idx = 'COLUMN()-' + (first - 1) + '-MOD(COLUMN()-' + first + ',' + L.DAY_SPAN + ')';
  const rowRef = function (r, abs) {
    const a = abs ? '$' : '';
    return '$' + colLetter_(first) + a + r + ':$' + colLetter_(last) + a + r;
  };
  const todoRow = rowRef(L.TODO_START, false);
  const doneCond = '=INDEX(' + todoRow + ',1,' + idx + ')=TRUE';
  const missedCond = '=LEFT(INDEX(' + todoRow + ',1,' + idx + '+2),' + CONFIG.CARRY.MISSED.trim().length +
    ')="' + CONFIG.CARRY.MISSED.trim() + '"';
  const score = 'IFERROR(VALUE(LEFT(INDEX(' + rowRef(L.MOOD, true) + ',1,' + idx + '),1)),0)';

  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(doneCond)
      .setStrikethrough(true).setFontColor(C.SUB).setBackground(pale).setRanges([todo]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(missedCond)
      .setFontColor(C.SUB).setRanges([todo]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + a1_(L.HABIT_START, L.HABIT_DAY_COL) + '=TRUE')
      .setBackground(pale).setRanges([sh.getRange(L.HABIT_START, L.HABIT_DAY_COL, CONFIG.HABIT_COUNT, 7)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=INDEX(' + rowRef(L.DAY_HEADER, true) + ',1,' + idx + ')=TODAY()')
      .setBackground(shade_(theme, 0.35)).setRanges([sh.getRange(L.DAY_HEADER, first, 1, width)]).build(),
  ].concat(moodRules_(sh, [sh.getRange(L.MOOD, first, 2, width)], score, theme));
  sh.setConditionalFormatRules(rules);
}

/* ---------- 레이아웃 버전 이전 (예전 입력값 보존) ---------- */

/** 아직 예전 레이아웃인지 */
function needsMigration_() {
  return props_().getProperty('LAYOUT_VERSION') !== CONFIG.LAYOUT_VERSION;
}

/** 시트 내용·서식·검증·숨김 열을 모두 비움 (마이그레이션 전용) */
function wipeSheet_(sh) {
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart();
  sh.clear();
  all.clearDataValidations();
  sh.showColumns(1, sh.getMaxColumns());
}

/** 예전(2열) 위클리에서 입력값 읽기 */
function collectWeeklyV1_(sh) {
  const L = WEEKLY, V = WEEKLY_V1, n = CONFIG.TODO_ROWS;
  const data = {
    weekStart: sh.getRange(V.WEEK_START.row, V.WEEK_START.col).getValue(),
    info: [L.GOAL_ROW, L.REWARD_ROW, L.AFFIRM_ROW].map(function (r) {
      return sh.getRange(r, V.INFO_VALUE_COL).getValue();
    }),
    must: sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 2).getValues(),
    habits: sh.getRange(L.HABIT_START, V.HABIT_DAY_COL, CONFIG.HABIT_COUNT, 7).getValues(),
    days: [],
  };
  for (let d = 0; d < 7; d++) {
    data.days.push({
      checks: sh.getRange(L.TODO_START, V.checkCol(d), n, 1).getValues(),
      texts: sh.getRange(L.TODO_START, V.textCol(d), n, 1).getValues(),
      mood: sh.getRange(L.MOOD, V.checkCol(d), 2, 1).getValues(),
    });
  }
  return data;
}

/** 읽어 둔 입력값을 새(3열) 위치에 쓰기 */
function restoreWeekly_(sh, data) {
  const L = WEEKLY, n = CONFIG.TODO_ROWS;
  if (data.weekStart) sh.getRange(L.WEEK_START.row, L.WEEK_START.col).setValue(data.weekStart);
  [L.GOAL_ROW, L.REWARD_ROW, L.AFFIRM_ROW].forEach(function (r, i) {
    sh.getRange(r, L.INFO_VALUE_COL).setValue(data.info[i]);
  });
  sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 2).setValues(data.must);
  sh.getRange(L.HABIT_START, L.HABIT_DAY_COL, CONFIG.HABIT_COUNT, 7).setValues(data.habits);
  data.days.forEach(function (day, d) {
    sh.getRange(L.TODO_START, L.checkCol(d), n, 1).setValues(day.checks);
    sh.getRange(L.TODO_START, L.textCol(d), n, 1).setValues(day.texts);
    sh.getRange(L.MOOD, L.checkCol(d), 2, 1).setValues(day.mood);
  });
}

/* ---------- 데이터 읽기·쓰기 ---------- */

/** 위클리 시트가 보여주는 주의 월요일 (없으면 null) */
/** weeklyMonday_: 실행 한 번 동안 기억 */
function weeklyMonday_() { return memo_('weeklyMonday', weeklyMondayRaw_); }

function weeklyMondayRaw_() {
  const sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const d = dateOnly_(sh.getRange(WEEKLY.WEEK_START.row, WEEKLY.WEEK_START.col).getValue());
  return d ? mondayOf_(d) : null;
}

/** 날짜가 위클리의 몇 번째 요일인지 (0=월, 범위 밖이면 -1) */
function weeklyIndexOf_(date) {
  const mon = weeklyMonday_();
  if (!mon || !date) return -1;
  const d = diffDays_(mon, date);
  return d >= 0 && d <= 6 ? d : -1;
}

/** 위클리 수정 시: 수동 실행 / 할 일 정렬 / 오늘 시트 반영 / 기록 즉시 저장 */
function onWeeklyEdit_(e) {
  const L = WEEKLY, sh = e.range.getSheet();
  const out = function (cell) { return { row: cell.row, col: L.ACTION_RESULT_COL }; };

  if (justChecked_(e, L.ACTION_SAVE.row, L.ACTION_SAVE.col)) {
    return runAction_(sh, L.ACTION_SAVE, out(L.ACTION_SAVE), function () {
      return '저장됨 ✓ ' + saveAll_() + '행';
    });
  }
  if (justChecked_(e, L.ACTION_SYNC.row, L.ACTION_SYNC.col)) {
    return runAction_(sh, L.ACTION_SYNC, out(L.ACTION_SYNC), function () {
      return '불러옴 ✓ ' + syncCalendar_() + '개';
    });
  }
  // 주 시작을 바꾸면 그 주를 저장소에서 다시 그림 (데이터는 그대로)
  if (hit_(e.range, L.WEEK_START.row, L.WEEK_START.col, 1, 3)) {
    renderWeekView_();
    refreshToday_();
    return;
  }

  const mon = weeklyMonday_();
  if (!mon) return;
  const todayKey = dateKey_(todaySheetDate_() || today_());
  let touchedToday = false;
  const todayParts = {};

  for (let d = 0; d < 7; d++) {
    const date = addDays_(mon, d);
    const inTodo = hit_(e.range, L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, L.DAY_SPAN) ||
      hit_(e.range, L.TODO_START, L.memoCol(d), CONFIG.TODO_ROWS, 1);
    const inMood = hit_(e.range, L.MOOD, L.checkCol(d), 2, L.DAY_SPAN);
    const inHabit = hit_(e.range, L.HABIT_START, L.HABIT_DAY_COL + d, CONFIG.HABIT_COUNT, 1);
    if (!inTodo && !inMood && !inHabit) continue;
    if (inTodo) syncTodosFromSheet_(sh, L.TODO_START, L.checkCol(d), L.memoCol(d), date);
    if (inMood) setMoodData_(date, moodScore_(sh.getRange(L.MOOD, L.checkCol(d)).getValue()), sh.getRange(L.MOOD_MEMO, L.checkCol(d)).getValue());
    if (inHabit) {
      const checks = sh.getRange(L.HABIT_START, L.HABIT_DAY_COL + d, CONFIG.HABIT_COUNT, 1).getValues();
      const map = {};
      getHabits_().forEach(function (n, i) { if (n) map[n] = checks[i][0] === true; });
      setHabitsDone_(date, map);
    }
    archiveDayLive_(date);
    // 할 일은 정렬이 바뀔 수 있어 다시 그림, 기분·습관은 사용자가 쓴 그대로라 위클리는 안 다시 그림
    if (inTodo) renderWeekView_({ days: [d], todos: true });
    if (dateKey_(date) === todayKey) {
      touchedToday = true;
      todayParts.todos = todayParts.todos || inTodo;
      todayParts.mood = todayParts.mood || inMood;
      todayParts.habits = todayParts.habits || inHabit;
    }
  }
  if (hit_(e.range, L.MUST_START, 1, CONFIG.MUST_COUNT, L.MUST_WIDTH)) {
    syncMustFromSheet_(sh, mon, L.MUST_START, 1, 2);
    markWeekGoalsDirty_(mon);
    renderWeekView_({ days: [], must: true });
    if (weeklyIndexOf_(todaySheetDate_() || today_()) >= 0) { touchedToday = true; todayParts.must = true; }
  }
  if (hit_(e.range, L.GOAL_ROW, L.INFO_VALUE_COL, L.AFFIRM_ROW - L.GOAL_ROW + 1, 1)) {
    setWeekInfo_(mon, [L.GOAL_ROW, L.REWARD_ROW, L.AFFIRM_ROW].map(function (r) { return sh.getRange(r, L.INFO_VALUE_COL).getValue(); }));
  }
  if (touchedToday) renderTodayView_(todayParts);
}

/**
 * 시트의 할 일 10줄 → 저장소 (row/col = 체크박스 첫 칸, memoCol = 같은 줄 메모 칸)
 * 보이는 줄 i = 저장소 i번째 할 일. 내용을 지우면 그 할 일(과 메모) 삭제, 빈 줄에 적으면 추가.
 * 10개 넘게 있어서 안 보이는 할 일은 그대로 둠
 */
function syncTodosFromSheet_(sh, row, col, memoCol, date) {
  const n = CONFIG.TODO_ROWS;
  const items = getTodos_(date);
  const block = sh.getRange(row, col, n, 3);
  const vals = block.getValues(), disp = block.getDisplayValues();
  const memos = memoRange_(sh, row, memoCol).getValues();
  const shown = Math.min(items.length, n);
  const next = [];
  for (let i = 0; i < n; i++) {
    const text = String(vals[i][2]).trim();
    if (!text) continue; // 지운 줄 → 그 할 일과 메모 삭제
    const r = splitRange_(disp[i][1]);
    next.push({ done: vals[i][0] === true, start: r.start, end: r.end, text: text, memo: String(memos[i][0]),
      flag: i < shown ? items[i].flag : '' });
  }
  setTodos_(date, next.concat(items.slice(shown)));
}

/** 시트의 꼭 끝낼 일 5줄 → 저장소 (안 보이는 6번째부터는 그대로) */
function syncMustFromSheet_(sh, monday, row, checkCol, textCol) {
  const n = CONFIG.MUST_COUNT;
  const list = getWeekData_(monday).must;
  const shown = Math.min(list.length, n);
  const checks = sh.getRange(row, checkCol, n, 1).getValues(), texts = sh.getRange(row, textCol, n, 1).getValues();
  const next = [];
  for (let i = 0; i < n; i++) {
    const text = String(texts[i][0]).trim();
    if (text) next.push({ done: checks[i][0] === true, text: text });
  }
  setWeekMust_(monday, next.concat(list.slice(shown)));
}

/** 완료율 글자 (○◔◑◕● + %) */
function rateText_(done, total) {
  const p = total ? done / total : 0;
  return ['○', '◔', '◑', '◕', '●'][Math.round(p * 4)] + '  ' + Math.round(p * 100) + '%';
}

/**
 * 저장소 → 위클리 시트 (위클리가 보여 주는 주)
 * parts = {days: [요일 번호…], habits, must, info} — 없으면 전부 다시 그림
 * 할 일은 10줄, 꼭 끝낼 일은 5줄까지 보이고 넘치면 "+N개 더", 완료율은 전체 기준
 */
function renderWeekView_(parts) {
  const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY), n = CONFIG.TODO_ROWS;
  const mon = weeklyMonday_();
  if (!mon) return;
  const all = !parts;
  const P = parts || {};
  const days = all ? [0, 1, 2, 3, 4, 5, 6] : P.days || [];
  // 요일만 주면 할 일·기분을 그림 (예전 호출 방식)
  const doTodos = all || P.todos || (days.length && P.todos === undefined && P.mood === undefined && P.habits === undefined);
  const doMood = all || P.mood || (days.length && P.todos === undefined && P.mood === undefined && P.habits === undefined);
  const doHabits = all || P.habits;
  const needTodos = doTodos && days.length, needDay = (doMood && days.length) || doHabits;
  const todos = needTodos ? todosByDate_(mon, addDays_(mon, 6)) : {};
  const dayData = needDay ? dayDataByDate_(mon, addDays_(mon, 6)) : {};
  const habits = doHabits ? getHabits_() : [];
  const first = L.checkCol(0), width = L.LAST_DAY_COL - first + 1;

  // 하루치 값: 할 일 10줄 + 합계 줄(3열), 메모, 완료율, 차트(완료·남음)
  const dayBlock = function (d) {
    const list = todos[dateKey_(addDays_(mon, d))] || [], rows = [], memos = [];
    for (let i = 0; i < n; i++) {
      const t = list[i];
      rows.push(t ? [t.done, todoTime_(t), todoDisplay_(t)] : [false, '', '']);
      memos.push(t ? t.memo : '');
    }
    const done = list.filter(function (t) { return t.done; }).length;
    rows.push([done + ' 완료 · ' + (list.length - done) + ' 남음' + (list.length > n ? ' · +' + (list.length - n) + '개 더' : ''), '', '']);
    return { rows: rows, memos: memos, rate: rateText_(done, list.length), chart: [done, list.length - done] };
  };
  const moodPair = function (d) {
    const day = dayData[dateKey_(addDays_(mon, d))] || emptyDay_();
    return [day.mood.score ? CONFIG.MOODS[day.mood.score - 1] : '', day.mood.memo];
  };

  if (doTodos && days.length === 7) {
    // 7일을 한 번에: 할 일 영역·메모·완료율·차트를 각각 setValues 1번씩
    const blocks = days.map(dayBlock);
    const grid = [], rate = [], memo = [];
    for (let r = 0; r <= n; r++) grid.push([].concat.apply([], blocks.map(function (b) { return b.rows[r]; })));
    rate.push([].concat.apply([], blocks.map(function (b) { return [b.rate, '', '']; })));
    for (let r = 0; r < n; r++) memo.push(blocks.map(function (b) { return b.memos[r]; }));
    sh.getRange(L.TODO_START, first, n + 1, width).setValues(grid);
    sh.getRange(L.RATE, first, 1, width).setValues(rate);
    memoRange_(sh, L.TODO_START, L.MEMO_COL0 + 6);
    sh.getRange(L.TODO_START, L.MEMO_COL0, n, 7).setValues(memo);
    sh.getRange(L.DAY_HEADER + 1, L.CHART_DATA_COL + 1, 7, 2).setValues(blocks.map(function (b) { return b.chart; }));
  } else if (doTodos) {
    days.forEach(function (d) {
      const b = dayBlock(d);
      sh.getRange(L.TODO_START, L.checkCol(d), n + 1, 3).setValues(b.rows);
      memoRange_(sh, L.TODO_START, L.memoCol(d)).setValues(b.memos.map(function (m) { return [m]; }));
      sh.getRange(L.RATE, L.checkCol(d)).setValue(b.rate);
      sh.getRange(L.DAY_HEADER + 1 + d, L.CHART_DATA_COL + 1, 1, 2).setValues([b.chart]);
    });
  }
  if (doMood && days.length === 7) {
    const row1 = [], row2 = [];
    days.forEach(function (d) { const p = moodPair(d); row1.push(p[0], '', ''); row2.push(p[1], '', ''); });
    sh.getRange(L.MOOD, first, 2, width).setValues([row1, row2]);
  } else if (doMood) {
    days.forEach(function (d) { const p = moodPair(d); sh.getRange(L.MOOD, L.checkCol(d), 2, 1).setValues([[p[0]], [p[1]]]); });
  }
  if (doHabits) {
    // 습관은 7일 × 10개를 한 번에 (바뀐 날만이 아니라도 1번이면 충분)
    sh.getRange(L.HABIT_START, L.HABIT_DAY_COL, CONFIG.HABIT_COUNT, 7).setValues(habits.map(function (h) {
      const row = [];
      for (let d = 0; d < 7; d++) {
        const day = dayData[dateKey_(addDays_(mon, d))];
        row.push(!!(h && day && day.habits[h]));
      }
      return row;
    }));
  }
  if (all || P.must || P.info) {
    const week = getWeekData_(mon), m = CONFIG.MUST_COUNT;
    if (all || P.must) {
      const rows = [];
      for (let i = 0; i < m; i++) rows.push(week.must[i] ? [week.must[i].done, week.must[i].text] : [false, '']);
      sh.getRange(L.MUST_START, 1, m, 2).setValues(rows);
      const done = week.must.filter(function (x) { return x.done; }).length, total = week.must.length;
      const p = total ? done / total : 0;
      sh.getRange(L.MUST_PROGRESS, 2).setFormula('=SPARKLINE(' + p + ',{"charttype","bar";"max",1;"color1","' + getTheme_() + '"})');
      sh.getRange(L.MUST_PROGRESS, L.MUST_WIDTH - 1).setValue(Math.round(p * 100) + '%' + (total > m ? ' · +' + (total - m) + '개 더' : ''));
    }
    if (all || P.info) sh.getRange(L.GOAL_ROW, L.INFO_VALUE_COL, 3, 1).setValues(week.info.map(function (v) { return [v]; }));
  }
}

/** 체크박스 실행 공통: 체크 해제 → 실행 → 결과를 result {row, col} 셀에 글자로 */
function runAction_(sh, cell, result, fn) {
  sh.getRange(cell.row, cell.col).setValue(false);
  const out = sh.getRange(result.row, result.col);
  out.setValue('실행 중…');
  SpreadsheetApp.flush();
  try {
    out.setValue(fn() + '  ' + stamp_());
  } catch (err) {
    out.setValue('오류: ' + err.message);
  }
}

/**
 * 새 주 시작: 지난주 기록 저장 → 저장 확인 → 날짜를 이번 주로 (데이터는 저장소에 남음)
 * 위클리가 이미 이번 주면 아무것도 안 함 (여러 번 실행해도 안전)
 */
function startNewWeek_() {
  const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const cur = mondayOf_(today_());
  const mon = weeklyMonday_();
  if (mon && mon >= cur) return '이미 이번 주예요';

  if (mon) {
    // 1) 지난주 7일 전체 + 꼭 끝낼 일 기록 저장 (저장소 데이터는 지우지 않음)
    const n = saveRange_(mon, addDays_(mon, 6));
    archiveWeekGoals_(mon);
    // 2) 저장 확인: 내용이 있었는데 기록에 없으면 중단
    if (n > 0 && countArchiveWeek_(dateKey_(mon)) === 0) throw new Error('지난주 기록 저장을 확인하지 못해 새 주 전환을 멈췄어요');
    sheet_(CONFIG.SHEETS.LASTWEEK).getRange(LASTWEEK.SELECT.row, LASTWEEK.SELECT.col).setValue(dateKey_(mon));
  }
  // 3) 날짜를 이번 주 월요일로 → 저장소에서 이번 주를 그림 (비우기 필요 없음)
  sh.getRange(L.WEEK_START.row, L.WEEK_START.col).setValue(cur);
  forget_('weeklyMonday');
  renderWeekView_();
  return '새 주 시작 ✓ ' + dateKey_(cur);
}
