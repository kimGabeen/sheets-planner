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
  if (old) restoreWeekly_(sh, old);
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
  const mc = 'A' + L.MUST_START + ':A' + (L.MUST_PROGRESS - 1);
  const mt = 'B' + L.MUST_START + ':B' + (L.MUST_PROGRESS - 1);
  const p = 'IF(COUNTIF(' + mt + ',"?*")=0,0,COUNTIFS(' + mc + ',TRUE,' + mt + ',"?*")/COUNTIF(' + mt + ',"?*"))';
  sh.getRange(L.MUST_PROGRESS, 1).setValue('진행률').setFontColor(C.SUB);
  sh.getRange(L.MUST_PROGRESS, 2, 1, MW - 3).merge()
    .setFormula('=SPARKLINE(' + p + ',{"charttype","bar";"max",1;"color1","' + theme + '"})');
  sh.getRange(L.MUST_PROGRESS, MW - 1, 1, 2).merge().setFormula('=' + p).setNumberFormat('0%')
    .setHorizontalAlignment('right').setFontWeight('bold');
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

    // 2) 그날 완료율 (○◔◑◕● 원형 + %)
    sh.getRange(L.RATE, c, 1, S).merge().setFormula(
      '=LET(done,' + done + ',tot,' + tot + ',p,IF(tot=0,0,done/tot),' +
      'CHOOSE(ROUND(p*4)+1,"○","◔","◑","◕","●")&"  "&TEXT(p,"0%"))')
      .setHorizontalAlignment('center').setFontSize(13).setFontColor(theme).setFontWeight('bold');

    // 3) 일정 (먼슬리와 같은 데이터, 읽기 전용 회색)
    for (let k = 1; k <= CONFIG.SCHEDULE_LINES; k++) {
      sh.getRange(L.SCHED_START + k - 1, c, 1, S).merge().setFormula(
        '=IFERROR(INDEX(FILTER(' + colRef_(ev, 3) + ',' + colRef_(ev, 1) + '=' + date + '),' + k + '),"")')
        .setFontColor(C.SUB).setFontSize(9).setWrap(false);
    }

    // 4) 할 일 10줄 (체크박스 + 시간(선택) + 내용) — 시간은 글자로 저장해 정렬
    checkbox_(sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, 1));
    sh.getRange(L.TODO_START, L.timeCol(d), CONFIG.TODO_ROWS, 1).setNumberFormat('@')
      .setHorizontalAlignment('center').setFontSize(9).setFontColor(theme);
    underline_(sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, S));

    // 5) 완료 / 남은 개수
    sh.getRange(L.COUNT, c, 1, S).merge()
      .setFormula('=LET(done,' + done + ',tot,' + tot + ',done&" 완료 · "&(tot-done)&" 남음")')
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
    sh.getRange(top + 1 + d, dc, 1, 3).setFormulas([[
      '=TEXT(' + a1_(L.DAY_HEADER, c) + ',"ddd")',
      '=COUNTIFS(' + chk + ',TRUE,' + txt + ',"?*")',
      '=COUNTIF(' + txt + ',"?*")-' + a1_(top + 1 + d, dc + 1),
    ]]);
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
function weeklyMonday_() {
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
  if (hit_(e.range, L.WEEK_START.row, L.WEEK_START.col, 1, 3)) {
    refreshToday_();
    return;
  }

  const mon = weeklyMonday_();
  if (!mon) return;
  const todayIdx = weeklyIndexOf_(todaySheetDate_());
  let touchedToday = false;

  for (let d = 0; d < 7; d++) {
    const inTodo = hit_(e.range, L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, L.DAY_SPAN);
    const inMood = hit_(e.range, L.MOOD, L.checkCol(d), 2, L.DAY_SPAN);
    const inHabit = hit_(e.range, L.HABIT_START, L.HABIT_DAY_COL + d, CONFIG.HABIT_COUNT, 1);
    if (inTodo) sortTodos_(sh, L.TODO_START, L.checkCol(d)); // 미완료(시간순) → 체크된 일
    if (inTodo || inMood || inHabit) {
      archiveDayLive_(addDays_(mon, d));
      if (d === todayIdx) touchedToday = true;
    }
  }
  if (hit_(e.range, L.MUST_START, 1, CONFIG.MUST_COUNT, L.MUST_WIDTH)) {
    archiveWeekGoals_();
    if (todayIdx >= 0) touchedToday = true;
  }
  // 정렬로 줄이 움직였을 수 있어 오늘 칸 전체를 다시 복사
  if (touchedToday) copyBlocks_(sh, sheet_(CONFIG.SHEETS.TODAY), todayBlocks_(todayIdx), 'w');
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
 * 새 주 시작: 지난주 기록 저장 → 저장 확인 → 할 일·체크·기분 비우기 → 날짜를 이번 주로
 * 위클리가 이미 이번 주면 아무것도 안 함 (여러 번 실행해도 안전)
 */
function startNewWeek_() {
  const L = WEEKLY, sh = sheet_(CONFIG.SHEETS.WEEKLY);
  const cur = mondayOf_(today_());
  const mon = weeklyMonday_();
  if (mon && mon >= cur) return '이미 이번 주예요';

  if (mon) {
    // 1) 지난주 7일 전체 저장
    const n = saveWeek_(addDays_(mon, 6));
    // 2) 저장 확인: 내용이 있었는데 기록에 없으면 비우지 않고 중단
    const saved = countArchiveWeek_(dateKey_(mon));
    if (n > 0 && saved === 0) throw new Error('지난주 기록 저장을 확인하지 못해 초기화를 멈췄어요');
    sheet_(CONFIG.SHEETS.LASTWEEK).getRange(LASTWEEK.SELECT.row, LASTWEEK.SELECT.col).setValue(dateKey_(mon));
  }

  // 3) 비우기 (체크박스는 FALSE, 글자는 지우기)
  for (let d = 0; d < 7; d++) {
    sh.getRange(L.TODO_START, L.checkCol(d), CONFIG.TODO_ROWS, 1).setValue(false);
    sh.getRange(L.TODO_START, L.timeCol(d), CONFIG.TODO_ROWS, 2).clearContent();
    sh.getRange(L.MOOD, L.checkCol(d), 2, 1).clearContent();
  }
  sh.getRange(L.HABIT_START, L.HABIT_DAY_COL, CONFIG.HABIT_COUNT, 7).setValue(false);
  sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 1).setValue(false);
  sh.getRange(L.MUST_START, 2, CONFIG.MUST_COUNT, 1).clearContent();

  // 4) 날짜를 이번 주 월요일로
  sh.getRange(L.WEEK_START.row, L.WEEK_START.col).setValue(cur);
  SpreadsheetApp.flush();
  return '새 주 시작 ✓ ' + dateKey_(cur);
}
