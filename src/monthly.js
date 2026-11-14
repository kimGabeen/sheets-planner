/**
 * monthly.js
 * 먼슬리 시트: 연/월 드롭다운 달력, 이번 달 목표, 월간 습관 결과, 일별 차트
 * - 달력 칸 = 날짜+기분 이모지 / 캘린더 일정(자동) / 내 메모(직접 입력)
 * - 습관·기분은 여기서 입력하지 않고 기록 시트 결과만 보여줌
 */

/** 달력 첫 칸(월요일) 날짜 수식 */
function monthStartExpr_() {
  const y = a1_(MONTHLY.YEAR.row, MONTHLY.YEAR.col, true);
  const m = a1_(MONTHLY.MONTH.row, MONTHLY.MONTH.col, true);
  return '(DATE(' + y + ',' + m + ',1)-WEEKDAY(DATE(' + y + ',' + m + ',1),3))';
}

/** 먼슬리 시트 레이아웃·수식·서식 그리기 (입력 값은 유지) */
function buildMonthly_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.MONTHLY);
  const L = MONTHLY, C = CONFIG.COLORS;
  ensureSize_(sh, L.CHART_ROW + 20, L.LAST_COL);
  resetLayout_(sh);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(10).setVerticalAlignment('middle');
  sh.setColumnWidths(1, L.CHART_DATA_COL - 1, CONFIG.SIZE.GRID_COL);
  sh.setRowHeights(1, L.CHART_ROW, 26);

  buildMonthlyCalendar_(sh, theme);
  buildMonthlyPanel_(sh, theme);
  buildMonthlyHabits_(sh, theme);
  buildMonthlyCharts_(sh, theme);
}

/** 상단 드롭다운 + 6주 달력 */
function buildMonthlyCalendar_(sh, theme) {
  const L = MONTHLY, C = CONFIG.COLORS;
  const y = a1_(L.YEAR.row, L.YEAR.col, true), m = a1_(L.MONTH.row, L.MONTH.col, true);
  const start = monthStartExpr_();
  const arc = CONFIG.SHEETS.ARCHIVE, ev = CONFIG.SHEETS.EVENTS;
  const emoji = CONFIG.MOOD_EMOJI.map(function (x) { return '"' + x + '"'; }).join(',');

  // 연도·월 드롭다운
  const thisYear = new Date().getFullYear();
  const years = [];
  for (let i = -2; i <= 3; i++) years.push(String(thisYear + i));
  sh.getRange(1, 1, 1, 2).merge().setValue('연도').setFontColor(C.SUB).setHorizontalAlignment('right');
  sh.getRange(L.YEAR.row, L.YEAR.col, 1, 2).merge().setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(years, true).build()).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, 5, 1, 2).merge().setValue('월').setFontColor(C.SUB).setHorizontalAlignment('right');
  sh.getRange(L.MONTH.row, L.MONTH.col, 1, 2).merge().setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'], true).build())
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, L.TITLE_COL, 1, 28 - L.TITLE_COL + 1).merge()
    .setFormula('=IFERROR(TEXT(DATE(' + y + ',' + m + ',1),"yyyy년 m월"),"")')
    .setFontSize(18).setFontWeight('bold').setFontColor(theme).setHorizontalAlignment('right');
  sh.setRowHeight(1, 36);

  // 요일 머리글
  ['월', '화', '수', '목', '금', '토', '일'].forEach(function (n, d) {
    styleHeader_(sh.getRange(L.WEEKDAY_ROW, L.dayCol(d), 1, L.DAY_SPAN).merge().setValue(n)
      .setHorizontalAlignment('center'), theme);
  });

  // 날짜 칸
  for (let w = 0; w < L.WEEKS; w++) {
    sh.setRowHeight(L.dateRow(w), 22).setRowHeight(L.eventRow(w), 56).setRowHeight(L.memoRow(w), 40);
    for (let d = 0; d < 7; d++) {
      const c = L.dayCol(d);
      const dt = start + '+' + (7 * w + d);
      // 1) 날짜 + 기분 이모지 (기록 시트에서)
      sh.getRange(L.dateRow(w), c, 1, L.DAY_SPAN).merge().setFormula(
        '=IFERROR(LET(dt,' + dt + ',s,IFERROR(INDEX(FILTER(' + colRef_(arc, 5) + ',' + colRef_(arc, 1) +
        '=dt,' + colRef_(arc, 2) + '="' + CONFIG.CATEGORY.MOOD + '"),1),""),' +
        'DAY(dt)&IF(s="",""," "&CHOOSE(s,' + emoji + '))),"")')
        .setFontWeight('bold').setFontSize(10);
      // 2) 구글 캘린더 일정 (자동)
      sh.getRange(L.eventRow(w), c, 1, L.DAY_SPAN).merge().setFormula(
        '=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER(' + colRef_(ev, 3) + ',' + colRef_(ev, 1) + '=' + dt + ',' +
        colRef_(ev, 2) + '="' + CONFIG.EVENT_TYPE.CAL + '")),"")')
        .setFontSize(8).setFontColor(shade_(theme, 0.1)).setWrap(true).setVerticalAlignment('top');
      // 3) 내 메모 (직접 입력, 스크립트가 날짜별로 저장)
      underline_(sh.getRange(L.memoRow(w), c, 1, L.DAY_SPAN).merge().setFontSize(8).setWrap(true)
        .setVerticalAlignment('top'));
    }
  }

  // 조건부 서식: 다른 달 날짜는 회색, 오늘은 연한 테마색
  const cellDate = start + '+7*INT((ROW()-' + L.CAL_START + ')/3)+INT((COLUMN()-1)/' + L.DAY_SPAN + ')';
  const calRange = sh.getRange(L.CAL_START, 1, L.WEEKS * 3, 7 * L.DAY_SPAN);
  const dateRows = [];
  for (let w = 0; w < L.WEEKS; w++) dateRows.push(sh.getRange(L.dateRow(w), 1, 1, 7 * L.DAY_SPAN));
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + cellDate + '=TODAY()')
      .setBackground(tint_(theme, CONFIG.TINT.LIGHT)).setRanges(dateRows).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=MONTH(' + cellDate + ')<>VALUE(' + m + ')')
      .setFontColor('#C8C8C8').setRanges([calRange]).build(),
  ]);
}

/** 오른쪽 패널: 이번 달 목표, 새 일정 추가, 캘린더 새로고침 */
function buildMonthlyPanel_(sh, theme) {
  const L = MONTHLY, C = CONFIG.COLORS, P = L.PANEL_COL, W = L.PANEL_END - P + 1;

  styleHeader_(sh.getRange(L.GOAL_HEADER, P, 1, W).merge().setValue('이번 달 목표'), theme);
  for (let i = 0; i < CONFIG.MONTH_GOALS; i++) {
    checkbox_(sh.getRange(L.GOAL_START + i, P));
    underline_(sh.getRange(L.GOAL_START + i, P + 1, 1, W - 1).merge());
  }

  styleHeader_(sh.getRange(L.ADD_HEADER, P, 1, W).merge().setValue('새 일정 추가 (구글 캘린더)'), theme);
  const labels = { date: '날짜', start: '시작', end: '끝(선택)', title: '제목', check: '추가' };
  Object.keys(L.ADD).forEach(function (k) {
    sh.getRange(L.ADD_LABEL, L.ADD[k], 1, L.ADD_SPAN[k]).merge().setValue(labels[k])
      .setFontColor(C.SUB).setFontSize(8).setHorizontalAlignment('center');
    const input = sh.getRange(L.ADD_ROW, L.ADD[k], 1, L.ADD_SPAN[k]).merge().setHorizontalAlignment('center');
    if (k !== 'check') underline_(input);
  });
  dateInput_(sh.getRange(L.ADD_ROW, L.ADD.date, 1, L.ADD_SPAN.date));
  sh.getRange(L.ADD_ROW, L.ADD.start, 1, 4).setNumberFormat('HH:mm');
  checkbox_(sh.getRange(L.ADD_ROW, L.ADD.check));
  sh.getRange(L.ADD_RESULT, P, 1, W).merge().setFontColor(C.SUB).setFontSize(9)
    .setNote('시작+끝: 그 시간대 일정 / 시작만: 종일 + 제목 앞에 시간 / 둘 다 없음: 종일');
  sh.setRowHeight(L.ADD_ROW, 30);

  checkbox_(sh.getRange(L.SYNC_ROW, P));
  sh.getRange(L.SYNC_ROW, P + 1, 1, 4).merge().setValue('캘린더 새로고침');
  sh.getRange(L.SYNC_ROW, P + 5, 1, W - 5).merge().setFontColor(C.SUB).setFontSize(9);
}

/** 월간 습관 트래커 (기록 결과만 표시, 직접 체크 X) */
function buildMonthlyHabits_(sh, theme) {
  const L = MONTHLY, C = CONFIG.COLORS, arc = CONFIG.SHEETS.ARCHIVE;
  const y = a1_(L.YEAR.row, L.YEAR.col, true), m = a1_(L.MONTH.row, L.MONTH.col, true);
  const days = 'DAY(EOMONTH(DATE(' + y + ',' + m + ',1),0))';
  const H = L.HABIT_HEADER;

  styleHeader_(sh.getRange(H, 1, 1, L.HABIT_NAME_SPAN).merge().setValue('습관 (자동)'), theme);
  for (let n = 1; n <= 31; n++) {
    styleHeader_(sh.getRange(H, L.HABIT_DAY_COL + n - 1).setFormula('=IF(' + n + '>' + days + ',"",' + n + ')')
      .setHorizontalAlignment('center').setFontSize(8), theme);
  }
  styleHeader_(sh.getRange(H, L.HABIT_BAR_COL, 1, 3).merge().setValue('달성률')
    .setHorizontalAlignment('center'), theme);

  // 경과 일수 (이번 달이면 오늘까지, 지난 달이면 전체)
  const elapsed = 'LET(f,DATE(' + y + ',' + m + ',1),l,EOMONTH(f,0),IF(TODAY()<f,0,MIN(l,TODAY())-f+1))';
  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    const r = L.HABIT_START + i, name = a1_(r, 1, false).replace(/^A/, '$A');
    sh.getRange(r, 1, 1, L.HABIT_NAME_SPAN).merge().setFormula('=' + habitRef_(i) + '&""').setFontSize(9);
    const formulas = [];
    for (let n = 1; n <= 31; n++) {
      formulas.push('=IF(OR(' + name + '="",' + n + '>' + days + '),"",IF(COUNTIFS(' +
        colRef_(arc, 1) + ',DATE(' + y + ',' + m + ',' + n + '),' + colRef_(arc, 2) + ',"' + CONFIG.CATEGORY.HABIT +
        '",' + colRef_(arc, 3) + ',' + name + ',' + colRef_(arc, 4) + ',TRUE)>0,"●",""))');
    }
    sh.getRange(r, L.HABIT_DAY_COL, 1, 31).setFormulas([formulas])
      .setFontColor(theme).setHorizontalAlignment('center').setFontSize(9);
    const row = a1_(r, L.HABIT_DAY_COL) + ':' + a1_(r, L.HABIT_DAY_COL + 30);
    const p = 'COUNTIF(' + row + ',"●")/' + elapsed;
    sh.getRange(r, L.HABIT_BAR_COL, 1, 2).merge().setFormula('=IFERROR(IF(' + name + '="","",SPARKLINE(' + p +
      ',{"charttype","bar";"max",1;"color1","' + theme + '"})),"")');
    sh.getRange(r, L.HABIT_PCT_COL).setFormula('=IFERROR(IF(' + name + '="","",' + p + '),"")')
      .setNumberFormat('0%').setFontSize(8).setFontColor(C.SUB);
    underline_(sh.getRange(r, 1, 1, L.HABIT_PCT_COL));
  }
}

/** 일별 할 일 달성률 막대 + 기분 꺾은선 차트 */
function buildMonthlyCharts_(sh, theme) {
  const L = MONTHLY, arc = CONFIG.SHEETS.ARCHIVE, dc = L.CHART_DATA_COL;
  const y = a1_(L.YEAR.row, L.YEAR.col, true), m = a1_(L.MONTH.row, L.MONTH.col, true);
  const days = 'DAY(EOMONTH(DATE(' + y + ',' + m + ',1),0))';
  const A = colRef_(arc, 1), B = colRef_(arc, 2), D = colRef_(arc, 4), E = colRef_(arc, 5);

  // 숨김 데이터: 일 | 할 일 % | 기분
  const rows = [['일', '할 일 달성률', '기분']];
  for (let n = 1; n <= 31; n++) {
    const dt = 'DATE(' + y + ',' + m + ',' + n + ')';
    rows.push([
      String(n),
      '=IF(' + n + '>' + days + ',"",LET(t,COUNTIFS(' + A + ',' + dt + ',' + B + ',"' + CONFIG.CATEGORY.TODO +
        '"),IF(t=0,"",COUNTIFS(' + A + ',' + dt + ',' + B + ',"' + CONFIG.CATEGORY.TODO + '",' + D + ',TRUE)/t)))',
      '=IF(' + n + '>' + days + ',"",IFERROR(INDEX(FILTER(' + E + ',' + A + '=' + dt + ',' + B + '="' +
        CONFIG.CATEGORY.MOOD + '"),1),""))',
    ]);
  }
  sh.getRange(1, dc, 32, 1).setNumberFormat('@');
  sh.getRange(1, dc, 32, 3).setValues(rows);
  sh.hideColumns(dc, 3);

  const bar = sh.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(1, dc, 32, 2))
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setPosition(L.CHART_ROW, 1, 0, 0)
    .setOption('title', '일별 할 일 달성률')
    .setOption('colors', [theme])
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { min: 0, max: 1 }, gridlines: { color: '#F2F2F2' } })
    .setOption('width', 640).setOption('height', 260)
    .build();
  sh.insertChart(bar);

  const line = sh.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(1, dc, 32, 1))
    .addRange(sh.getRange(1, dc + 2, 32, 1))
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setPosition(L.CHART_ROW, 21, 0, 0)
    .setOption('title', '일별 기분 점수')
    .setOption('colors', [theme])
    .setOption('legend', { position: 'none' })
    .setOption('pointSize', 6)
    .setOption('vAxis', { viewWindow: { min: 1, max: 5 }, gridlines: { count: 5, color: '#F2F2F2' } })
    .setOption('width', 640).setOption('height', 260)
    .build();
  sh.insertChart(line);
}

/* ---------- 데이터 ---------- */

/** 먼슬리에서 고른 연·월 ({y, m}, 없으면 null) */
/** getMonthlyYM_: 실행 한 번 동안 기억 */
function getMonthlyYM_() { return memo_('monthlyYM', getMonthlyYMRaw_); }

function getMonthlyYMRaw_() {
  const sh = sheet_(CONFIG.SHEETS.MONTHLY);
  if (!sh) return null;
  // 연도(C1)·월(G1)을 한 번에 읽기
  const v = sh.getRange(MONTHLY.YEAR.row, MONTHLY.YEAR.col, 1, MONTHLY.MONTH.col - MONTHLY.YEAR.col + 1).getValues()[0];
  const y = Number(v[0]), m = Number(v[MONTHLY.MONTH.col - MONTHLY.YEAR.col]);
  return y && m >= 1 && m <= 12 ? { y: y, m: m } : null;
}

/** 달력 첫 칸(월요일) 날짜 */
function monthGridStart_(ym) {
  const first = new Date(ym.y, ym.m - 1, 1);
  return mondayOf_(first);
}

/** 달력의 메모 칸 날짜들을 _일정 시트의 메모로 다시 채우기 */
function refreshMonthMemos_() {
  const sh = sheet_(CONFIG.SHEETS.MONTHLY), L = MONTHLY;
  const ym = getMonthlyYM_();
  if (!ym) return;
  const start = monthGridStart_(ym);
  const memos = {};
  const rows = readEvents_();
  const hidden = hiddenMemoTest_(rows); // 캘린더 줄에 이미 보이는 메모는 메모 칸에 또 쓰지 않음 (중복 표시 방지)
  rows.forEach(function (r) {
    if (r[1] !== CONFIG.EVENT_TYPE.MEMO || hidden(r)) return;
    const k = dateKey_(r[0]);
    memos[k] = memos[k] ? memos[k] + '\n' + r[2] : r[2]; // 한 줄에 일정 하나
  });
  for (let w = 0; w < L.WEEKS; w++) {
    const row = [];
    for (let d = 0; d < 7; d++) {
      const key = dateKey_(addDays_(start, 7 * w + d));
      row.push(memos[key] || '');
      for (let k = 1; k < L.DAY_SPAN; k++) row.push('');
    }
    sh.getRange(L.memoRow(w), 1, 1, row.length).setValues([row]);
  }
}

/** 수정된 메모 칸들을 _일정 시트에 저장 */
function saveEditedMemos_(range) {
  const L = MONTHLY, sh = range.getSheet();
  const ym = getMonthlyYM_();
  if (!ym) return;
  const start = monthGridStart_(ym);
  const entries = [];
  for (let w = 0; w < L.WEEKS; w++) {
    for (let d = 0; d < 7; d++) {
      if (!hit_(range, L.memoRow(w), L.dayCol(d), 1, L.DAY_SPAN)) continue;
      entries.push({
        date: addDays_(start, 7 * w + d),
        text: String(sh.getRange(L.memoRow(w), L.dayCol(d)).getValue()).trim(),
      });
    }
  }
  if (entries.length) saveMemos_(entries);
}

/** 이번 달 목표 칸을 기록 시트 값으로 채우기 */
function loadMonthGoals_() {
  const sh = sheet_(CONFIG.SHEETS.MONTHLY), L = MONTHLY;
  const ym = getMonthlyYM_();
  if (!ym) return;
  const key = dateKey_(new Date(ym.y, ym.m - 1, 1));
  const goals = readArchive_().filter(function (r) {
    return r.key === key && r.cat === CONFIG.CATEGORY.MONTH_GOAL;
  }).slice(0, CONFIG.MONTH_GOALS);
  const vals = [];
  for (let i = 0; i < CONFIG.MONTH_GOALS; i++) {
    vals.push(goals[i] ? [goals[i].done === true, goals[i].item] : [false, '']);
  }
  sh.getRange(L.GOAL_START, L.PANEL_COL, CONFIG.MONTH_GOALS, 1).setValues(vals.map(function (v) { return [v[0]]; }));
  sh.getRange(L.GOAL_START, L.PANEL_COL + 1, CONFIG.MONTH_GOALS, 1).setValues(vals.map(function (v) { return [v[1]]; }));
}

/** 연·월이 바뀌었을 때: 메모·목표 다시 채우고 그 달 일정 불러오기 */
function onMonthChange_() {
  refreshMonthMemos_();
  loadMonthGoals_();
  const sh = sheet_(CONFIG.SHEETS.MONTHLY), L = MONTHLY;
  const out = sh.getRange(L.SYNC_ROW, L.PANEL_COL + 5);
  try {
    out.setValue('불러옴 ✓ ' + syncCalendar_() + '개  ' + stamp_());
  } catch (err) {
    out.setValue('오류: ' + err.message);
  }
}

/** 먼슬리 수정 시 */
function onMonthlyEdit_(e) {
  const L = MONTHLY, sh = e.range.getSheet();
  if (hit_(e.range, L.YEAR.row, L.YEAR.col, 1, 2) || hit_(e.range, L.MONTH.row, L.MONTH.col, 1, 2)) {
    return onMonthChange_();
  }
  if (justChecked_(e, L.ADD_ROW, L.ADD.check)) {
    return addEventFromForm_(sh, {
      date: [L.ADD_ROW, L.ADD.date], start: [L.ADD_ROW, L.ADD.start], end: [L.ADD_ROW, L.ADD.end],
      title: [L.ADD_ROW, L.ADD.title], check: [L.ADD_ROW, L.ADD.check], result: [L.ADD_RESULT, L.PANEL_COL],
    });
  }
  if (justChecked_(e, L.SYNC_ROW, L.PANEL_COL)) {
    return runAction_(sh, { row: L.SYNC_ROW, col: L.PANEL_COL }, { row: L.SYNC_ROW, col: L.PANEL_COL + 5 }, function () {
      return '불러옴 ✓ ' + syncCalendar_() + '개';
    });
  }
  if (hit_(e.range, L.CAL_START, 1, L.WEEKS * 3, 7 * L.DAY_SPAN)) saveEditedMemos_(e.range);
  if (hit_(e.range, L.GOAL_START, L.PANEL_COL, CONFIG.MONTH_GOALS, L.PANEL_END - L.PANEL_COL + 1)) {
    archiveMonthGoals_();
  }
}
