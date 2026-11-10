/**
 * today.js
 * 오늘 시트: 폰 세로 화면용 (3열만 사용, 세로로 쭉)
 * 위클리의 "오늘 요일" 칸과 양방향으로 맞춰진다.
 */

/** 오늘 시트 레이아웃·수식·서식 그리기 (입력 값은 유지) */
function buildToday_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.TODAY);
  const T = TODAY, C = CONFIG.COLORS;
  const pale = tint_(theme, CONFIG.TINT.PALE);
  const ev = CONFIG.SHEETS.EVENTS;
  ensureSize_(sh, T.LAST_ROW + 2, 3);
  resetLayout_(sh);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), 3).setFontSize(11).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 52).setColumnWidth(2, 220).setColumnWidth(3, 100);
  sh.setRowHeights(1, T.LAST_ROW + 1, CONFIG.SIZE.TODAY_ROW);
  sh.setRowHeight(T.DATE_ROW, 48);
  sh.setFrozenRows(1);

  const header = function (row, text) {
    styleHeader_(sh.getRange(row, 1, 1, 3), theme);
    sh.getRange(row, 1, 1, 2).merge().setValue(text);
    return sh.getRange(row, 3).setHorizontalAlignment('right');
  };
  const date = a1_(T.DATE_ROW, 1, true);

  // 오늘 날짜 (자정에 스크립트가 갱신)
  sh.getRange(T.DATE_ROW, 1, 1, 3).merge().setNumberFormat('m월 d일 dddd')
    .setFontSize(18).setFontWeight('bold').setFontColor(theme);

  // 오늘 일정 (자동, 읽기 전용)
  header(T.SCHED_HEADER, '오늘 일정');
  for (let k = 1; k <= CONFIG.SCHEDULE_LINES; k++) {
    sh.getRange(T.SCHED_START + k - 1, 1, 1, 3).merge().setFormula(
      '=IFERROR(INDEX(FILTER(' + colRef_(ev, 3) + ',' + colRef_(ev, 1) + '=' + date + '),' + k + '),' +
      (k === 1 ? '"일정 없음"' : '""') + ')').setFontColor(C.SUB).setFontSize(10);
  }

  // 오늘 할 일 10줄
  const tc = 'A' + T.TODO_START + ':A' + (T.TODO_START + CONFIG.TODO_ROWS - 1);
  const tt = 'B' + T.TODO_START + ':B' + (T.TODO_START + CONFIG.TODO_ROWS - 1);
  header(T.TODO_HEADER, '오늘 할 일')
    .setFormula('=COUNTIFS(' + tc + ',TRUE,' + tt + ',"?*")&" / "&COUNTIF(' + tt + ',"?*")');
  for (let i = 0; i < CONFIG.TODO_ROWS; i++) {
    checkbox_(sh.getRange(T.TODO_START + i, 1));
    underline_(sh.getRange(T.TODO_START + i, 2, 1, 2).merge());
  }

  // 오늘 습관
  const hc = 'A' + T.HABIT_START + ':A' + (T.HABIT_START + CONFIG.HABIT_COUNT - 1);
  const hn = 'B' + T.HABIT_START + ':B' + (T.HABIT_START + CONFIG.HABIT_COUNT - 1);
  header(T.HABIT_HEADER, '오늘 습관')
    .setFormula('=COUNTIFS(' + hc + ',TRUE,' + hn + ',"?*")&" / "&COUNTIF(' + hn + ',"?*")');
  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    checkbox_(sh.getRange(T.HABIT_START + i, 1));
    underline_(sh.getRange(T.HABIT_START + i, 2, 1, 2).merge().setFormula('=' + habitRef_(i) + '&""'));
  }

  // 이번 주 꼭 끝낼 일
  header(T.MUST_HEADER, '이번 주 꼭 끝낼 일');
  for (let i = 0; i < CONFIG.MUST_COUNT; i++) {
    checkbox_(sh.getRange(T.MUST_START + i, 1));
    underline_(sh.getRange(T.MUST_START + i, 2, 1, 2).merge());
  }

  // 기분
  header(T.MOOD_HEADER, '오늘 기분');
  sh.getRange(T.MOOD, 1).setValue('점수').setFontColor(C.SUB).setFontSize(9);
  moodInput_(sh.getRange(T.MOOD, 2, 1, 2).merge());
  sh.getRange(T.MOOD_MEMO, 1).setValue('메모').setFontColor(C.SUB).setFontSize(9);
  underline_(sh.getRange(T.MOOD_MEMO, 2, 1, 2).merge());

  // 새 일정 추가 (세로 입력 폼)
  const A = T.ADD;
  header(T.ADD_HEADER, '새 일정 추가 (구글 캘린더)');
  [[A.date, '날짜'], [A.start, '시작'], [A.end, '끝'], [A.title, '제목'], [A.result, '결과']]
    .forEach(function (a) {
      sh.getRange(a[0], 1).setValue(a[1]).setFontColor(C.SUB).setFontSize(9);
      underline_(sh.getRange(a[0], 2, 1, 2).merge());
    });
  dateInput_(sh.getRange(A.date, 2, 1, 2));
  sh.getRange(A.start, 2, 2, 2).setNumberFormat('HH:mm');
  checkbox_(sh.getRange(A.check, 1));
  sh.getRange(A.check, 2, 1, 2).merge().setValue('체크하면 캘린더에 추가 (끝시간 비우면 종일)')
    .setFontSize(9).setFontColor(C.SUB);
  sh.getRange(A.result, 2).setFontColor(C.SUB).setFontSize(9).setWrap(true);

  // 도구 (메뉴 대신 체크박스)
  header(T.TOOLS_HEADER, '도구');
  [[T.SYNC_ROW, '캘린더·오늘 새로고침'], [T.SAVE_ROW, '지금 기록 저장']].forEach(function (a) {
    checkbox_(sh.getRange(a[0], 1));
    sh.getRange(a[0], 2).setValue(a[1]);
    sh.getRange(a[0], 3).setFontColor(C.SUB).setFontSize(8).setWrap(true);
  });

  // 조건부 서식: 체크하면 취소선, 습관 체크 배경, 기분 배경
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A' + T.TODO_START + '=TRUE')
      .setStrikethrough(true).setFontColor(C.SUB).setBackground(pale)
      .setRanges([sh.getRange(T.TODO_START, 1, CONFIG.TODO_ROWS, 3),
        sh.getRange(T.MUST_START, 1, CONFIG.MUST_COUNT, 3)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A' + T.HABIT_START + '=TRUE')
      .setBackground(pale).setRanges([sh.getRange(T.HABIT_START, 1, CONFIG.HABIT_COUNT, 3)]).build(),
  ].concat(moodRules_(sh, [sh.getRange(T.MOOD, 1, 2, 3)],
    'IFERROR(VALUE(LEFT($B$' + T.MOOD + ',1)),0)', theme));
  sh.setConditionalFormatRules(rules);
}

/** 오늘 시트에 적힌 날짜 */
function todaySheetDate_() {
  return dateOnly_(sheet_(CONFIG.SHEETS.TODAY).getRange(TODAY.DATE_ROW, 1).getValue());
}

/**
 * 오늘 시트 ↔ 위클리 짝 목록 (d = 오늘의 요일 번호)
 * t: 오늘 시트 [행, 열], w: 위클리 [행, 열], n: 세로 칸 수
 */
function todayBlocks_(d) {
  const T = TODAY, W = WEEKLY;
  return [
    { t: [T.TODO_START, 1], w: [W.TODO_START, W.checkCol(d)], n: CONFIG.TODO_ROWS },
    { t: [T.TODO_START, 2], w: [W.TODO_START, W.textCol(d)], n: CONFIG.TODO_ROWS },
    { t: [T.HABIT_START, 1], w: [W.HABIT_START, W.HABIT_DAY_COL + d], n: CONFIG.HABIT_COUNT },
    { t: [T.MUST_START, 1], w: [W.MUST_START, 1], n: CONFIG.MUST_COUNT },
    { t: [T.MUST_START, 2], w: [W.MUST_START, 2], n: CONFIG.MUST_COUNT },
    { t: [T.MOOD, 2], w: [W.MOOD, W.checkCol(d)], n: 1 },
    { t: [T.MOOD_MEMO, 2], w: [W.MOOD_MEMO, W.checkCol(d)], n: 1 },
  ];
}

/**
 * 짝 목록대로 값 복사 (from='t'면 오늘→위클리, 'w'면 위클리→오늘)
 * edited가 있으면 그 범위와 겹치는 묶음만 복사
 */
function copyBlocks_(fromSheet, toSheet, blocks, from, edited) {
  const to = from === 't' ? 'w' : 't';
  blocks.forEach(function (b) {
    if (edited && !hit_(edited, b[from][0], b[from][1], b.n, 1)) return;
    const vals = fromSheet.getRange(b[from][0], b[from][1], b.n, 1).getValues();
    toSheet.getRange(b[to][0], b[to][1], b.n, 1).setValues(vals);
  });
}

/** 오늘 시트를 오늘 날짜로 바꾸고 위클리 값으로 다시 채우기 */
function refreshToday_() {
  const sh = sheet_(CONFIG.SHEETS.TODAY), T = TODAY;
  const today = today_();
  sh.getRange(T.DATE_ROW, 1).setValue(today);
  const d = weeklyIndexOf_(today);
  const blocks = todayBlocks_(Math.max(d, 0));
  if (d >= 0) {
    copyBlocks_(sheet_(CONFIG.SHEETS.WEEKLY), sh, blocks, 'w');
  } else {
    // 위클리가 아직 다른 주면 비워 둠 (새 주 시작 후 다시 채워짐)
    blocks.forEach(function (b) { sh.getRange(b.t[0], b.t[1], b.n, 1).clearContent(); });
  }
}

/** 오늘 시트 수정 시: 도구 실행 / 일정 추가 / 위클리에 반영 + 기록 저장 */
function onTodayEdit_(e) {
  const T = TODAY, sh = e.range.getSheet();

  if (justChecked_(e, T.SYNC_ROW, 1)) {
    return runAction_(sh, { row: T.SYNC_ROW, col: 1 }, 3, function () {
      const n = syncCalendar_();
      ensureCurrentWeek_();
      refreshToday_();
      return '✓ 일정 ' + n + '개';
    });
  }
  if (justChecked_(e, T.SAVE_ROW, 1)) {
    return runAction_(sh, { row: T.SAVE_ROW, col: 1 }, 3, function () {
      return '저장됨 ✓ ' + saveAll_() + '행';
    });
  }
  if (justChecked_(e, T.ADD.check, 1)) {
    return addEventFromForm_(sh, {
      date: [T.ADD.date, 2], start: [T.ADD.start, 2], end: [T.ADD.end, 2],
      title: [T.ADD.title, 2], check: [T.ADD.check, 1], result: [T.ADD.result, 2],
    });
  }

  // 오늘 시트에 적힌 날짜의 위클리 요일 칸으로 반영
  const date = todaySheetDate_();
  const d = weeklyIndexOf_(date);
  if (d < 0) return;
  const blocks = todayBlocks_(d);
  const touched = blocks.some(function (b) { return hit_(e.range, b.t[0], b.t[1], b.n, 1); });
  if (!touched) return;
  copyBlocks_(sh, sheet_(CONFIG.SHEETS.WEEKLY), blocks, 't', e.range);
  archiveDayLive_(date);
  if (hit_(e.range, T.MUST_START, 1, CONFIG.MUST_COUNT, 3)) archiveWeekGoals_();
}
