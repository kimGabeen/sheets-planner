/**
 * today.js
 * 오늘 시트: 폰 세로 화면용 (3열만 사용, 세로로 쭉)
 * 위클리의 "오늘 요일" 칸과 양방향으로 맞춰진다.
 */

/** 오늘 시트 레이아웃·수식·서식 그리기 (입력 값은 유지) */
function buildToday_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.TODAY);
  const T = TODAY, C = CONFIG.COLORS, n = CONFIG.TODO_ROWS;
  const pale = tint_(theme, CONFIG.TINT.PALE);
  const ev = CONFIG.SHEETS.EVENTS;
  ensureSize_(sh, T.LAST_ROW + 2, T.MEMO_COL);
  // 예전 레이아웃이면 비우고 새로 그림 (값은 위클리에서 다시 채워짐)
  if (needsMigration_() && sh.getLastRow() > 0) wipeSheet_(sh);
  resetLayout_(sh);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), 3).setFontSize(11).setVerticalAlignment('middle');
  sh.setColumnWidth(1, 48).setColumnWidth(2, 84).setColumnWidth(3, 210);
  sh.setRowHeights(1, T.LAST_ROW + 1, CONFIG.SIZE.TODAY_ROW);
  sh.setRowHeight(T.DATE_ROW, 48);
  sh.setFrozenRows(1);
  sh.getRange(T.TODO_HEADER, T.MEMO_COL).setValue('메모');
  sh.hideColumns(T.MEMO_COL); // 할 일 메모 (숨김)

  // 섹션 머리글 (3열 병합, count가 있으면 "제목  3 / 10")
  const header = function (row, text, count) {
    const r = sh.getRange(row, 1, 1, 3).merge();
    styleHeader_(r, theme);
    if (count) r.setFormula('="' + text + '   "&' + count);
    else r.setValue(text);
  };
  // 라벨(A) + 입력칸(B:C 병합)
  const field = function (row, label) {
    sh.getRange(row, 1).setValue(label).setFontColor(C.SUB).setFontSize(9);
    return underline_(sh.getRange(row, 2, 1, 2).merge());
  };
  const date = a1_(T.DATE_ROW, 1, true);

  // 오늘 날짜 (자정에 스크립트가 갱신)
  sh.getRange(T.DATE_ROW, 1, 1, 3).merge().setNumberFormat('m월 d일 dddd')
    .setFontSize(18).setFontWeight('bold').setFontColor(theme);

  // 오늘 일정 (자동, 읽기 전용)
  header(T.SCHED_HEADER, '오늘 일정');
  for (let k = 1; k <= CONFIG.SCHEDULE_LINES; k++) {
    sh.getRange(T.SCHED_START + k - 1, 1, 1, 3).merge().setFormula(scheduleLineFormula_(date, k, k === 1 ? '일정 없음' : ''))
      .setFontColor(C.SUB).setFontSize(10);
  }

  // 오늘 할 일 10줄: 체크 | 시간(선택) | 내용
  const tEnd = T.TODO_START + n - 1;
  header(T.TODO_HEADER, '오늘 할 일'); // 개수는 renderTodayView_가 전체 기준으로
  sh.getRange(T.TODO_MORE, 1, 1, 3).merge().setFontColor(C.SUB).setFontSize(9).setHorizontalAlignment('right');
  checkbox_(sh.getRange(T.TODO_START, 1, n, 1));
  sh.getRange(T.TODO_START, 2, n, 1).setNumberFormat('@').setHorizontalAlignment('center')
    .setFontSize(10).setFontColor(theme);
  underline_(sh.getRange(T.TODO_START, 1, n, 3));

  // 오늘 습관
  const hEnd = T.HABIT_START + CONFIG.HABIT_COUNT - 1;
  header(T.HABIT_HEADER, '오늘 습관',
    'COUNTIFS(A' + T.HABIT_START + ':A' + hEnd + ',TRUE,B' + T.HABIT_START + ':B' + hEnd + ',"?*")&" / "&COUNTIF(B' +
    T.HABIT_START + ':B' + hEnd + ',"?*")');
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
  moodInput_(field(T.MOOD, '점수'));
  field(T.MOOD_MEMO, '메모');

  // 새 일정 추가 (세로 입력 폼)
  const A = T.ADD;
  header(T.ADD_HEADER, '새 일정 추가 (구글 캘린더)');
  dateInput_(field(A.date, '날짜'));
  field(A.start, '시작').setNumberFormat('HH:mm');
  field(A.end, '끝(선택)').setNumberFormat('HH:mm');
  field(A.title, '제목');
  checkbox_(sh.getRange(A.check, 1));
  sh.getRange(A.check, 2, 1, 2).merge().setValue('체크하면 캘린더에 추가 (시작만 넣으면 종일 + 제목 앞 시간)')
    .setFontSize(9).setFontColor(C.SUB).setWrap(true);
  field(A.result, '결과').setFontColor(C.SUB).setFontSize(9).setWrap(true);

  // 도구 (메뉴 대신 체크박스, 결과는 아래 줄)
  header(T.TOOLS_HEADER, '도구');
  [[T.SYNC_ROW, T.SYNC_RESULT, '캘린더·오늘 새로고침'], [T.SAVE_ROW, T.SAVE_RESULT, '지금 기록 저장']]
    .forEach(function (a) {
      checkbox_(sh.getRange(a[0], 1));
      sh.getRange(a[0], 2, 1, 2).merge().setValue(a[2]);
      sh.getRange(a[1], 2, 1, 2).merge().setFontColor(C.SUB).setFontSize(9);
    });

  // 예전 "넘어온 일" 칸은 없앰 (전날 미완료는 전부 그날 할 일로 넘어감)
  sh.getRange(T.LAST_ROW + 1, 1, T.OLD_OVERFLOW_ROWS, 3).clearContent().setBackground(null).setBorder(false, false, false, false, false, false);

  // 조건부 서식: 체크하면 취소선, 미완 회색, 습관 체크 배경, 기분 배경
  const missed = CONFIG.CARRY.MISSED.trim();
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A' + T.TODO_START + '=TRUE')
      .setStrikethrough(true).setFontColor(C.SUB).setBackground(pale)
      .setRanges([sh.getRange(T.TODO_START, 1, n, 3),
        sh.getRange(T.MUST_START, 1, CONFIG.MUST_COUNT, 3)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($C' + T.TODO_START + ',' + missed.length + ')="' + missed + '"')
      .setFontColor(C.SUB).setRanges([sh.getRange(T.TODO_START, 1, n, 3)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A' + T.HABIT_START + '=TRUE')
      .setBackground(pale).setRanges([sh.getRange(T.HABIT_START, 1, CONFIG.HABIT_COUNT, 3)]).build(),
  ].concat(moodRules_(sh, [sh.getRange(T.MOOD, 1, 2, 3)],
    'IFERROR(VALUE(LEFT($B$' + T.MOOD + ',1)),0)', theme));
  sh.setConditionalFormatRules(rules);
  if (props_().getProperty('STORE_VERSION') === CONFIG.STORE_VERSION) renderTodayView_();
}

/** 오늘 시트에 적힌 날짜 */
function todaySheetDate_() {
  return dateOnly_(sheet_(CONFIG.SHEETS.TODAY).getRange(TODAY.DATE_ROW, 1).getValue());
}



/** 오늘 시트를 오늘 날짜로 바꾸고 위클리 값으로 다시 채우기 */
function refreshToday_() {
  sheet_(CONFIG.SHEETS.TODAY).getRange(TODAY.DATE_ROW, 1).setValue(today_());
  renderTodayView_();
}

/**
 * 저장소 → 오늘 시트 (적힌 날짜 기준)
 * 할 일 10줄·꼭 끝낼 일 5줄까지 보이고 넘치면 "+N개 더", 개수·완료율은 전체 기준
 */
function renderTodayView_() {
  const T = TODAY, sh = sheet_(CONFIG.SHEETS.TODAY), n = CONFIG.TODO_ROWS, m = CONFIG.MUST_COUNT;
  const date = todaySheetDate_();
  if (!date) return;
  const list = getTodos_(date), day = getDayData_(date), week = getWeekData_(mondayOf_(date));
  const rows = [], memos = [];
  for (let i = 0; i < n; i++) {
    const t = list[i];
    rows.push(t ? [t.done, todoTime_(t), todoDisplay_(t)] : [false, '', '']);
    memos.push([t ? t.memo : '']);
  }
  sh.getRange(T.TODO_START, 1, n, 3).setValues(rows);
  memoRange_(sh, T.TODO_START, T.MEMO_COL).setValues(memos);
  const done = list.filter(function (t) { return t.done; }).length;
  sh.getRange(T.TODO_HEADER, 1).setValue('오늘 할 일   ' + done + ' / ' + list.length);
  sh.getRange(T.TODO_MORE, 1).setValue(list.length > n ? '+' + (list.length - n) + '개 더 (웹앱에서 전부 보기)' : '');
  sh.getRange(T.HABIT_START, 1, CONFIG.HABIT_COUNT, 1)
    .setValues(getHabits_().map(function (h) { return [!!(h && day.habits[h])]; }));
  const must = [];
  for (let i = 0; i < m; i++) must.push(week.must[i] ? [week.must[i].done, week.must[i].text] : [false, '']);
  sh.getRange(T.MUST_START, 1, m, 2).setValues(must);
  sh.getRange(T.MUST_HEADER, 1).setValue('이번 주 꼭 끝낼 일' + (week.must.length > m ? '   +' + (week.must.length - m) + '개 더' : ''));
  sh.getRange(T.MOOD, 2).setValue(day.mood.score ? CONFIG.MOODS[day.mood.score - 1] : '');
  sh.getRange(T.MOOD_MEMO, 2).setValue(day.mood.memo);
}


/** 오늘 시트 수정 시: 도구 실행 / 일정 추가 / 할 일 정렬 / 위클리에 반영 + 기록 저장 */
function onTodayEdit_(e) {
  const T = TODAY, sh = e.range.getSheet();

  if (justChecked_(e, T.SYNC_ROW, 1)) {
    return runAction_(sh, { row: T.SYNC_ROW, col: 1 }, { row: T.SYNC_RESULT, col: 2 }, function () {
      const n = syncCalendar_();
      dailyRollover_();
      refreshToday_();
      return '✓ 일정 ' + n + '개';
    });
  }
  if (justChecked_(e, T.SAVE_ROW, 1)) {
    return runAction_(sh, { row: T.SAVE_ROW, col: 1 }, { row: T.SAVE_RESULT, col: 2 }, function () {
      return '저장됨 ✓ ' + saveAll_() + '행';
    });
  }
  if (justChecked_(e, T.ADD.check, 1)) {
    return addEventFromForm_(sh, {
      date: [T.ADD.date, 2], start: [T.ADD.start, 2], end: [T.ADD.end, 2],
      title: [T.ADD.title, 2], check: [T.ADD.check, 1], result: [T.ADD.result, 2],
    });
  }

  // 오늘 시트에 적힌 날짜의 저장소로 반영 → 위클리 시트도 다시 그림
  const date = todaySheetDate_();
  if (!date) return;
  let changed = false;
  if (hit_(e.range, T.TODO_START, 1, CONFIG.TODO_ROWS, T.MEMO_COL)) {
    syncTodosFromSheet_(sh, T.TODO_START, 1, T.MEMO_COL, date);
    changed = true;
  }
  if (hit_(e.range, T.HABIT_START, 1, CONFIG.HABIT_COUNT, 1)) {
    const checks = sh.getRange(T.HABIT_START, 1, CONFIG.HABIT_COUNT, 1).getValues();
    const map = {};
    getHabits_().forEach(function (n, i) { if (n) map[n] = checks[i][0] === true; });
    setHabitsDone_(date, map);
    changed = true;
  }
  if (hit_(e.range, T.MOOD, 2, 2, 2)) {
    setMoodData_(date, moodScore_(sh.getRange(T.MOOD, 2).getValue()), sh.getRange(T.MOOD_MEMO, 2).getValue());
    changed = true;
  }
  if (hit_(e.range, T.MUST_START, 1, CONFIG.MUST_COUNT, 3)) {
    syncMustFromSheet_(sh, mondayOf_(date), T.MUST_START, 1, 2);
    archiveWeekGoals_(mondayOf_(date));
    changed = true;
  }
  if (!changed) return;
  archiveDayLive_(date);
  renderTodayView_();
  const d = weeklyIndexOf_(date);
  if (d >= 0) renderWeekView_({ days: [d], habits: true, must: true });
}
