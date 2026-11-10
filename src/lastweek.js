/**
 * lastweek.js
 * 지난 주 보기: 주차 드롭다운을 고르면 기록 시트에서 그 주 결과를
 * FILTER 수식으로 불러와 위클리와 비슷한 모양으로 보여줌 (읽기 전용)
 */

/** 지난 주 보기 시트 그리기 */
function buildLastWeek_(theme) {
  const sh = ensureSheet_(CONFIG.SHEETS.LASTWEEK);
  const L = LASTWEEK, W = WEEKLY, C = CONFIG.COLORS, Z = CONFIG.SIZE, CAT = CONFIG.CATEGORY;
  const arc = CONFIG.SHEETS.ARCHIVE;
  const A = colRef_(arc, 1), B = colRef_(arc, 2), I = colRef_(arc, 3), D = colRef_(arc, 4),
    E = colRef_(arc, 5), F = colRef_(arc, 6);
  const pale = tint_(theme, CONFIG.TINT.PALE);
  ensureSize_(sh, L.MOOD_MEMO + 5, L.LAST_COL + 1);
  resetLayout_(sh);
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(10).setVerticalAlignment('middle');

  // 열 너비 (위클리와 같게)
  sh.setColumnWidth(1, Z.LABEL_COL);
  for (let d = 0; d < 7; d++) {
    sh.setColumnWidth(W.checkCol(d), Z.CHECK_COL).setColumnWidth(W.textCol(d), Z.TEXT_COL);
  }
  sh.setColumnWidth(L.HABIT_NAME_COL - 1, Z.GAP_COL).setColumnWidth(L.HABIT_NAME_COL, Z.HABIT_NAME_COL);
  sh.setColumnWidths(L.HABIT_DAY_COL, 7, Z.HABIT_DAY_COL);
  sh.setColumnWidth(L.HABIT_PCT_COL, 50);
  sh.setRowHeights(1, L.MOOD_MEMO + 1, Z.ROW);

  // 숨김 도우미: 주차 목록(최신순), 선택 주 월요일
  const sel = a1_(L.SELECT.row, L.SELECT.col, true);
  const mon = a1_(L.MONDAY.row, L.MONDAY.col, true);
  const key = 'TEXT(' + mon + ',"yyyy-mm-dd")';
  sh.getRange(1, L.HELPER_COL).setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + F + ',' + F + '<>"")),1,FALSE),"")');
  sh.getRange(L.MONDAY.row, L.MONDAY.col).setFormula(
    '=IF(' + sel + '="","",IF(ISNUMBER(' + sel + '),' + sel + ',DATEVALUE(' + sel + ')))');
  sh.hideColumns(L.HELPER_COL, 2);

  // 상단: 제목 + 주차 드롭다운
  sh.getRange(1, 1, 1, 2).merge().setValue('지난 주 보기').setFontSize(16).setFontWeight('bold').setFontColor(theme);
  sh.getRange(1, 3).setValue('주차 (월요일)').setFontColor(C.SUB).setHorizontalAlignment('right');
  sh.getRange(L.SELECT.row, L.SELECT.col, 1, 3).merge().setNumberFormat('@')
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(sh.getRange(1, L.HELPER_COL, 300, 1), true).build())
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, 7, 1, 9).merge().setValue('기록 시트에서 자동으로 불러와요 (읽기 전용)')
    .setFontColor(C.SUB).setFontSize(9);

  // 그 주 꼭 끝낼 일
  styleHeader_(sh.getRange(L.MUST_HEADER, 1, 1, 6).merge().setValue('그 주 꼭 끝낼 일'), theme);
  sh.getRange(L.MUST_START, 1).setFormula('=IFERROR(ARRAY_CONSTRAIN(ARRAYFORMULA(FILTER({IF(' + D +
    ',"✓","·"),' + I + '},' + B + '="' + CAT.WEEK_GOAL + '",' + F + '=' + key + ')),' + CONFIG.MUST_COUNT + ',2),"")');
  sh.getRange(L.MUST_START, 1, CONFIG.MUST_COUNT, 1).setHorizontalAlignment('center').setFontColor(theme);

  // 요일 블록
  [[L.RATE, '완료율'], [L.TODO_START, '할 일'], [L.COUNT, '합계'], [L.MOOD, '기분'], [L.MOOD_MEMO, '메모']]
    .forEach(function (a) { sh.getRange(a[0], 1).setValue(a[1]).setFontColor(C.SUB).setFontSize(9); });
  const emoji = CONFIG.MOOD_EMOJI.map(function (x) { return '"' + x + '"'; }).join(',');
  for (let d = 0; d < 7; d++) {
    const c = W.checkCol(d), t = W.textCol(d);
    const date = colLetter_(c) + '$' + L.DAY_HEADER;
    const mk = a1_(L.TODO_START, c) + ':' + a1_(L.COUNT - 1, c);
    const tx = a1_(L.TODO_START, t) + ':' + a1_(L.COUNT - 1, t);
    const done = 'COUNTIF(' + mk + ',"✓")', tot = 'COUNTIF(' + tx + ',"?*")';

    styleHeader_(sh.getRange(L.DAY_HEADER, c, 1, 2).merge()
      .setFormula('=IF(' + mon + '="","",' + mon + '+' + d + ')').setNumberFormat('m/d (ddd)')
      .setHorizontalAlignment('center'), theme);
    sh.getRange(L.RATE, c, 1, 2).merge().setFormula('=LET(done,' + done + ',tot,' + tot +
      ',p,IF(tot=0,0,done/tot),IF(tot=0,"",CHOOSE(ROUND(p*4)+1,"○","◔","◑","◕","●")&"  "&TEXT(p,"0%")))')
      .setHorizontalAlignment('center').setFontSize(13).setFontColor(theme).setFontWeight('bold');
    // 할 일: ✓/· 와 내용을 2열로 펼침
    sh.getRange(L.TODO_START, c).setFormula('=IFERROR(ARRAY_CONSTRAIN(ARRAYFORMULA(FILTER({IF(' + D +
      ',"✓","·"),' + I + '},' + A + '=' + date + ',' + B + '="' + CAT.TODO + '")),' + CONFIG.TODO_ROWS + ',2),"")');
    sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, 1).setHorizontalAlignment('center').setFontColor(theme);
    underline_(sh.getRange(L.TODO_START, c, CONFIG.TODO_ROWS, 2));
    sh.getRange(L.COUNT, c, 1, 2).merge().setFormula('=IF(' + tot + '=0,"",' + done + '&" 완료 · "&(' + tot +
      '-' + done + ')&" 남음")').setHorizontalAlignment('center').setFontColor(C.SUB).setFontSize(9);
    const s = 'IFERROR(INDEX(FILTER(' + E + ',' + A + '=' + date + ',' + B + '="' + CAT.MOOD + '"),1),"")';
    sh.getRange(L.MOOD, c, 1, 2).merge().setFormula('=LET(s,' + s + ',IF(s="","",s&"점 "&CHOOSE(s,' + emoji + ')))')
      .setHorizontalAlignment('center');
    sh.getRange(L.MOOD_MEMO, c, 1, 2).merge().setFormula('=IFERROR(INDEX(FILTER(' + I + ',' + A + '=' + date +
      ',' + B + '="' + CAT.MOOD + '"),1),"")').setFontSize(9);
  }

  // 습관 결과
  styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_NAME_COL).setValue('습관'), theme);
  for (let d = 0; d < 7; d++) {
    styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_DAY_COL + d).setFormula('=' + a1_(L.DAY_HEADER, W.checkCol(d)))
      .setNumberFormat('ddd').setHorizontalAlignment('center'), theme);
  }
  styleHeader_(sh.getRange(L.DAY_HEADER, L.HABIT_PCT_COL).setValue('%').setHorizontalAlignment('center'), theme);
  sh.getRange(L.HABIT_START, L.HABIT_NAME_COL).setFormula('=IFERROR(ARRAY_CONSTRAIN(UNIQUE(FILTER(' + I + ',' +
    B + '="' + CAT.HABIT + '",' + F + '=' + key + ')),' + CONFIG.HABIT_COUNT + ',1),"")');
  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    const r = L.HABIT_START + i, name = '$' + colLetter_(L.HABIT_NAME_COL) + r;
    const f = [];
    for (let d = 0; d < 7; d++) {
      const day = colLetter_(L.HABIT_DAY_COL + d) + '$' + L.DAY_HEADER;
      f.push('=IF(' + name + '="","",IF(COUNTIFS(' + A + ',' + day + ',' + B + ',"' + CAT.HABIT + '",' + I + ',' +
        name + ',' + D + ',TRUE)>0,"●","·"))');
    }
    sh.getRange(r, L.HABIT_DAY_COL, 1, 7).setFormulas([f]).setHorizontalAlignment('center').setFontColor(theme);
    sh.getRange(r, L.HABIT_PCT_COL).setFormula('=IF(' + name + '="","",COUNTIF(' + a1_(r, L.HABIT_DAY_COL) + ':' +
      a1_(r, L.HABIT_DAY_COL + 6) + ',"●")/7)').setNumberFormat('0%').setFontColor(C.SUB).setFontSize(9);
    underline_(sh.getRange(r, L.HABIT_NAME_COL, 1, 9));
  }

  // 조건부 서식: 완료 항목 취소선, 기분 배경
  const first = W.checkCol(0), last = W.textCol(6);
  const tl = a1_(L.TODO_START, first), left = a1_(L.TODO_START, first - 1);
  const idx = 'COLUMN()-1-N(ISODD(COLUMN()))';
  const moodRow = '$' + colLetter_(first) + '$' + L.MOOD + ':$' + colLetter_(last) + '$' + L.MOOD;
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR(AND(ISEVEN(COLUMN()),' + tl + '="✓"),AND(ISODD(COLUMN()),' + left + '="✓"))')
      .setStrikethrough(true).setFontColor(C.SUB).setBackground(pale)
      .setRanges([sh.getRange(L.TODO_START, first, CONFIG.TODO_ROWS, last - first + 1)]).build(),
  ].concat(moodRules_(sh, [sh.getRange(L.MOOD, first, 2, last - first + 1)],
    'IFERROR(VALUE(LEFT(INDEX(' + moodRow + ',1,' + idx + '),1)),0)', theme)));
}

/** 드롭다운이 비어 있으면 가장 최근 지난 주로 채우기 */
function initLastWeekSelect_() {
  const sh = sheet_(CONFIG.SHEETS.LASTWEEK), L = LASTWEEK;
  const cell = sh.getRange(L.SELECT.row, L.SELECT.col);
  if (cell.getValue()) return;
  const cur = weekKey_(today_());
  const weeks = readArchive_().map(function (r) { return r.week; })
    .filter(function (w) { return w && w < cur; }).sort();
  if (weeks.length) cell.setValue(weeks[weeks.length - 1]);
}
