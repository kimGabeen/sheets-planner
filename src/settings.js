/**
 * settings.js
 * 설정 시트: 테마색, 습관 목록, 자동 실행 상태 표시
 */

const DEFAULT_HABITS_ = ['물 8잔 마시기', '운동 30분', '독서 20분', '영어 공부', '일기 쓰기'];
const THEME_PRESETS_ = '초록 #2E7D32 · 파랑 #1565C0 · 남색 #283593 · 보라 #6A1B9A · 주황 #EF6C00 · 회색 #455A64';

/** 설정 시트 만들기 (이미 입력한 값은 유지) → 테마색 반환 */
function buildSettings_() {
  const sh = ensureSheet_(CONFIG.SHEETS.SETTINGS);
  const S = SETTINGS;
  ensureSize_(sh, S.TT_START + S.TT_COUNT + 2, 4);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 130).setColumnWidth(2, 280).setColumnWidth(3, 260).setColumnWidth(4, 200);
  sh.setRowHeights(1, 30, CONFIG.SIZE.ROW);

  // 처음 만들 때만 기본값 채우기
  const themeCell = sh.getRange(S.THEME.row, S.THEME.col);
  if (!themeCell.getValue()) themeCell.setValue(CONFIG.DEFAULT_THEME);
  const habitRange = sh.getRange(S.HABIT_START, 2, CONFIG.HABIT_COUNT, 1);
  if (habitRange.getValues().every(function (r) { return !r[0]; })) {
    habitRange.setValues(habitRange.getValues().map(function (r, i) { return [DEFAULT_HABITS_[i] || '']; }));
  }

  const cats = sh.getRange(S.CAT_START, 2, CONFIG.TIMELINE.CATEGORY_COUNT, 2);
  if (cats.getValues().every(function (r) { return !r[0]; })) {
    cats.setValues(cats.getValues().map(function (r, i) { return CONFIG.TIMELINE.DEFAULT_CATEGORIES[i] || ['', '']; }));
  }

  const theme = getTheme_();
  styleSettings_(sh, theme);
  return theme;
}

/** 설정 시트 서식 (테마 바뀔 때도 다시 실행) */
function styleSettings_(sh, theme) {
  const S = SETTINGS;
  sh.getRange(1, 1, 1, 3).merge().setValue('설정').setFontSize(18).setFontWeight('bold').setFontColor(theme);
  sh.getRange(S.THEME.row, 1).setValue('테마색').setFontWeight('bold');
  sh.getRange(S.THEME.row, 2).setBackground(theme).setFontColor(CONFIG.COLORS.WHITE)
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(S.THEME.row, 3).setValue('← #2E7D32 같은 색 코드를 넣으면 전체 색이 바뀌어요')
    .setFontColor(CONFIG.COLORS.SUB).setFontSize(9);

  // 못 한 일 자동 넘기기 (처음엔 켜짐)
  const carry = sh.getRange(S.CARRY.row, S.CARRY.col);
  if (!carry.getDataValidation()) carry.setValue(true);
  sh.getRange(S.CARRY.row, 1).setValue('못 한 일 자동 넘기기').setFontWeight('bold');
  checkbox_(carry);
  sh.getRange(S.CARRY.row, 3).setValue('← 켜면 매일 0시에 전날 미완료를 다음 날로 넘겨요')
    .setFontColor(CONFIG.COLORS.SUB).setFontSize(9);

  styleHeader_(sh.getRange(S.HABIT_HEADER, 1, 1, 3).merge()
    .setValue('습관 목록 (최대 ' + CONFIG.HABIT_COUNT + '개) — 위클리·오늘·먼슬리에 자동 반영'), theme);
  for (let i = 0; i < CONFIG.HABIT_COUNT; i++) {
    sh.getRange(S.HABIT_START + i, 1).setValue(i + 1).setHorizontalAlignment('center')
      .setFontColor(CONFIG.COLORS.SUB);
    underline_(sh.getRange(S.HABIT_START + i, 2, 1, 2).merge());
  }

  sh.getRange(S.PRESET_ROW, 1).setValue('테마 예시').setFontColor(CONFIG.COLORS.SUB);
  sh.getRange(S.PRESET_ROW, 2, 1, 2).merge().setValue(THEME_PRESETS_).setFontColor(CONFIG.COLORS.SUB)
    .setFontSize(9);

  styleHeader_(sh.getRange(S.STATUS_HEADER, 1, 1, 3).merge().setValue('자동 실행 상태 (읽기 전용)'), theme);
  const labels = { setup: '처음 설정', save: '기록 저장', sync: '캘린더 불러오기', week: '새 주 시작', carry: '못 한 일 넘기기', error: '최근 오류' };
  Object.keys(S.STATUS).forEach(function (k) {
    sh.getRange(S.STATUS[k], 1).setValue(labels[k]).setFontColor(CONFIG.COLORS.SUB);
    sh.getRange(S.STATUS[k], 2, 1, 2).merge().setFontColor(CONFIG.COLORS.SUB).setFontSize(9);
  });

  styleHeader_(sh.getRange(S.CAT_HEADER, 1, 1, 3).merge()
    .setValue('타임라인 카테고리 (웹앱) — 이름 | 색 코드'), theme);
  const n = CONFIG.TIMELINE.CATEGORY_COUNT;
  for (let i = 0; i < n; i++) {
    sh.getRange(S.CAT_START + i, 1).setValue(i + 1).setHorizontalAlignment('center').setFontColor(CONFIG.COLORS.SUB);
  }
  underline_(sh.getRange(S.CAT_START, 2, n, 2));
  // 색 코드 칸은 그 색으로 칠해서 미리보기
  getCategories_().forEach(function (c, i) {
    sh.getRange(S.CAT_START + i, 3).setBackground(c.color).setFontColor(CONFIG.COLORS.WHITE);
  });

  // 고정 시간표 (수업용): 학기 기간 + 요일 | 시작 | 끝 | 이름
  styleHeader_(sh.getRange(S.TT_HEADER, 1, 1, 4).merge()
    .setValue('고정 시간표 — 학기 기간에만 위클리·오늘 "일정"에 매주 표시'), theme);
  [[S.TT_SEM_START, '학기 시작'], [S.TT_SEM_END, '학기 종료']].forEach(function (a) {
    sh.getRange(a[0], 1).setValue(a[1]).setFontWeight('bold');
    dateInput_(sh.getRange(a[0], 2));
    sh.getRange(a[0], 3).setValue('비우면 기간 제한 없음').setFontColor(CONFIG.COLORS.SUB).setFontSize(9);
  });
  sh.getRange(S.TT_LABEL, 1, 1, 4).setValues([['요일', '시작', '끝', '이름']])
    .setFontColor(CONFIG.COLORS.SUB).setFontSize(9);
  sh.getRange(S.TT_START, 1, S.TT_COUNT, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['월', '화', '수', '목', '금', '토', '일'], true).build()).setHorizontalAlignment('center');
  sh.getRange(S.TT_START, 2, S.TT_COUNT, 2).setNumberFormat('@').setHorizontalAlignment('center');
  underline_(sh.getRange(S.TT_START, 1, S.TT_COUNT, 4));
}

const WEEKDAY_NAMES_ = ['월', '화', '수', '목', '금', '토', '일'];

/** 고정 시간표 읽기 → {start, end, rows:[{day(0=월), start, end, name}]} (시간순) */
function getTimetable_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS), S = SETTINGS;
  if (!sh || sh.getMaxColumns() < 4 || sh.getMaxRows() < S.TT_START + S.TT_COUNT - 1) return { start: null, end: null, rows: [] };
  const rows = [];
  sh.getRange(S.TT_START, 1, S.TT_COUNT, 4).getDisplayValues().forEach(function (r) {
    const day = WEEKDAY_NAMES_.indexOf(String(r[0]).trim());
    const st = normTime_(r[1]), name = String(r[3]).trim();
    if (day < 0 || !st || !name) return;
    rows.push({ day: day, start: st, end: normTime_(r[2]), name: name });
  });
  rows.sort(function (a, b) { return (a.day - b.day) || (a.start < b.start ? -1 : a.start > b.start ? 1 : 0); });
  return {
    start: dateOnly_(sh.getRange(S.TT_SEM_START, 2).getValue()),
    end: dateOnly_(sh.getRange(S.TT_SEM_END, 2).getValue()),
    rows: rows,
  };
}

/** 고정 시간표 저장 (start·end는 Date 또는 null) */
function setTimetable_(tt) {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS), S = SETTINGS;
  if (tt.rows.length > S.TT_COUNT) throw new Error('시간표는 최대 ' + S.TT_COUNT + '개예요');
  sh.getRange(S.TT_SEM_START, 2).setValue(tt.start || '');
  sh.getRange(S.TT_SEM_END, 2).setValue(tt.end || '');
  const vals = [];
  for (let i = 0; i < S.TT_COUNT; i++) {
    const r = tt.rows[i];
    vals.push(r ? [WEEKDAY_NAMES_[r.day], r.start, r.end || '', r.name] : ['', '', '', '']);
  }
  sh.getRange(S.TT_START, 1, S.TT_COUNT, 4).setValues(vals);
}

/**
 * 습관 목록 저장 (웹앱 편집): list = [{name, from(예전 순서 번호 또는 null)}]
 * 위클리 체크도 습관을 따라 함께 옮김. 기록 시트의 예전 기록은 건드리지 않음
 */
function setHabits_(list) {
  if (list.length > CONFIG.HABIT_COUNT) throw new Error('습관은 최대 ' + CONFIG.HABIT_COUNT + '개예요');
  const set = sheet_(CONFIG.SHEETS.SETTINGS), W = WEEKLY, n = CONFIG.HABIT_COUNT;
  const weekly = sheet_(CONFIG.SHEETS.WEEKLY);
  const checks = weekly.getRange(W.HABIT_START, W.HABIT_DAY_COL, n, 7).getValues();
  const names = [], moved = [];
  for (let i = 0; i < n; i++) {
    const h = list[i];
    names.push([h ? String(h.name).trim() : '']);
    const from = h && h.from !== null && h.from !== undefined ? Number(h.from) : -1;
    moved.push(from >= 0 && from < n ? checks[from].map(function (v) { return v === true; }) : [false, false, false, false, false, false, false]);
  }
  set.getRange(SETTINGS.HABIT_START, 2, n, 1).setValues(names);
  weekly.getRange(W.HABIT_START, W.HABIT_DAY_COL, n, 7).setValues(moved);
}

/** 타임라인 카테고리 목록 [{name, color}] (이름 없는 줄은 제외) */
function getCategories_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  return sh.getRange(SETTINGS.CAT_START, 2, CONFIG.TIMELINE.CATEGORY_COUNT, 2).getValues()
    .map(function (r) {
      const color = String(r[1]).trim();
      return { name: String(r[0]).trim(), color: /^#[0-9a-fA-F]{6}$/.test(color) ? color : '#C9C3B6' };
    })
    .filter(function (c) { return c.name; });
}

/** 테마색 읽기 (형식이 틀리면 기본색) */
function getTheme_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  const v = sh ? String(sh.getRange(SETTINGS.THEME.row, SETTINGS.THEME.col).getValue()).trim() : '';
  return /^#[0-9a-fA-F]{6}$/.test(v) ? v.toUpperCase() : CONFIG.DEFAULT_THEME;
}

/** 습관 이름 목록 (빈 칸은 '') */
function getHabits_() {
  const sh = sheet_(CONFIG.SHEETS.SETTINGS);
  return sh.getRange(SETTINGS.HABIT_START, 2, CONFIG.HABIT_COUNT, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
}

/** 다른 시트 수식에서 쓸 습관 이름 셀 참조 */
function habitRef_(i) {
  return "'" + CONFIG.SHEETS.SETTINGS + "'!$B$" + (SETTINGS.HABIT_START + i);
}

/** 설정 시트 수정 시: 테마색이 바뀌면 전체 다시 칠하기 */
function onSettingsEdit_(e) {
  const S = SETTINGS;
  // 고정 시간표가 바뀌면 일정 칸 다시 만들기
  if (hit_(e.range, S.TT_SEM_START, 1, S.TT_START + S.TT_COUNT - S.TT_SEM_START, 4)) {
    refreshTimetableEvents_();
    return;
  }
  // 타임라인 카테고리 색이 바뀌면 미리보기 다시 칠하기
  if (hit_(e.range, S.CAT_START, 2, CONFIG.TIMELINE.CATEGORY_COUNT, 2)) {
    const sh = e.range.getSheet();
    sh.getRange(S.CAT_START, 3, CONFIG.TIMELINE.CATEGORY_COUNT, 1).setBackground(null);
    getCategories_().forEach(function (c) {
      const i = sh.getRange(S.CAT_START, 2, CONFIG.TIMELINE.CATEGORY_COUNT, 1).getValues()
        .findIndex(function (r) { return String(r[0]).trim() === c.name; });
      if (i >= 0) sh.getRange(S.CAT_START + i, 3).setBackground(c.color).setFontColor(CONFIG.COLORS.WHITE);
    });
    return;
  }
  if (!hit_(e.range, S.THEME.row, S.THEME.col)) return;
  const v = String(e.range.getSheet().getRange(S.THEME.row, S.THEME.col).getValue()).trim();
  const note = e.range.getSheet().getRange(S.THEME.row, 3);
  if (!/^#[0-9a-fA-F]{6}$/.test(v)) {
    note.setValue('색 코드 형식이 아니에요 (예: #2E7D32)');
    return;
  }
  note.setValue('적용 중…');
  SpreadsheetApp.flush();
  buildAllLayouts_(getTheme_());
  note.setValue('적용됨 ✓ ' + stamp_());
}
